<#
.SYNOPSIS
    Step 26 - Analysis of IIS logs and SMTP protocol logs (2013 vs 2019 distribution).

.DESCRIPTION
    Analyzes the IIS logs and the SMTP Protocol Logs (Receive + Send) to check
    the 2013 vs 2019 distribution.

    PURPOSE: confirm after the cutover that clients only hit the 2019 servers
    and that there are no residual connections left on the 2013 servers.

    Log SOURCES analyzed:
      - IIS Default Web Site (Front End)         : HTTP/HTTPS client traffic
        => MAPI, OWA, EWS, EAS (Microsoft-Server-ActiveSync), OAB, ECP, RPC,
           Autodiscover, PowerShell.
        Path: retrieved DYNAMICALLY on the server side via
        (Get-Website -Name 'Default Web Site').logFile.directory (environment
        variables are expanded). Step06 may have relocated this path (LogPaths)
        on 2019; on 2013 it is usually the system default.
      - SMTP Receive Protocol Log                : inbound SMTP traffic
      - SMTP Send Protocol Log                   : outbound SMTP traffic
        Path via Get-TransportService + Get-FrontendTransportService.

    NOT included: MessageTracking (mail transactions), Exchange Back End HTTP
    Logs (internal 2013<->2019 proxy, not driven by clients).

    Modes:
      - Inventory : full scan (paths, ProtocolLoggingLevel, stats) without any
                    change. Reports the SMTP connectors with
                    ProtocolLoggingLevel='None' (invisible logs).
      - Simulate  : -WhatIf on Set-ReceiveConnector / Set-SendConnector.
      - Apply     : sets ProtocolLoggingLevel='Verbose' on the connectors where
                    it is disabled (according to
                    Config.LogAnalysis.SmtpAutoEnableProtocolLogs) AND produces
                    the stats report. The first run after enabling will have
                    little or no SMTP data - run it again after N hours.

    Environment variables:
      LOG_HOURS=<n>          : overrides the sliding window (default from config).

    CSV report: one row per scanned target (server + source). The global HTML
    shows a summary table per Exchange version.

    Exchange/PowerShell pitfalls handled:
      - $cnx.Server can be a TransportServer object (.Name) or a string
        depending on the version. Systematic [string] cast.
      - %SystemRoot%/%SystemDrive% resolution in logFile.directory via
        [Environment]::ExpandEnvironmentVariables() on the remote side.
      - Get-WebSite over WinRM requires the WebAdministration module on the target.
      - IIS W3C logs: the #Fields header varies by version (it can be re-emitted
        in the middle of a file). It is parsed on every #Fields line encountered.
      - SMTP logs: CSV format with a #Fields line, RFC 5321 ASCII events.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    Part of : Exchange 2013/2016 to 2019 Migration - Deploy-Exchange2019.ps1
#>

function Test-ProtocolLogExcludedIdentity {
    <#
        Returns true for a technical identity or an empty protocol-log identity.
        The same helper is used for IIS cs-username and SMTP MAIL FROM values so
        the report cannot count a system object through a different log format.
    #>
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Identity,
        [string[]]$Patterns = @()
    )

    $value = if ($null -eq $Identity) { '' } else { $Identity.Trim().Trim('"') }
    if ([string]::IsNullOrWhiteSpace($value) -or $value -in @('-', '<>', 'ANONYMOUS', 'ANONYMOUS LOGON')) {
        return $true
    }

    $candidates = [Collections.Generic.List[string]]::new()
    [void]$candidates.Add($value)
    if ($value.Contains('\')) { [void]$candidates.Add($value.Substring($value.LastIndexOf('\') + 1)) }
    if ($value -match '^(?:<)?([^>@\s]+)(?:@[^>\s]+)?>?$') { [void]$candidates.Add($Matches[1]) }

    foreach ($candidate in $candidates) {
        foreach ($pattern in $Patterns) {
            if ($candidate -match $pattern) { return $true }
        }
    }
    return $false
}

function Invoke-Step {
    [CmdletBinding()]
    param(
        [string]$Mode,
        [string]$OutputFolder,
        [string]$CsvFolder,
        [hashtable]$Config,
        [System.Management.Automation.PSCredential]$Credential
    )

    # --- 1) Analysis parameters --------------------------------------------
    $logCfg = $Config.LogAnalysis
    $hours = if ($env:LOG_HOURS -match '^\d+$') { [int]$env:LOG_HOURS }
             elseif ($logCfg -and $logCfg.DefaultHoursWindow) { [int]$logCfg.DefaultHoursWindow }
             else { 24 }
    $since = (Get-Date).AddHours(-$hours)

    $vDirs = if ($logCfg -and $logCfg.IISVdirsToTrack) { @($logCfg.IISVdirsToTrack) }
             else { @('Mapi','OWA','EWS','OAB','ECP','RPC','Autodiscover','Microsoft-Server-ActiveSync','PowerShell') }
    # Lowercase index for fast matching
    $vDirsLower = @{}
    foreach ($v in $vDirs) { $vDirsLower[$v.ToLowerInvariant()] = $v }

    $autoEnable = if ($logCfg -and $null -ne $logCfg.SmtpAutoEnableProtocolLogs) {
        [bool]$logCfg.SmtpAutoEnableProtocolLogs
    } else { $true }

    # Filters: monitoring probes / technical mailboxes / SMTP probes
    $excludeUri  = @(); $excludeUser = @(); $excludeUA = @()
    $excludeSmtpSender = @()
    if ($logCfg) {
        if ($logCfg.IISExcludeUriPatterns)       { $excludeUri  = @($logCfg.IISExcludeUriPatterns) }
        if ($logCfg.IISExcludeUserPatterns)      { $excludeUser = @($logCfg.IISExcludeUserPatterns) }
        if ($logCfg.IISExcludeUserAgentPatterns) { $excludeUA   = @($logCfg.IISExcludeUserAgentPatterns) }
        if ($logCfg.SmtpExcludeSenderPatterns)  { $excludeSmtpSender = @($logCfg.SmtpExcludeSenderPatterns) }
    }
    $requireAuthenticatedUser = if ($logCfg -and $null -ne $logCfg.IISRequireAuthenticatedUser) {
        [bool]$logCfg.IISRequireAuthenticatedUser
    } else { $true }
    $smtpRequireMailFrom = if ($logCfg -and $null -ne $logCfg.SmtpRequireMailFromForValid) {
        [bool]$logCfg.SmtpRequireMailFromForValid
    } else { $true }
    $smtpRequireSenderAddress = if ($logCfg -and $null -ne $logCfg.SmtpRequireSenderAddress) {
        [bool]$logCfg.SmtpRequireSenderAddress
    } else { $true }
    # IISRequireSuccessStatus: only count the 2xx sc-status values (successful auth).
    # Without this filter, the 401 challenges of the Negotiate/NTLM round-trip are counted
    # as user traffic (1 real hit -> 2 IIS lines: 401 then 200).
    # Worse: the HealthMailbox/AMProbe probes send a 401 without cs-username
    # which slips through the user/UA filters and skews the metric.
    $requireSuccess = if ($logCfg -and $null -ne $logCfg.IISRequireSuccessStatus) {
        [bool]$logCfg.IISRequireSuccessStatus
    } else { $true }

    Write-StepBanner -StepName '26' -Title 'IIS + SMTP log analysis - 2013 vs 2019 distribution' -Mode $Mode -Actions @(
        "Analysis period:     $hours rolling hour(s) (since $($since.ToString('yyyy-MM-dd HH:mm')))",
        "IIS vDirs tracked:   $($vDirs -join ', ')",
        "IIS excluded filters: $($excludeUri.Count) URI / $($excludeUser.Count) identities / $($excludeUA.Count) UserAgent",
        ("IIS auth. identity required: {0} (anonymous requests are excluded)" -f $requireAuthenticatedUser),
        ("IIS sc-status 2xx required: {0} (excludes 401 challenges + errors)" -f $requireSuccess),
        "SMTP scope:          Receive + Send Protocol Logs (NOT message tracking)",
        ("SMTP valid sessions = MAIL FROM observed: {0}" -f $smtpRequireMailFrom),
        ("SMTP user sender required: {0} (system filters: {1})" -f $smtpRequireSenderAddress, $excludeSmtpSender.Count),
        "Apply SMTP:          auto-enable Verbose if disabled = $autoEnable",
        "Inventory: read-only scan. Apply: enable disabled SMTP logs + stats."
    )

    Reset-Report
    Initialize-ExchangeShell -Credential $Credential

    # --- 2) Exchange server inventory (2013 + 2019, Edge excluded) ---------
    $allServers = @(Get-ExchangeServer -ErrorAction Stop |
                    Where-Object { $_.ServerRole -notlike '*Edge*' })
    if (-not $allServers) {
        Write-Log "No Exchange server found." -Level Warning
        Save-Report -StepName 'Step26-AnalyzeProtocolLogs' | Out-Null
        return 0
    }

    $serverVersion = @{}
    foreach ($s in $allServers) {
        $v = if     ($s.AdminDisplayVersion -like 'Version 15.2*') { '2019' }
             elseif ($s.AdminDisplayVersion -like 'Version 15.1*') { '2016' }
             elseif ($s.AdminDisplayVersion -like 'Version 15.0*') { '2013' }
             else { 'Other' }
        $serverVersion[[string]$s.Name] = $v
    }
    Write-Log ('=' * 70) -Level Step
    Write-Log "Exchange servers detected:" -Level Step
    foreach ($name in ($serverVersion.Keys | Sort-Object)) {
        Write-Log ("  {0,-12} -> {1}" -f $name, $serverVersion[$name]) -Level Info
    }
    Write-Log ('=' * 70) -Level Step

    # ========================================================================
    #  PART 1: SMTP Protocol Logs
    # ========================================================================
    Write-Log '===== SMTP Protocol Logs =====' -Level Step

    # 1.1 ProtocolLoggingLevel inventory per connector
    $receiveCnx = @(Get-ReceiveConnector -ErrorAction SilentlyContinue)
    $sendCnx    = @(Get-SendConnector -ErrorAction SilentlyContinue)
    Write-Log ("ReceiveConnectors detected: {0} | SendConnectors: {1}" -f $receiveCnx.Count, $sendCnx.Count) -Level Info

    # Helper: extracts the server name from an Identity 'Server\Connector' or a TransportServer object
    function Get-ServerNameFromConnector {
        param($Connector)
        if ($Connector.Server) {
            $sv = $Connector.Server
            if ($sv.Name) { return [string]$sv.Name }
            return ([string]$sv -split '\\')[0]
        }
        # Identity 'EXCH201901\Default EXCH201901'
        return ([string]$Connector.Identity -split '\\')[0]
    }

    # Receive connectors
    foreach ($cnx in $receiveCnx) {
        $srvName = Get-ServerNameFromConnector -Connector $cnx
        $cnxName = [string]$cnx.Name
        $level   = [string]$cnx.ProtocolLoggingLevel
        $ver     = if ($serverVersion.ContainsKey($srvName)) { $serverVersion[$srvName] } else { '?' }

        Add-Report -Step '26-LogAnalysis' -Target ("$srvName\$cnxName") -Action 'ReceiveConnector inventory' -Status Inventoried `
            -AfterValue ("Server=$srvName ($ver) | ProtocolLoggingLevel=$level")

        if ($level -eq 'None' -and ($Mode -eq 'Apply' -or $Mode -eq 'Simulate') -and $autoEnable) {
            Invoke-Action -Step '26-LogAnalysis' -Target ("$srvName\$cnxName") -Action 'Enable Receive ProtocolLog' `
                -Detail "Server=$srvName ($ver); before=None, target=Verbose" `
                -InventoryScript {
                    $cur = Get-ReceiveConnector -Identity $cnx.Identity -ErrorAction SilentlyContinue
                    if ($cur) { "ProtocolLoggingLevel=$([string]$cur.ProtocolLoggingLevel)" } else { '<absent>' }
                } `
                -PreCheckScript {
                    $cur = Get-ReceiveConnector -Identity $cnx.Identity -ErrorAction SilentlyContinue
                    return ($cur -and [string]$cur.ProtocolLoggingLevel -eq 'Verbose')
                } `
                -ActionScript {
                    Set-ReceiveConnector -Identity $cnx.Identity -ProtocolLoggingLevel Verbose -Confirm:$false -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                }
        }
    }

    # Send connectors
    foreach ($cnx in $sendCnx) {
        $cnxName = [string]$cnx.Name
        $level   = [string]$cnx.ProtocolLoggingLevel

        Add-Report -Step '26-LogAnalysis' -Target $cnxName -Action 'SendConnector inventory' -Status Inventoried `
            -AfterValue ("ProtocolLoggingLevel=$level | SourceTransportServers=" + (@($cnx.SourceTransportServers | ForEach-Object { [string]$_.Name }) -join ','))

        if ($level -eq 'None' -and ($Mode -eq 'Apply' -or $Mode -eq 'Simulate') -and $autoEnable) {
            Invoke-Action -Step '26-LogAnalysis' -Target $cnxName -Action 'Enable Send ProtocolLog' `
                -Detail "Connector=$cnxName; before=None, target=Verbose" `
                -InventoryScript {
                    $cur = Get-SendConnector -Identity $cnx.Identity -ErrorAction SilentlyContinue
                    if ($cur) { "ProtocolLoggingLevel=$([string]$cur.ProtocolLoggingLevel)" } else { '<absent>' }
                } `
                -PreCheckScript {
                    $cur = Get-SendConnector -Identity $cnx.Identity -ErrorAction SilentlyContinue
                    return ($cur -and [string]$cur.ProtocolLoggingLevel -eq 'Verbose')
                } `
                -ActionScript {
                    Set-SendConnector -Identity $cnx.Identity -ProtocolLoggingLevel Verbose -Confirm:$false -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                }
        }
    }

    # 1.2 Retrieve the SMTP paths per server (Hub + Frontend)
    Write-Log '----- Retrieving SMTP paths per server -----' -Level Sub
    $smtpPaths = @()
    foreach ($srv in $allServers) {
        $srvName = [string]$srv.Name
        $ver     = $serverVersion[$srvName]
        try {
            $ts = Get-TransportService -Identity $srvName -ErrorAction Stop
            if ($ts.ReceiveProtocolLogPath) { $smtpPaths += [PSCustomObject]@{ Server=$srvName; Version=$ver; Kind='HubReceive'; Path=[string]$ts.ReceiveProtocolLogPath } }
            if ($ts.SendProtocolLogPath)    { $smtpPaths += [PSCustomObject]@{ Server=$srvName; Version=$ver; Kind='HubSend';    Path=[string]$ts.SendProtocolLogPath } }
        } catch {
            Write-Log "  Get-TransportService $srvName - $($_.Exception.Message)" -Level Warning
        }
        try {
            $fe = Get-FrontendTransportService -Identity $srvName -ErrorAction Stop
            if ($fe.ReceiveProtocolLogPath) { $smtpPaths += [PSCustomObject]@{ Server=$srvName; Version=$ver; Kind='FERecv'; Path=[string]$fe.ReceiveProtocolLogPath } }
            if ($fe.SendProtocolLogPath)    { $smtpPaths += [PSCustomObject]@{ Server=$srvName; Version=$ver; Kind='FESend'; Path=[string]$fe.SendProtocolLogPath } }
        } catch {
            # Not all servers have FrontendTransport (Edge is excluded anyway)
        }
    }
    foreach ($p in $smtpPaths) {
        Write-Log ("  {0,-12} ({1}) {2,-10} -> {3}" -f $p.Server, $p.Version, $p.Kind, $p.Path) -Level Info
    }

    # 1.3 Parse the SMTP logs filtered by window
    Write-Log '----- Parsing SMTP logs -----' -Level Sub
    # Aggregates per server+kind: unique sessions, top remote IPs
    $smtpStats = @()
    foreach ($p in $smtpPaths) {
        $unc = Convert-LocalPathToUnc -ComputerName $p.Server -LocalPath $p.Path
        if (-not (Test-Path $unc -ErrorAction SilentlyContinue)) {
            Write-Log "  [$($p.Server)/$($p.Kind)] Path not accessible: $unc" -Level Warning
            $smtpStats += [PSCustomObject]@{ Server=$p.Server; Version=$p.Version; Kind=$p.Kind; Path=$p.Path; Sessions=0; UniqueRemoteIps=0; TopRemoteIps=''; Note='Path not accessible' }
            continue
        }
        $logFiles = @(Get-ChildItem -Path $unc -Filter '*.log' -File -ErrorAction SilentlyContinue |
                      Where-Object { $_.LastWriteTime -ge $since })
        if (-not $logFiles) {
            $smtpStats += [PSCustomObject]@{ Server=$p.Server; Version=$p.Version; Kind=$p.Kind; Path=$p.Path; Sessions=0; UniqueRemoteIps=0; TopRemoteIps=''; Note='No log in window' }
            continue
        }
        # Per-session tracking: mark a session as "valid" if it has seen MAIL FROM:.
        # A Managed Availability probe usually only does EHLO/NOOP/QUIT
        # without ever reaching MAIL FROM, so they are excluded this way.
        $sessions       = @{}        # sessionId -> $true (all sessions seen)
        $validSessions  = @{}        # sessionId -> $true (MAIL FROM observed)
        $sessionRemote  = @{}        # sessionId -> remoteIp (so that IPs only count for valid sessions)
        $sessionConn    = @{}        # sessionId -> connector-id (for per-connector drill-down)
        $mailFromSessions = @{}      # sessionId -> a MAIL FROM was observed
        $systemSessions = @{}        # sessionId -> MAIL FROM belongs to a system object
        $remoteIpsAll   = @{}        # IP -> count, all sessions
        $remoteIpsValid = @{}        # IP -> count, valid sessions
        $connectorsAll  = @{}        # connectorId -> count, all sessions
        $connectorsValid= @{}        # connectorId -> count, valid sessions
        foreach ($lf in $logFiles) {
            try {
                Get-Content -LiteralPath $lf.FullName -ErrorAction Stop | ForEach-Object {
                    $line = $_
                    if ($line.StartsWith('#') -or [string]::IsNullOrWhiteSpace($line)) { return }
                    # Format: date-time,connector-id,session-id,sequence-number,local-endpoint,remote-endpoint,event,data,context
                    $parts = $line -split ','
                    if ($parts.Count -lt 8) { return }
                    $dt = $null
                    try { $dt = [DateTime]::Parse($parts[0]) } catch { return }
                    if ($dt -lt $since) { return }
                    $connId    = $parts[1]
                    $sessionId = $parts[2]
                    $remoteEp  = $parts[5]
                    $ev        = $parts[6]
                    $data      = $parts[7]
                    if ($sessionId) { $sessions[$sessionId] = $true }
                    if ($sessionId -and $connId -and -not $sessionConn.ContainsKey($sessionId)) {
                        $sessionConn[$sessionId] = $connId
                    }
                    if ($remoteEp -and -not $sessionRemote.ContainsKey($sessionId)) {
                        $sessionRemote[$sessionId] = ($remoteEp -split ':')[0]
                    }
                    if ($remoteEp) {
                        $ip = ($remoteEp -split ':')[0]
                        if ($ip) { $remoteIpsAll[$ip] = ($remoteIpsAll[$ip] + 1) }
                    }
                    # A valid user session must submit a non-empty MAIL FROM.
                    # System mailboxes and monitoring senders are excluded even
                    # when the probe completes a full SMTP transaction.
                    if ($sessionId -and $ev -eq '<' -and $data -and $data -match '^\s*MAIL FROM:') {
                        $mailFromSessions[$sessionId] = $true
                        $mailFrom = ''
                        if ($data -match '^\s*MAIL FROM:\s*<?([^>\s]+)>?') { $mailFrom = $Matches[1].Trim() }
                        if ($smtpRequireSenderAddress -and
                            (Test-ProtocolLogExcludedIdentity -Identity $mailFrom -Patterns $excludeSmtpSender)) {
                            $systemSessions[$sessionId] = $true
                        }
                    }
                }
            } catch {
                Write-Log "  [$($p.Server)/$($p.Kind)] Parsing error in $($lf.Name): $($_.Exception.Message)" -Level Warning
            }
        }
        # Derive valid sessions only after all MAIL FROM commands are known.
        foreach ($sid in $mailFromSessions.Keys) {
            if (-not $systemSessions.ContainsKey($sid)) { $validSessions[$sid] = $true }
        }
        # Re-aggregate remoteIps + connectors by valid / total sessions
        foreach ($sid in $sessions.Keys) {
            $c = $sessionConn[$sid]; if ($c) { $connectorsAll[$c] = ($connectorsAll[$c] + 1) }
        }
        foreach ($sid in $validSessions.Keys) {
            $ip = $sessionRemote[$sid]
            if ($ip) { $remoteIpsValid[$ip] = ($remoteIpsValid[$ip] + 1) }
            $c = $sessionConn[$sid]
            if ($c) { $connectorsValid[$c] = ($connectorsValid[$c] + 1) }
        }
        # System sessions are never reportable. The switch only controls whether
        # sessions without MAIL FROM are visible; strict mode remains the default.
        $eligibleSessions = @{}
        foreach ($sid in $sessions.Keys) {
            if (-not $systemSessions.ContainsKey($sid)) { $eligibleSessions[$sid] = $true }
        }
        $sessionsToReport     = if ($smtpRequireMailFrom) { $validSessions } else { $eligibleSessions }
        $ipsToReport          = if ($smtpRequireMailFrom) { $remoteIpsValid } else { $remoteIpsAll }
        $connectorsToReport   = if ($smtpRequireMailFrom) { $connectorsValid } else { $connectorsAll }
        $top = ($ipsToReport.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 5 |
                ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ' ; '
        $smtpStats += [PSCustomObject]@{
            Server=$p.Server; Version=$p.Version; Kind=$p.Kind; Path=$p.Path
            Sessions=$sessionsToReport.Count; SessionsTotal=$sessions.Count
            ExcludedSystem=$systemSessions.Count; ExcludedNoMailFrom=($sessions.Count - $mailFromSessions.Count)
            UniqueRemoteIps=$ipsToReport.Count; TopRemoteIps=$top
            ByConnector=$connectorsToReport; Note=''
        }
    }

    foreach ($s in $smtpStats) {
        $filteredOut = [Math]::Max(0, $s.SessionsTotal - $s.Sessions)
        $detail = "ValidSessions={0} | SessionsTotal={1} | SystemObjects={2} | NoMailFrom={3} | UniqueRemoteIps={4} | TopIps={5}" `
                  -f $s.Sessions, $s.SessionsTotal, $s.ExcludedSystem, $s.ExcludedNoMailFrom, $s.UniqueRemoteIps, $s.TopRemoteIps
        if ($s.Note) { $detail = "$detail | Note=$($s.Note)" }
        Add-Report -Step '26-LogAnalysis' -Target ("$($s.Server)\$($s.Kind)") -Action ("SMTP stats {0}" -f $s.Kind) -Status Inventoried `
            -AfterValue ("Version=$($s.Version) | $detail")
        Write-Log ("  {0,-12} ({1}) {2,-10}: Valid={3,5} (Total={4,5}, Non-user={5,5}) | IPs={6,4} | Top: {7}" `
            -f $s.Server, $s.Version, $s.Kind, $s.Sessions, $s.SessionsTotal, $filteredOut, $s.UniqueRemoteIps, $s.TopRemoteIps) -Level Info
    }

    # ========================================================================
    #  PART 2: IIS Logs (Default Web Site - Front End)
    # ========================================================================
    Write-Log '===== IIS Logs (Default Web Site / Front End) =====' -Level Step

    # 2.1 Retrieve the IIS paths dynamically on the server side (WinRM)
    Write-Log '----- Retrieving IIS paths per server -----' -Level Sub
    $iisPaths = @()
    foreach ($srv in $allServers) {
        $srvName = [string]$srv.Name
        $ver     = $serverVersion[$srvName]
        try {
            $remoteResult = Invoke-Command -ComputerName $srvName -ErrorAction Stop -ScriptBlock {
                Import-Module WebAdministration -ErrorAction Stop
                $w = Get-Website -Name 'Default Web Site' -ErrorAction SilentlyContinue
                if ($w) {
                    $raw = [string]$w.logFile.directory
                    $expanded = [Environment]::ExpandEnvironmentVariables($raw)
                    [PSCustomObject]@{ RootDir = $expanded; SiteId = [int]$w.Id }
                }
            }
            if ($remoteResult) {
                $iisPaths += [PSCustomObject]@{ Server=$srvName; Version=$ver; Path=$remoteResult.RootDir; SiteId=$remoteResult.SiteId }
            } else {
                Write-Log "  [$srvName] Default Web Site not found (IIS not started?)" -Level Warning
            }
        } catch {
            Write-Log "  [$srvName] Invoke-Command WebAdministration: $($_.Exception.Message)" -Level Warning
        }
    }
    foreach ($p in $iisPaths) {
        Write-Log ("  {0,-12} ({1}) -> {2}  (SiteId={3})" -f $p.Server, $p.Version, $p.Path, $p.SiteId) -Level Info
    }

    # 2.2 Parse the IIS logs filtered by window + vDir + status
    Write-Log '----- Parsing IIS logs -----' -Level Sub
    # Structure per server: @{ vDirName = count }
    $iisStats = @()
    foreach ($p in $iisPaths) {
        # Local server path converted to an administrative share path.
        # logFile.directory points to the root; the .log files are in <root>\W3SVC<SiteId>.
        $w3svcDirs = @()
        try {
            $unc = Convert-LocalPathToUnc -ComputerName $p.Server -LocalPath $p.Path
            if (Test-Path $unc -ErrorAction SilentlyContinue) {
                $expected = Join-Path $unc ("W3SVC" + [string]$p.SiteId)
                if (Test-Path $expected -ErrorAction SilentlyContinue) {
                    $w3svcDirs = @([PSCustomObject]@{ FullName = $expected })
                } else {
                    # Fallback: scan the W3SVC* folders (non-standard case or root already at site level)
                    $found = @(Get-ChildItem -Path $unc -Directory -ErrorAction SilentlyContinue |
                              Where-Object { $_.Name -like 'W3SVC*' })
                    if ($found) { $w3svcDirs = $found }
                    else        { $w3svcDirs = @([PSCustomObject]@{ FullName = $unc }) }
                }
            } else {
                Write-Log "  [$($p.Server)] Path not accessible: $unc" -Level Warning
                $iisStats += [PSCustomObject]@{ Server=$p.Server; Version=$p.Version; Path=$p.Path; TotalRequests=0; ByVdir=@{}; TopClientIps=''; TopStatus=''; Note='Path not accessible' }
                continue
            }
        } catch {
            Write-Log "  [$($p.Server)] Access error: $($_.Exception.Message)" -Level Warning
            continue
        }

        $byVdir       = @{}        # vDir -> count (after filter)
        $clientIps    = @{}        # IP -> global count (after filter)
        $byVdirIPs    = @{}        # vDir -> @{ ip -> count }  (per-protocol drill-down)
        $byStatus     = @{}        # status -> count (after filter)
        $totalReq     = 0          # after filter
        $totalRaw     = 0          # before filter
        $excludedByUri    = 0
        $excludedByUser   = 0
        $excludedByUA     = 0
        $excludedByStatus = 0
        foreach ($d in $w3svcDirs) {
            $logFiles = @(Get-ChildItem -Path $d.FullName -Filter '*.log' -File -ErrorAction SilentlyContinue |
                          Where-Object { $_.LastWriteTime -ge $since })
            foreach ($lf in $logFiles) {
                try {
                    $fields = $null
                    $iDate = -1; $iTime = -1; $iUri = -1; $iIp = -1; $iStatus = -1; $iUser = -1; $iUA = -1
                    Get-Content -LiteralPath $lf.FullName -ErrorAction Stop | ForEach-Object {
                        $line = $_
                        if ($line.StartsWith('#Fields:')) {
                            $fields = $line.Substring(8).Trim().Split(' ')
                            $iDate   = [Array]::IndexOf($fields, 'date')
                            $iTime   = [Array]::IndexOf($fields, 'time')
                            $iUri    = [Array]::IndexOf($fields, 'cs-uri-stem')
                            $iIp     = [Array]::IndexOf($fields, 'c-ip')
                            $iStatus = [Array]::IndexOf($fields, 'sc-status')
                            $iUser   = [Array]::IndexOf($fields, 'cs-username')
                            $iUA     = [Array]::IndexOf($fields, 'cs(User-Agent)')
                            return
                        }
                        if ($line.StartsWith('#') -or [string]::IsNullOrWhiteSpace($line)) { return }
                        if (-not $fields -or $iUri -lt 0) { return }
                        $parts = $line.Split(' ')
                        # Date filter
                        if ($iDate -ge 0 -and $iTime -ge 0 -and $iDate -lt $parts.Count -and $iTime -lt $parts.Count) {
                            $dt = $null
                            try { $dt = [DateTime]::ParseExact("$($parts[$iDate]) $($parts[$iTime])", 'yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) } catch { return }
                            if ($dt -lt $since) { return }
                        }
                        if ($iUri -ge $parts.Count) { return }
                        $uri = $parts[$iUri]
                        if (-not $uri -or -not $uri.StartsWith('/')) { return }

                        $totalRaw++

                        # --- Strict user-identity filters ----------------------
                        # 1) Healthcheck URI
                        $matchedUri = $false
                        foreach ($pat in $excludeUri) {
                            if ($uri -match $pat) { $matchedUri = $true; break }
                        }
                        if ($matchedUri) { $excludedByUri++; return }

                        # 2) Authenticated user identity (missing/anonymous and
                        # known service identities are never real-user traffic).
                        $userRaw = if ($iUser -ge 0 -and $iUser -lt $parts.Count) { $parts[$iUser] } else { '-' }
                        $userValue = if ($null -eq $userRaw) { '' } else { $userRaw.Trim().Trim('"') }
                        $hasAuthenticatedUser = -not [string]::IsNullOrWhiteSpace($userValue) -and
                            $userValue -notin @('-', 'ANONYMOUS', 'ANONYMOUS LOGON')
                        if ($requireAuthenticatedUser -and -not $hasAuthenticatedUser) {
                            $excludedByUser++; return
                        }
                        if ($hasAuthenticatedUser -and
                            (Test-ProtocolLogExcludedIdentity -Identity $userValue -Patterns $excludeUser)) {
                            $excludedByUser++; return
                        }

                        # 3) User-Agent probe
                        if ($iUA -ge 0 -and $iUA -lt $parts.Count) {
                            $ua = $parts[$iUA]
                            if ($ua -and $ua -ne '-') {
                                $matchedUA = $false
                                foreach ($pat in $excludeUA) {
                                    # contains case-insensitive
                                    if ($ua.IndexOf($pat, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                                        $matchedUA = $true; break
                                    }
                                }
                                if ($matchedUA) { $excludedByUA++; return }
                            }
                        }

                        # sc-status extraction (used for filter 4 + byStatus tracking)
                        $stRaw = $null
                        if ($iStatus -ge 0 -and $iStatus -lt $parts.Count) {
                            $stRaw = $parts[$iStatus]
                            if ($stRaw) { $byStatus[$stRaw] = ($byStatus[$stRaw] + 1) }
                        }
                        # 4) sc-status: only count the 2xx requests.
                        # The 401s are the Negotiate/NTLM round-trip 1 challenges
                        # (followed by a 200 when the auth succeeds). Counting them
                        # doubles the real users and lets through the HealthMailbox
                        # probes whose 1st round-trip has no cs-username.
                        if ($requireSuccess -and $stRaw -and -not ($stRaw -match '^2\d\d$')) {
                            $excludedByStatus++; return
                        }

                        # --- Counting ----------------------------------------
                        $seg1 = ($uri -split '/')[1]
                        if (-not $seg1) { return }
                        $key = $vDirsLower[$seg1.ToLowerInvariant()]
                        if (-not $key) {
                            $key = '<Other>'
                        }
                        $byVdir[$key] = ($byVdir[$key] + 1)
                        $totalReq++

                        if ($iIp -ge 0 -and $iIp -lt $parts.Count) {
                            $ip = $parts[$iIp]
                            if ($ip) {
                                $clientIps[$ip] = ($clientIps[$ip] + 1)
                                if (-not $byVdirIPs.ContainsKey($key)) { $byVdirIPs[$key] = @{} }
                                $byVdirIPs[$key][$ip] = ($byVdirIPs[$key][$ip] + 1)
                            }
                        }
                    }
                } catch {
                    Write-Log "  [$($p.Server)] Parsing error in $($lf.Name): $($_.Exception.Message)" -Level Warning
                }
            }
        }

        $topClients = ($clientIps.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 5 |
                       ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ' ; '
        $topStatus  = ($byStatus.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 5 |
                       ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ' ; '
        $iisStats += [PSCustomObject]@{
            Server=$p.Server; Version=$p.Version; Path=$p.Path
            TotalRequests=$totalReq; TotalRaw=$totalRaw
            ExcludedUri=$excludedByUri; ExcludedUser=$excludedByUser; ExcludedUA=$excludedByUA
            ExcludedStatus=$excludedByStatus
            ByVdir=$byVdir; ByVdirIPs=$byVdirIPs
            TopClientIps=$topClients; TopStatus=$topStatus; Note=''
        }
    }

    # 2.3 Log + report per server
    foreach ($s in $iisStats) {
        $vdirSummary = ($s.ByVdir.GetEnumerator() | Sort-Object -Property Value -Descending |
                        ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ' ; '
        $excludedTotal = [int]$s.ExcludedUri + [int]$s.ExcludedUser + [int]$s.ExcludedUA + [int]$s.ExcludedStatus
        Write-Log ("  {0,-12} ({1}): User={2,7} (Total={3}, Probes={4}: URI={5}+User={6}+UA={7}+Status={8}) | {9}" `
            -f $s.Server, $s.Version, $s.TotalRequests, $s.TotalRaw, $excludedTotal,
               $s.ExcludedUri, $s.ExcludedUser, $s.ExcludedUA, $s.ExcludedStatus, $vdirSummary) -Level Info
        Add-Report -Step '26-LogAnalysis' -Target $s.Server -Action 'Stats IIS FrontEnd' -Status Inventoried `
            -AfterValue ("Version=$($s.Version) | UserRequests=$($s.TotalRequests) | TotalRaw=$($s.TotalRaw) | ExcludedProbes=$excludedTotal (URI=$($s.ExcludedUri), User=$($s.ExcludedUser), UA=$($s.ExcludedUA), Status=$($s.ExcludedStatus)) | ByVdir=$vdirSummary | TopClients=$($s.TopClientIps) | TopStatus=$($s.TopStatus)")
    }

    # ========================================================================
    #  PART 3: 2013 vs 2019 summary
    # ========================================================================
    Write-Log ('=' * 70) -Level Step
    Write-Log '===== SUMMARY 2013 vs 2019 =====' -Level Step

    # Versions present (table columns, stable order 2013 -> 2016 -> 2019)
    $versionOrder = @('2013','2016','2019')
    $versionsPresent = @($versionOrder | Where-Object { $serverVersion.ContainsValue($_) })

    # --- IIS: vDir x version matrix (uses the byVdir aggregated per server) ----
    Write-Log '----- IIS: user requests matrix per vDir x version -----' -Level Sub
    # Aggregate byVdir per version
    $iisMatrix = @{}      # vDirKey -> @{ '2013'=N; '2016'=N; '2019'=N }
    $iisByVer  = @{}      # version -> total
    foreach ($s in $iisStats) {
        $v = [string]$s.Version
        if (-not $iisByVer.ContainsKey($v)) { $iisByVer[$v] = 0 }
        $iisByVer[$v] += [int]$s.TotalRequests
        foreach ($kv in $s.ByVdir.GetEnumerator()) {
            $k = [string]$kv.Key
            if (-not $iisMatrix.ContainsKey($k)) {
                $iisMatrix[$k] = @{}
                foreach ($vv in $versionsPresent) { $iisMatrix[$k][$vv] = 0 }
            }
            $iisMatrix[$k][$v] += [int]$kv.Value
        }
    }
    # Ordered table display: configured vDirs first (all of them, even at 0), then <Other> at the end.
    # Force the presence of the configured vDirs in $iisMatrix so that they appear even without a hit.
    $vDirRowsOrder = @($vDirs) + @('<Other>')
    foreach ($v in $vDirRowsOrder) {
        if (-not $iisMatrix.ContainsKey($v)) {
            $iisMatrix[$v] = @{}
            foreach ($vv in $versionsPresent) { $iisMatrix[$v][$vv] = 0 }
        }
    }
    $rowsToShow = @($vDirRowsOrder) +
                  @($iisMatrix.Keys | Where-Object { $_ -notin $vDirRowsOrder })

    # Dynamically builds '  {0,-30} {1,10} {2,10} ...' (each column has its own unique index)
    $cols = @('  {0,-30}')
    for ($i = 1; $i -le $versionsPresent.Count; $i++) { $cols += " {${i},10}" }
    $headerFmt = $cols -join ''
    Write-Log ($headerFmt -f (@('vDir') + $versionsPresent)) -Level Info
    Write-Log ('  ' + ('-' * (30 + 11 * $versionsPresent.Count))) -Level Info
    foreach ($row in $rowsToShow) {
        $vals = @($row) + @($versionsPresent | ForEach-Object { $iisMatrix[$row][$_] })
        Write-Log ($headerFmt -f $vals) -Level Info
    }
    Write-Log ('  ' + ('-' * (30 + 11 * $versionsPresent.Count))) -Level Info
    $totalRow = @('TOTAL') + @($versionsPresent | ForEach-Object { if ($iisByVer.ContainsKey($_)) { $iisByVer[$_] } else { 0 } })
    Write-Log ($headerFmt -f $totalRow) -Level Info

    # CSV line: IIS matrix
    $matrixSummary = ($rowsToShow | ForEach-Object {
        $row = $_
        $cells = $versionsPresent | ForEach-Object { "$_=$($iisMatrix[$row][$_])" }
        "$row[$($cells -join ',')]"
    }) -join ' ; '
    Add-Report -Step '26-LogAnalysis' -Target 'GLOBAL' -Action 'IIS vDir x version matrix' -Status Inventoried `
        -AfterValue $matrixSummary

    # --- SMTP: direction (Receive / Send) x version matrix -----------------
    Write-Log '----- SMTP: valid sessions matrix per direction x version -----' -Level Sub
    # Group HubReceive + FERecv => Receive, HubSend + FESend => Send
    $smtpMatrix = @{
        'Receive' = @{}
        'Send'    = @{}
    }
    foreach ($vv in $versionsPresent) {
        $smtpMatrix['Receive'][$vv] = 0
        $smtpMatrix['Send'][$vv]    = 0
    }
    $smtpByVer = @{}
    foreach ($s in $smtpStats) {
        $v = [string]$s.Version
        if (-not $smtpByVer.ContainsKey($v)) { $smtpByVer[$v] = 0 }
        $smtpByVer[$v] += [int]$s.Sessions
        $dir = if ($s.Kind -match 'Send') { 'Send' } else { 'Receive' }
        $smtpMatrix[$dir][$v] += [int]$s.Sessions
    }
    Write-Log ($headerFmt -f (@('Direction') + $versionsPresent)) -Level Info
    Write-Log ('  ' + ('-' * (30 + 11 * $versionsPresent.Count))) -Level Info
    foreach ($dir in @('Receive','Send')) {
        $vals = @($dir) + @($versionsPresent | ForEach-Object { $smtpMatrix[$dir][$_] })
        Write-Log ($headerFmt -f $vals) -Level Info
    }
    Write-Log ('  ' + ('-' * (30 + 11 * $versionsPresent.Count))) -Level Info
    $totalSmtp = @('TOTAL') + @($versionsPresent | ForEach-Object { if ($smtpByVer.ContainsKey($_)) { $smtpByVer[$_] } else { 0 } })
    Write-Log ($headerFmt -f $totalSmtp) -Level Info

    $smtpSummary = (@('Receive','Send') | ForEach-Object {
        $dir = $_
        $cells = $versionsPresent | ForEach-Object { "$_=$($smtpMatrix[$dir][$_])" }
        "$dir[$($cells -join ',')]"
    }) -join ' ; '
    Add-Report -Step '26-LogAnalysis' -Target 'GLOBAL' -Action 'SMTP Receive/Send x version matrix' -Status Inventoried `
        -AfterValue $smtpSummary

    # --- Cutover diagnosis (legacy = 2013 + 2016 combined) -----------------
    $sumLegacy = {
        param($map)
        $s = 0
        foreach ($k in @('2013','2016')) { if ($map.ContainsKey($k)) { $s += [int]$map[$k] } }
        return $s
    }
    $iisLegacy  = & $sumLegacy $iisByVer
    $iis2019    = if ($iisByVer.ContainsKey('2019'))  { $iisByVer['2019'] }  else { 0 }
    $smtpLegacy = & $sumLegacy $smtpByVer
    $smtp2019   = if ($smtpByVer.ContainsKey('2019')) { $smtpByVer['2019'] } else { 0 }
    $legacyVersions = @($iisByVer.Keys + $smtpByVer.Keys | Where-Object { $_ -in @('2013','2016') } | Sort-Object -Unique)
    $legacyLabel = if ($legacyVersions) { $legacyVersions -join '+' } else { '2013/2016' }

    # Autodiscover focus (key to validate that Step25-CleanupAutodiscoverSCP worked)
    $autodLegacy = 0; $autod2019 = 0
    if ($iisMatrix.ContainsKey('Autodiscover')) {
        foreach ($vv in @('2013','2016')) { if ($iisMatrix['Autodiscover'].ContainsKey($vv)) { $autodLegacy += [int]$iisMatrix['Autodiscover'][$vv] } }
        if ($iisMatrix['Autodiscover'].ContainsKey('2019')) { $autod2019 = [int]$iisMatrix['Autodiscover']['2019'] }
    }

    Write-Log ('-' * 70) -Level Sub
    Write-Log ("  TOTAL IIS:  legacy($legacyLabel)={0,7} | 2019={1,7}" -f $iisLegacy, $iis2019) -Level Info
    Write-Log ("  TOTAL SMTP: legacy($legacyLabel)={0,7} | 2019={1,7}" -f $smtpLegacy, $smtp2019) -Level Info
    Write-Log ("  Autodiscover: legacy={0,7} | 2019={1,7} (SCP key indicator)" -f $autodLegacy, $autod2019) -Level Info
    Write-Log ('-' * 70) -Level Sub
    if ($iisLegacy -eq 0 -and $smtpLegacy -eq 0) {
        Write-Log "  OK: no IIS/SMTP activity detected on the legacy Exchange servers ($legacyLabel)." -Level Success
    } else {
        Write-Log ("  WARNING: residual legacy traffic detected ($legacyLabel) - IIS={0} | SMTP={1}" -f $iisLegacy, $smtpLegacy) -Level Warning
        if ($autodLegacy -gt 0) {
            Write-Log "  -> Autodiscover legacy=$autodLegacy - consider running Step25-CleanupAutodiscoverSCP." -Level Sub
        }
        Write-Log "  -> Check internal DNS, Outlook URLs, AutoDiscover, non-MDM mobile devices, scripts." -Level Sub
    }
    Add-Report -Step '26-LogAnalysis' -Target 'GLOBAL' -Action 'Legacy vs 2019 summary' -Status Inventoried `
        -AfterValue ("IIS legacy($legacyLabel)=$iisLegacy / 2019=$iis2019 | SMTP legacy($legacyLabel)=$smtpLegacy / 2019=$smtp2019 | Autodiscover legacy=$autodLegacy / 2019=$autod2019 | Window=${hours}h")

    # --- 4) Standalone HTML report -----------------------------------------
    # Dedicated subfolder: <OutputFolder>\Step26-ProtocolLogs\<yyyyMMdd_HHmmss>\
    # (same pattern as Step24-HealthChecker to keep the run history).
    $htmlParent = Join-Path $OutputFolder 'Step26-ProtocolLogs'
    $htmlDir    = Join-Path $htmlParent (Get-Date -Format 'yyyyMMdd_HHmmss')
    if (-not (Test-Path $htmlDir)) { New-Item -ItemType Directory -Path $htmlDir -Force | Out-Null }
    $htmlPath = Join-Path $htmlDir 'ProtocolLogsAnalysis.html'
    New-ProtocolLogsHtml -OutputPath $htmlPath `
        -Hours $hours -Since $since -ServerVersion $serverVersion `
        -VDirsOrder $vDirRowsOrder -VersionsPresent $versionsPresent `
        -IisStats $iisStats -SmtpStats $smtpStats `
        -IisMatrix $iisMatrix -SmtpMatrix $smtpMatrix `
        -IisByVer $iisByVer -SmtpByVer $smtpByVer `
        -IisLegacy $iisLegacy -Iis2019 $iis2019 -SmtpLegacy $smtpLegacy -Smtp2019 $smtp2019 `
        -AutodLegacy $autodLegacy -Autod2019 $autod2019 -LegacyLabel $legacyLabel
    Write-Log "Detailed HTML report: $htmlPath" -Level Success
    try { Invoke-Item -Path $htmlPath -ErrorAction Stop } catch { Write-Log "  (Unable to open the HTML automatically: $($_.Exception.Message))" -Level Sub }
    Add-Report -Step '26-LogAnalysis' -Target 'GLOBAL' -Action 'Detailed HTML' -Status Inventoried -AfterValue $htmlPath

    Save-Report -StepName 'Step26-AnalyzeProtocolLogs' | Out-Null
    return 0
}

function Convert-LocalPathToUnc {
    <#
        Converts 'D:\Logs\IIS\...' (local on the remote server) to
        '\\SERVER\D$\Logs\IIS\...' for administrative access via SMB.
        If the path is already UNC or points to the local machine, it is returned as is.
    #>
    [CmdletBinding()]
    param([string]$ComputerName, [string]$LocalPath)
    if (-not $LocalPath) { return $null }
    if ($LocalPath.StartsWith('\\')) { return $LocalPath }
    if ($LocalPath -match '^([A-Za-z]):(\\.*)$') {
        $drive = $Matches[1]
        $rest  = $Matches[2]
        if ($ComputerName -eq $env:COMPUTERNAME) {
            return $LocalPath
        }
        return "\\$ComputerName\$drive`$$rest"
    }
    return $LocalPath
}

function New-ProtocolLogsHtml {
    <#
        Builds ProtocolLogsAnalysis.html with the shared report theme of the module
        (Get-ExHtmlHead / Get-ExHtmlHero / Get-ExHtmlFooter): header with the key figures,
        IIS and SMTP matrices, clickable server tiles and one accordion per server.
    #>
    [CmdletBinding()]
    param(
        [string]$OutputPath,
        [int]$Hours,
        [DateTime]$Since,
        [hashtable]$ServerVersion,
        [string[]]$VDirsOrder,
        [string[]]$VersionsPresent,
        [array]$IisStats,
        [array]$SmtpStats,
        [hashtable]$IisMatrix,
        [hashtable]$SmtpMatrix,
        [hashtable]$IisByVer,
        [hashtable]$SmtpByVer,
        [int]$IisLegacy, [int]$Iis2019, [int]$SmtpLegacy, [int]$Smtp2019,
        [int]$AutodLegacy, [int]$Autod2019,
        [string]$LegacyLabel
    )

    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
    function _H { param([string]$s) if ($null -eq $s) { return '' }
        [System.Web.HttpUtility]::HtmlEncode($s) }

    # Colors for the colored tiles (theme aligned with the other reports)
    $colorBad   = 'var(--cp-danger)'    # red - residual legacy traffic
    $colorGood  = 'var(--cp-success)'   # green - healthy 2019
    $colorWarn  = 'var(--cp-warning)'   # orange - Autodiscover SCP signal
    $colorInfo  = 'var(--cp-info)'   # blue - neutral info
    $colorLegacy= 'var(--cp-warning)'   # dark orange - legacy server
    $color2019  = 'var(--cp-success)'   # medium green - 2019 server

    # Stats index per server
    $iisByServer = @{}
    foreach ($s in $IisStats) { $iisByServer[[string]$s.Server] = $s }
    $smtpByServer = @{}
    foreach ($s in $SmtpStats) {
        $k = [string]$s.Server
        if (-not $smtpByServer.ContainsKey($k)) { $smtpByServer[$k] = @() }
        $smtpByServer[$k] += $s
    }
    $serverNames = @($ServerVersion.Keys | Sort-Object)

    # Global totals for the header
    $totalIIS  = $IisLegacy + $Iis2019
    $totalSMTP = $SmtpLegacy + $Smtp2019
    $now       = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((Get-ExHtmlHead -Title 'Exchange 2019 migration | Protocol-log analysis'))

    # --- HEADER: what the window shows, at a glance ----------------------
    # The goal of the cut-over is zero real-user traffic on the legacy servers.
    $legacyTotal = $IisLegacy + $SmtpLegacy
    $tiles = @(
        (Get-ExHtmlTile -Label "Real-user traffic on legacy ($LegacyLabel)" -Value (Format-ExNumber $legacyTotal) -Note ('{0} IIS requests + {1} SMTP sessions - target 0' -f (Format-ExNumber $IisLegacy), (Format-ExNumber $SmtpLegacy)) -Kind Primary -Icon 'server')
        (Get-ExHtmlTile -Label 'IIS requests on Exchange 2019' -Value (Format-ExNumber $Iis2019) -Note 'Authenticated users, HTTP 2xx, probes excluded' -Kind Tinted -Icon 'user')
        (Get-ExHtmlTile -Label 'SMTP sessions on Exchange 2019' -Value (Format-ExNumber $Smtp2019) -Note 'Sessions with a user MAIL FROM' -Kind Tinted -Icon 'mail')
        (Get-ExHtmlTile -Label 'Autodiscover on legacy' -Value (Format-ExNumber $AutodLegacy) -Note ('SCP indicator - Exchange 2019: {0}' -f (Format-ExNumber $Autod2019)) -Tone $(if ($AutodLegacy -gt 0) { $colorWarn } else { $colorGood }) -Icon 'alert')
    )
    $bands = foreach ($b in @(
            @('IIS - legacy', $IisLegacy, $totalIIS, $colorBad), @('IIS - Exchange 2019', $Iis2019, $totalIIS, $colorGood),
            @('SMTP - legacy', $SmtpLegacy, $totalSMTP, $colorBad), @('SMTP - Exchange 2019', $Smtp2019, $totalSMTP, $colorGood))) {
        $pct = if ($b[2] -gt 0) { [Math]::Round(100.0 * $b[1] / $b[2]) } else { 0 }
        "<div><div class='band-head'><span>$($b[0])</span><span>$(Format-ExNumber $b[1]) / $pct%</span></div><div class='track'><div class='fill' style='width:$pct%;background:$($b[3])'></div></div></div>"
    }
    $box = "Last <strong>$Hours h</strong><br>From $($Since.ToString('yyyy-MM-dd HH:mm'))<br>To $now"
    $subtitle = 'Real-user IIS and SMTP traffic per Exchange version, from the IIS logs (Default Web Site) and the SMTP protocol logs. Monitoring probes (AMProbe, HealthMailbox, system mailboxes...), anonymous requests and sessions without a user MAIL FROM are excluded.'
    [void]$sb.Append((Get-ExHtmlHero -Title 'Protocol-log analysis' -Subtitle $subtitle -BoxLabel 'Analysis window' -BoxHtml $box -Tiles $tiles `
        -ExtraHtml ("<p class='caption'>Share of the traffic, legacy versus Exchange 2019</p><div class='distribution'>" + ($bands -join '') + '</div>')))

    # --- CARD: IIS matrix ------------------------------------------------
    [void]$sb.AppendLine("<div class='card'>")
    [void]$sb.AppendLine("  <div class='card-hdr'>IIS &mdash; vDir x version matrix (user requests)</div>")
    [void]$sb.AppendLine("  <div class='tbl-wrap'><table>")
    [void]$sb.Append("<thead><tr><th>vDir</th>")
    foreach ($v in $VersionsPresent) { [void]$sb.Append("<th class='num'>$v</th>") }
    [void]$sb.AppendLine("</tr></thead><tbody>")
    foreach ($row in $VDirsOrder) {
        [void]$sb.Append("<tr><td>$(_H $row)</td>")
        foreach ($v in $VersionsPresent) {
            $val = 0
            if ($IisMatrix.ContainsKey($row) -and $IisMatrix[$row].ContainsKey($v)) { $val = [int]$IisMatrix[$row][$v] }
            $cls = if ($val -eq 0) { 'num zero' } else { 'num' }
            [void]$sb.Append("<td class='$cls'>$val</td>")
        }
        [void]$sb.AppendLine("</tr>")
    }
    [void]$sb.Append("<tr class='total'><td>TOTAL</td>")
    foreach ($v in $VersionsPresent) {
        $tot = if ($IisByVer.ContainsKey($v)) { [int]$IisByVer[$v] } else { 0 }
        [void]$sb.Append("<td class='num'>$tot</td>")
    }
    [void]$sb.AppendLine("</tr></tbody></table></div></div>")

    # --- CARD: SMTP matrix -----------------------------------------------
    [void]$sb.AppendLine("<div class='card'>")
    [void]$sb.AppendLine("  <div class='card-hdr'>SMTP &mdash; Direction x version matrix (valid sessions)</div>")
    [void]$sb.AppendLine("  <div class='tbl-wrap'><table>")
    [void]$sb.Append("<thead><tr><th>Direction</th>")
    foreach ($v in $VersionsPresent) { [void]$sb.Append("<th class='num'>$v</th>") }
    [void]$sb.AppendLine("</tr></thead><tbody>")
    foreach ($dir in @('Receive','Send')) {
        [void]$sb.Append("<tr><td>$dir</td>")
        foreach ($v in $VersionsPresent) {
            $val = 0
            if ($SmtpMatrix.ContainsKey($dir) -and $SmtpMatrix[$dir].ContainsKey($v)) { $val = [int]$SmtpMatrix[$dir][$v] }
            $cls = if ($val -eq 0) { 'num zero' } else { 'num' }
            [void]$sb.Append("<td class='$cls'>$val</td>")
        }
        [void]$sb.AppendLine("</tr>")
    }
    [void]$sb.Append("<tr class='total'><td>TOTAL</td>")
    foreach ($v in $VersionsPresent) {
        $tot = if ($SmtpByVer.ContainsKey($v)) { [int]$SmtpByVer[$v] } else { 0 }
        [void]$sb.Append("<td class='num'>$tot</td>")
    }
    [void]$sb.AppendLine("</tr></tbody></table></div></div>")

    # --- CARD: Clickable server tiles ------------------------------------
    [void]$sb.AppendLine("<div class='card'>")
    [void]$sb.AppendLine("  <div class='card-hdr'>Servers (click a tile to filter the detail below)</div>")
    [void]$sb.AppendLine("  <div class='tiles'>")
    foreach ($name in $serverNames) {
        $ver = $ServerVersion[$name]
        $bg  = if ($ver -in @('2013','2016')) { $colorLegacy } else { $color2019 }
        $iisUser = if ($iisByServer.ContainsKey($name)) { [int]$iisByServer[$name].TotalRequests } else { 0 }
        $smtpVal = 0
        if ($smtpByServer.ContainsKey($name)) { foreach ($x in $smtpByServer[$name]) { $smtpVal += [int]$x.Sessions } }
        [void]$sb.AppendLine("    <div class='tile tile-srv clickable' data-server='$(_H $name)' style='--tone:$bg' onclick='toggleSrv(this)'>")
        [void]$sb.AppendLine("      <div class='tile-hdr'>$(_H $name) &middot; $ver</div>")
        [void]$sb.AppendLine("      <div class='tile-row'><span class='lbl'>IIS</span><span class='val'>$iisUser</span></div>")
        [void]$sb.AppendLine("      <div class='tile-row'><span class='lbl'>SMTP</span><span class='val'>$smtpVal</span></div>")
        [void]$sb.AppendLine("    </div>")
    }
    [void]$sb.AppendLine("  </div>")
    [void]$sb.AppendLine("  <div class='filter-state' id='filterInfo'>Active filter: <strong id='filterName'></strong><button class='clear' onclick='clearFilter()'>Clear</button></div>")
    [void]$sb.AppendLine("</div>")

    # --- CARD: Detail per server (accordions) ----------------------------
    [void]$sb.AppendLine("<div class='card'>")
    [void]$sb.AppendLine("  <div class='card-hdr'>Detail per server &mdash; click to expand</div>")
    [void]$sb.AppendLine("  <div class='card-body'>")

    foreach ($name in $serverNames) {
        $ver = $ServerVersion[$name]
        $iisS  = $iisByServer[$name]
        $smtpS = if ($smtpByServer.ContainsKey($name)) { @($smtpByServer[$name]) } else { @() }

        $iisUserCnt = if ($iisS) { [int]$iisS.TotalRequests } else { 0 }
        $iisRawCnt  = if ($iisS) { [int]$iisS.TotalRaw } else { 0 }
        $iisExcl    = if ($iisS) {
            [int]$iisS.ExcludedUri + [int]$iisS.ExcludedUser + [int]$iisS.ExcludedUA + [int]$iisS.ExcludedStatus
        } else { 0 }
        $smtpValTotal = 0; $smtpRawTotal = 0
        $smtpSystemTotal = 0; $smtpNoMailFromTotal = 0
        foreach ($x in $smtpS) {
            $smtpValTotal += [int]$x.Sessions
            $smtpRawTotal += [int]$x.SessionsTotal
            $smtpSystemTotal += [int]$x.ExcludedSystem
            $smtpNoMailFromTotal += [int]$x.ExcludedNoMailFrom
        }

        [void]$sb.AppendLine("<details data-server='$(_H $name)'>")
        [void]$sb.AppendLine("  <summary>")
        [void]$sb.AppendLine("    <span class='acc-icon'>&#9654;</span>")
        [void]$sb.AppendLine("    <span class='bdg bdg-$ver'>$ver</span>")
        [void]$sb.AppendLine("    <span class='acc-name'>$(_H $name)</span>")
        $headerSummary = "IIS users=$iisUserCnt / raw=$iisRawCnt ($iisExcl non-user filtered) &middot; SMTP valid=$smtpValTotal / raw=$smtpRawTotal (system=$smtpSystemTotal, without MAIL FROM=$smtpNoMailFromTotal)"
        [void]$sb.AppendLine("    <span class='acc-summary'>$headerSummary</span>")
        [void]$sb.AppendLine("  </summary>")
        [void]$sb.AppendLine("  <div class='srv-body'>")

        # ===== IIS - block per vDir =====
        [void]$sb.AppendLine("    <h3>IIS &mdash; requests per virtual directory</h3>")
        if ($iisS -and $iisS.ByVdir.Count -gt 0) {
            $anyVdir = $false
            foreach ($kv in ($iisS.ByVdir.GetEnumerator() | Sort-Object -Property Value -Descending)) {
                $vdir = [string]$kv.Key
                $vCnt = [int]$kv.Value
                if ($vCnt -eq 0) { continue }
                $anyVdir = $true
                [void]$sb.AppendLine("    <div class='sub-title'>$(_H $vdir)<span class='sub-meta'>$vCnt user requests</span></div>")
                if ($iisS.ByVdirIPs -and $iisS.ByVdirIPs.ContainsKey($vdir) -and $iisS.ByVdirIPs[$vdir].Count -gt 0) {
                    [void]$sb.AppendLine("    <div class='tbl-wrap'><table>")
                    [void]$sb.AppendLine("      <thead><tr><th>Client IP</th><th class='num'>Requests</th></tr></thead><tbody>")
                    foreach ($ipKv in ($iisS.ByVdirIPs[$vdir].GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10)) {
                        [void]$sb.AppendLine("      <tr><td>$(_H ([string]$ipKv.Key))</td><td class='num'>$([int]$ipKv.Value)</td></tr>")
                    }
                    [void]$sb.AppendLine("    </tbody></table></div>")
                } else {
                    [void]$sb.AppendLine("    <div class='note'>No client IP identified.</div>")
                }
            }
            if (-not $anyVdir) { [void]$sb.AppendLine("    <div class='note'>No IIS protocol with user traffic in the window.</div>") }
            if ($iisS.TopStatus) { [void]$sb.AppendLine("    <div class='note' style='margin-top:10px'>Top HTTP status (server-wide): $(_H $iisS.TopStatus)</div>") }
        } else {
            [void]$sb.AppendLine("    <div class='note'>No IIS data for this server.</div>")
        }

        # ===== SMTP - block per kind (same structure as IIS) =====
        [void]$sb.AppendLine("    <h3>SMTP &mdash; valid sessions per connector</h3>")
        if ($smtpS.Count -gt 0) {
            $anyConn = $false
            foreach ($kind in @('HubReceive','FERecv','HubSend','FESend')) {
                $sk = $smtpS | Where-Object { $_.Kind -eq $kind }
                if (-not $sk) { continue }
                $skTotal = [int]$sk.Sessions
                $skRaw   = [int]$sk.SessionsTotal
                $skSystem = [int]$sk.ExcludedSystem
                $skNoMailFrom = [int]$sk.ExcludedNoMailFrom
                [void]$sb.AppendLine("    <div class='sub-title'>$kind<span class='sub-meta'>users=$skTotal / total=$skRaw &middot; system=$skSystem &middot; without MAIL FROM=$skNoMailFrom</span></div>")
                if ($sk.ByConnector -and $sk.ByConnector.Count -gt 0) {
                    $anyConn = $true
                    [void]$sb.AppendLine("    <div class='tbl-wrap'><table>")
                    [void]$sb.AppendLine("      <thead><tr><th>Connector ID</th><th class='num'>Valid sessions</th></tr></thead><tbody>")
                    foreach ($kv in ($sk.ByConnector.GetEnumerator() | Sort-Object -Property Value -Descending)) {
                        $val = [int]$kv.Value
                        $cls = if ($val -eq 0) { 'num zero' } else { 'num' }
                        [void]$sb.AppendLine("      <tr><td>$(_H ([string]$kv.Key))</td><td class='$cls'>$val</td></tr>")
                    }
                    [void]$sb.AppendLine("    </tbody></table></div>")
                } else {
                    [void]$sb.AppendLine("    <div class='note'>No valid session on this log stream.</div>")
                }
                if ($sk.TopRemoteIps) { [void]$sb.AppendLine("    <div class='note'>Top remote IPs: $(_H $sk.TopRemoteIps)</div>") }
            }
            if (-not $anyConn) { [void]$sb.AppendLine("    <div class='note'>No SMTP connector with a valid session in the window.</div>") }
        } else {
            [void]$sb.AppendLine("    <div class='note'>No SMTP data for this server.</div>")
        }

        [void]$sb.AppendLine("  </div>")
        [void]$sb.AppendLine("</details>")
    }

    [void]$sb.AppendLine("  </div>")
    [void]$sb.AppendLine("</div>")

    # --- Back to top + JS -------------------------------------------------
    [void]$sb.AppendLine(@'
<script>
function toggleSrv(tile){
    var name = tile.getAttribute('data-server');
    var allTiles = document.querySelectorAll('.tile.clickable');
    var allDet   = document.querySelectorAll('details[data-server]');
    var info     = document.getElementById('filterInfo');
    if (tile.classList.contains('tile-active')) {
        clearFilter();
        return;
    }
    allTiles.forEach(function(t){
        t.classList.remove('tile-active');
        t.classList.add('dimmed');
    });
    tile.classList.add('tile-active');
    tile.classList.remove('dimmed');
    allDet.forEach(function(d){
        if (d.getAttribute('data-server') === name) {
            d.classList.remove('filter-hidden');
            d.open = true;
        } else {
            d.classList.add('filter-hidden');
            d.open = false;
        }
    });
    document.getElementById('filterName').innerText = name;
    info.classList.add('on');
}
function clearFilter(){
    document.querySelectorAll('.tile.clickable').forEach(function(t){
        t.classList.remove('tile-active');
        t.classList.remove('dimmed');
    });
    document.querySelectorAll('details[data-server]').forEach(function(d){
        d.classList.remove('filter-hidden');
    });
    document.getElementById('filterInfo').classList.remove('on');
}
</script>
'@)
    [void]$sb.Append((Get-ExHtmlFooter -Text 'This file contains server names and client IP addresses: store and share it accordingly.'))

    [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($true))
}
