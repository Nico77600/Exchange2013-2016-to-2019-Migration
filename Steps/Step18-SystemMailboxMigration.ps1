<#
.SYNOPSIS
    Step 18 - System mailbox migration to Exchange 2019.

.DESCRIPTION
    Migrates the Arbitration, AuditLog, AuxAuditLog and Discovery mailboxes
    from the legacy Exchange servers to an Exchange 2019 database picked at
    random among the available databases.

    - Automatic completion (no manual suspension).
    - Idempotent: if the mailbox is already on Exchange 2019 -> AlreadyDone.
    - Apply mode: submits the requests, then waits for completion (30-minute timeout).
    - If a failed or suspended request exists, it is removed before resubmission.

    Simulate and Apply modes also generate an HTML plan of the system
    mailboxes (SystemMailboxPlan.html) before the 2019 database precheck.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    Part of : Exchange 2013/2016 to 2019 Migration - Deploy-Exchange2019.ps1
#>

function Invoke-Step {
    [CmdletBinding()]
    param(
        [string]$Mode,
        [string]$OutputFolder,
        [string]$CsvFolder,
        [hashtable]$Config,
        [System.Management.Automation.PSCredential]$Credential
    )

    # =========================================================================
    # Helper: system mailbox plan HTML report (Apply / Simulate)
    # =========================================================================
    function New-SystemMailboxPlanHtml {
        param(
            [System.Collections.IEnumerable]$Mailboxes,
            [int]$TargetDbCount,
            [string]$OutputPath,
            [string]$ReportMode
        )

        function Esc {
            param([string]$s)
            if ([string]::IsNullOrEmpty($s)) { return '' }
            $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
        }

        $now = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

        $typeClr = @{
            'ArbitrationMailbox' = '#7c3aed'
            'AuditLogMailbox'    = '#db2777'
            'AuxAuditLogMailbox' = '#ea580c'
            'DiscoveryMailbox'   = '#0891b2'
            'default'            = '#5f6b7a'
        }

        # Enrichment (Exchange version + status + move request)
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($m in $Mailboxes) {
            $srv = Get-ExchangeServer -Identity $m.ServerName -ErrorAction SilentlyContinue
            $ver = '<unknown>'
            $isOn2019 = $false
            if ($srv) {
                switch -Wildcard ([string]$srv.AdminDisplayVersion) {
                    'Version 15.0*' { $ver = 'Exchange 2013' }
                    'Version 15.1*' { $ver = 'Exchange 2016' }
                    'Version 15.2*' { $ver = 'Exchange 2019'; $isOn2019 = $true }
                    default         { $ver = [string]$srv.AdminDisplayVersion }
                }
            }
            $req = Get-MoveRequest -Identity $m.Identity -ErrorAction SilentlyContinue
            $rows.Add([PSCustomObject]@{
                Name                 = [string]$m.Name
                DisplayName          = [string]$m.DisplayName
                RecipientTypeDetails = [string]$m.RecipientTypeDetails
                ServerName           = [string]$m.ServerName
                ServerVersion        = $ver
                Database             = [string]$m.Database
                IsOn2019             = $isOn2019
                MoveRequestStatus    = if ($req) { [string]$req.Status } else { '' }
            })
        }

        $countByType = [ordered]@{}
        foreach ($t in 'ArbitrationMailbox','AuditLogMailbox','AuxAuditLogMailbox','DiscoveryMailbox') {
            $countByType[$t] = @($rows | Where-Object { $_.RecipientTypeDetails -eq $t }).Count
        }
        $totalToMigrate = @($rows | Where-Object { -not $_.IsOn2019 }).Count
        $totalAlready   = @($rows | Where-Object { $_.IsOn2019 }).Count

        $sb = [System.Text.StringBuilder]::new(60000)

        # ---- HEAD and HEADER: shared report theme of the module ------------
        # Table classes of this report, on top of the shared theme.
        $extraStyle = '.td-dn{font-weight:600}.st-mig{color:var(--cp-danger);font-weight:700}.st-ok{color:var(--cp-success);font-weight:700}' +
                      '.ver-13{color:var(--cp-danger);font-weight:600}.ver-16{color:var(--cp-warning);font-weight:600}.ver-19{color:var(--cp-success);font-weight:700}'
        [void]$sb.Append((Get-ExHtmlHead -Title 'Exchange 2019 migration | System mailbox plan' -ExtraStyle $extraStyle))
        $pills = foreach ($t in $countByType.Keys) { "<span class='pill' style='--tone:$($typeClr[$t])'>$t $($countByType[$t])</span>" }
        $tiles = @(
            (Get-ExHtmlTile -Label 'System mailboxes to move' -Value (Format-ExNumber $totalToMigrate) -Note 'Not on Exchange 2019 yet' -Kind Primary -Icon 'mail')
            (Get-ExHtmlTile -Label 'Already on Exchange 2019' -Value (Format-ExNumber $totalAlready) -Note 'Nothing to do' -Kind Tinted -Icon 'check')
            (Get-ExHtmlTile -Label 'System mailboxes' -Value (Format-ExNumber $rows.Count) -Note 'Arbitration, audit log, auxiliary audit log, discovery' -Kind Tinted -Icon 'list')
            (Get-ExHtmlTile -Label 'Target databases' -Value (Format-ExNumber $TargetDbCount) -Note 'Exchange 2019 databases available' -Icon 'server')
        )
        $box = "Mode <strong>$ReportMode</strong><br>$now"
        [void]$sb.Append((Get-ExHtmlHero -Title 'System mailbox plan' -Subtitle 'The system mailboxes of the organisation, their current server and their move to Exchange 2019 (Step 18). They must be on Exchange 2019 before the legacy servers are removed.' `
            -BoxLabel 'Plan' -BoxHtml $box -Tiles $tiles -ExtraHtml ("<div class='hdr-r2'>" + ($pills -join '') + '</div>')))

        # ---- DETAIL TABLE ---------------------------------------------------
        [void]$sb.AppendLine("<div class='card'><div class='card-hdr'>System mailbox detail</div>")
        [void]$sb.AppendLine("<div class='tbl-wrap'><table>")
        [void]$sb.AppendLine("  <thead><tr>")
        [void]$sb.AppendLine("    <th style='min-width:240px'>Name</th>")
        [void]$sb.AppendLine("    <th style='min-width:200px'>Display name</th>")
        [void]$sb.AppendLine("    <th style='min-width:160px'>Type</th>")
        [void]$sb.AppendLine("    <th style='min-width:140px'>Source server</th>")
        [void]$sb.AppendLine("    <th style='min-width:120px'>Version</th>")
        [void]$sb.AppendLine("    <th style='min-width:200px'>Source database</th>")
        [void]$sb.AppendLine("    <th style='min-width:130px'>Status</th>")
        [void]$sb.AppendLine("    <th style='min-width:120px'>Move request</th>")
        [void]$sb.AppendLine("  </tr></thead><tbody>")

        foreach ($r in ($rows | Sort-Object RecipientTypeDetails, Name)) {
            $tc = if ($typeClr[$r.RecipientTypeDetails]) { $typeClr[$r.RecipientTypeDetails] } else { $typeClr['default'] }
            $verClass = switch ($r.ServerVersion) {
                'Exchange 2013' { 'ver-13' }
                'Exchange 2016' { 'ver-16' }
                'Exchange 2019' { 'ver-19' }
                default         { '' }
            }
            $stClass = if ($r.IsOn2019) { 'st-ok' }  else { 'st-mig' }
            $stTxt   = if ($r.IsOn2019) { 'Already on 2019' } else { 'To migrate' }
            $mr      = if ($r.MoveRequestStatus) { Esc $r.MoveRequestStatus } else { "<span class='em'>&mdash;</span>" }

            [void]$sb.AppendLine("    <tr>")
            [void]$sb.AppendLine("      <td class='td-dn'>$(Esc $r.Name)</td>")
            [void]$sb.AppendLine("      <td><span class='val'>$(Esc $r.DisplayName)</span></td>")
            [void]$sb.AppendLine("      <td><span class='bdg' style='background:$tc'>$(Esc $r.RecipientTypeDetails)</span></td>")
            [void]$sb.AppendLine("      <td><span class='val'>$(Esc $r.ServerName)</span></td>")
            [void]$sb.AppendLine("      <td><span class='$verClass'>$(Esc $r.ServerVersion)</span></td>")
            [void]$sb.AppendLine("      <td><span class='val'>$(Esc $r.Database)</span></td>")
            [void]$sb.AppendLine("      <td><span class='$stClass'>$stTxt</span></td>")
            [void]$sb.AppendLine("      <td>$mr</td>")
            [void]$sb.AppendLine("    </tr>")
        }
        [void]$sb.AppendLine("</tbody></table></div></div>")

        [void]$sb.Append((Get-ExHtmlFooter -Text 'This file contains the names of the system mailboxes, servers and databases: store and share it accordingly.'))

        $sb.ToString() | Set-Content -Path $OutputPath -Encoding UTF8
    }

    Initialize-ExchangeShell -Credential $Credential

    $gc = $Config.PreferredGCs | Select-Object -First 1

    # Exchange 2019 databases available as targets
    $ex2019Dbs = @(Get-MailboxDatabase -ErrorAction Stop |
                   Where-Object { Test-IsExchange2019Server -ServerName $_.Server.Name })

    # System mailbox collection by type
    # DiscoveryMailbox has no dedicated switch: filter on RecipientTypeDetails
    $mbxArb = @(Get-Mailbox -Arbitration  -ErrorAction SilentlyContinue)
    $mbxAud = @(Get-Mailbox -AuditLog     -ErrorAction SilentlyContinue)
    $mbxAux = @(Get-Mailbox -AuxAuditLog  -ErrorAction SilentlyContinue)
    $mbxDis = @(Get-Mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue |
                Where-Object { $_.RecipientTypeDetails -eq 'DiscoveryMailbox' })
    $allMbx = @($mbxArb) + @($mbxAud) + @($mbxAux) + @($mbxDis)

    Write-StepBanner -StepName '18' -Title 'System mailbox migration to Exchange 2019' -Mode $Mode -Actions @(
        "Arbitration=$($mbxArb.Count) | AuditLog=$($mbxAud.Count) | AuxAuditLog=$($mbxAux.Count) | Discovery=$($mbxDis.Count)",
        "Target database: Exchange 2019, picked at random ($($ex2019Dbs.Count) database(s) available)",
        "New-MoveRequest without suspension - automatic completion",
        "Wait for completion of all requests (Apply only, 30-minute timeout)"
    )

    if (-not $allMbx) {
        Write-Log "No system mailbox found" -Level Warning
        Save-Report -StepName 'Step18-SystemMailboxMigration'
        return 0
    }

    # ==========================================================================
    # 0) HTML plan generation (Apply / Simulate)
    #    Generated BEFORE the 2019 database precheck to provide a preview even
    #    if Step10 (database creation) has not run yet.
    # ==========================================================================
    if ($Mode -in 'Simulate','Apply') {
        $htmlPath = Join-Path $OutputFolder 'SystemMailboxPlan.html'
        New-SystemMailboxPlanHtml -Mailboxes $allMbx -TargetDbCount $ex2019Dbs.Count `
                                  -OutputPath $htmlPath -ReportMode $Mode
        Write-Log "System mailbox HTML report: $htmlPath" -Level Success
        Add-Report -Step '18-SysMbx' -Target 'GLOBAL' -Action 'Export-Html' -Status Success `
                   -AfterValue $htmlPath -Detail 'Visual plan of the system mailboxes to migrate'
    }

    # In Apply/Simulate a target 2019 database is required to submit the move requests.
    # In Inventory the current state is simply reported.
    if (-not $ex2019Dbs -and $Mode -ne 'Inventory') {
        Write-Log "No Exchange 2019 database available - aborting (HTML plan generated anyway)" -Level Error
        Save-Report -StepName 'Step18-SystemMailboxMigration'
        return 1
    }

    # ==========================================================================
    # 1) Move request submission
    # ==========================================================================
    foreach ($mbx in $allMbx) {
        $identity   = $mbx.Identity
        $serverName = [string]$mbx.ServerName
        $database   = [string]$mbx.Database
        $targetDb   = if ($ex2019Dbs) { ($ex2019Dbs | Get-Random).Name } else { '<no 2019 database available>' }

        Invoke-Action -Step '18-SysMbx' -Target $identity -Action 'New-MoveRequest' `
            -Detail "TargetDatabase=$targetDb" `
            -InventoryScript {
                $srv = Get-ExchangeServer -Identity $serverName -ErrorAction SilentlyContinue
                $ver       = '<unknown>'
                $isOn2019  = $false
                if ($srv) {
                    switch -Wildcard ([string]$srv.AdminDisplayVersion) {
                        'Version 15.0*' { $ver = 'Exchange 2013' }
                        'Version 15.1*' { $ver = 'Exchange 2016' }
                        'Version 15.2*' { $ver = 'Exchange 2019'; $isOn2019 = $true }
                        default         { $ver = [string]$srv.AdminDisplayVersion }
                    }
                }
                $migState = if ($isOn2019) { 'ALREADY ON 2019 (skip)' } else { 'TO MIGRATE (New-MoveRequest required)' }
                $existing = Get-MoveRequest -Identity $identity -ErrorAction SilentlyContinue
                $reqInfo  = if ($existing) { "; MoveRequest=$($existing.Status)" } else { '' }
                "Server=$serverName [$ver]; Database=$database; Status=$migState$reqInfo"
            } `
            -PreCheckScript {
                return (Test-IsExchange2019Server -ServerName $serverName)
            } `
            -ActionScript {
                # Clean up a previous failed/suspended request
                $existing = Get-MoveRequest -Identity $identity -ErrorAction SilentlyContinue
                if ($existing) {
                    if ($existing.Status -in @('Completed','CompletedWithWarning')) { return }
                    if ($existing.Status -notin @('Failed','Suspended')) {
                        Write-Log "  Request already in progress ($($existing.Status)) for $identity" -Level Sub
                        return
                    }
                    Remove-MoveRequest -Identity $identity -Confirm:$false -ErrorAction Stop
                }

                $params = @{
                    Identity        = $identity
                    TargetDatabase  = $targetDb
                    AllowLargeItems = $true
                    BadItemLimit    = 100
                    ErrorAction     = 'Stop'
                }
                if ($gc) { $params['DomainController'] = $gc }
                New-MoveRequest @params -WhatIf:([bool]($Mode -eq 'Simulate'))
            }
    }

    # ==========================================================================
    # 2) Wait for completion (Apply only)
    # ==========================================================================
    if ($Mode -eq 'Apply') {
        Write-Log "Waiting for the system move requests to complete (30-minute timeout)..." -Level Sub
        $timeout    = (Get-Date).AddMinutes(30)
        $doneStatus = @('Completed','CompletedWithWarning','Failed','Suspended')

        do {
            Start-Sleep -Seconds 30
            $pending = @()
            foreach ($mbx in $allMbx) {
                $req = Get-MoveRequest -Identity $mbx.Identity -ErrorAction SilentlyContinue
                if ($req -and $req.Status -notin $doneStatus) { $pending += $req }
            }
            if ($pending.Count -gt 0) {
                Write-Log "  Pending: $($pending.Count) request(s)..." -Level Sub
            }
        } while ($pending.Count -gt 0 -and (Get-Date) -lt $timeout)

        # Completion report
        foreach ($mbx in $allMbx) {
            $req = Get-MoveRequest -Identity $mbx.Identity -ErrorAction SilentlyContinue
            if (-not $req) { continue }
            $ok = $req.Status -in @('Completed','CompletedWithWarning')
            Add-Report -Step '18-SysMbx' -Target $mbx.Identity -Action 'MoveRequest-Result' `
                -Status  $(if ($ok) { 'Success' } else { 'Failed' }) `
                -Detail  "Status=$($req.Status); TargetDB=$($req.TargetDatabase)"
        }
    }

    Save-Report -StepName 'Step18-SystemMailboxMigration'
    return 0
}
