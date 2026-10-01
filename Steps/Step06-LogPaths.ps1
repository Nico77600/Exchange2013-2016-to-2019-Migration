<#
.SYNOPSIS
    Step 06 - Reconfigure the Exchange 2019 log paths.

.DESCRIPTION
    Reconfigures the log paths of each Exchange 2019 server.

    Targets (from Config.LogPaths):
        - FrontendTransport (Set-FrontendTransportService)
        - Hub Transport     (Set-TransportService)
        - Mailbox Delivery  (Set-MailboxTransportService -MailboxDeliveryAgentLogPath / -PipelineTracingPath)
        - IIS - Default Web Site, Exchange Back End (via WebAdministration remoting)
        - POP3 / IMAP4      (Set-PopSettings / Set-ImapSettings)

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

    Write-StepBanner -StepName '06' -Title 'Log paths reconfiguration' -Mode $Mode -Actions @(
        'Frontend Transport: ConnectivityLogPath, ReceiveProtocolLogPath, SendProtocolLogPath, AgentLogPath',
        'Hub Transport: ConnectivityLogPath, ReceiveProtocolLogPath, SendProtocolLogPath, MessageTrackingLogPath, RoutingTableLogPath, QueueLogPath, IrmLogPath',
        'Mailbox Delivery: MailboxDeliveryAgentLogPath, PipelineTracingPath',
        'IIS Default Web Site and Exchange Back End: redirection to the defined paths',
        'POP3 / IMAP4 logs',
        'Restart of the affected services after the change (Apply mode only)'
    )

    Initialize-ExchangeShell

    $servers = @(Get-Exchange2019Servers)
    $L = $Config.LogPaths

    foreach ($srv in $servers) {
        $server = $srv.Name

        # --- Frontend Transport ---------------------------------------------
        Invoke-Action -Step '06-LogPaths' -Target $server -Action 'Set-FrontendTransportService' -Detail "Logs FE -> $($L.FrontEndTransportLogs)" `
            -InventoryScript {
                $s = Get-FrontendTransportService $server -ErrorAction Stop
                "Conn={0}; Recv={1}; Send={2}; Agent={3}" -f $s.ConnectivityLogPath, $s.ReceiveProtocolLogPath, $s.SendProtocolLogPath, $s.AgentLogPath
            } `
            -PreCheckScript {
                $s = Get-FrontendTransportService $server -ErrorAction Stop
                return (([string]$s.ConnectivityLogPath).StartsWith($L.FrontEndTransportLogs, 'OrdinalIgnoreCase'))
            } `
            -ActionScript {
                $base = $L.FrontEndTransportLogs
                Set-FrontendTransportService -Identity $server `
                    -ConnectivityLogPath        (Join-Path $base 'Connectivity') `
                    -ReceiveProtocolLogPath     (Join-Path $base 'ReceiveProtocol') `
                    -SendProtocolLogPath        (Join-Path $base 'SendProtocol') `
                    -AgentLogPath               (Join-Path $base 'AgentLog') `
                    -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
            }

        # --- Hub Transport --------------------------------------------------
        Invoke-Action -Step '06-LogPaths' -Target $server -Action 'Set-TransportService' -Detail "Logs Hub -> $($L.TransportLogs)" `
            -InventoryScript {
                $s = Get-TransportService $server -ErrorAction Stop
                "Conn={0}; Recv={1}; Send={2}; MsgTrack={3}; Queue={4}; Routing={5}; Irm={6}" -f `
                    $s.ConnectivityLogPath, $s.ReceiveProtocolLogPath, $s.SendProtocolLogPath, $s.MessageTrackingLogPath, $s.QueueLogPath, $s.RoutingTableLogPath, $s.IrmLogPath
            } `
            -PreCheckScript {
                $s = Get-TransportService $server -ErrorAction Stop
                return (([string]$s.ConnectivityLogPath).StartsWith($L.TransportLogs, 'OrdinalIgnoreCase'))
            } `
            -ActionScript {
                $base = $L.TransportLogs
                Set-TransportService -Identity $server `
                    -ConnectivityLogPath        (Join-Path $base 'Connectivity') `
                    -ReceiveProtocolLogPath     (Join-Path $base 'ReceiveProtocol') `
                    -SendProtocolLogPath        (Join-Path $base 'SendProtocol') `
                    -MessageTrackingLogPath     (Join-Path $base 'MessageTracking') `
                    -QueueLogPath               (Join-Path $base 'Queue') `
                    -RoutingTableLogPath        (Join-Path $base 'Routing') `
                    -IrmLogPath                 (Join-Path $base 'Irm') `
                    -PipelineTracingPath        (Join-Path $base 'PipelineTracing') `
                    -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
            }

        # --- Mailbox Delivery -----------------------------------------------
        Invoke-Action -Step '06-LogPaths' -Target $server -Action 'Set-MailboxTransportService' -Detail "Logs MboxDelivery -> $($L.MailboxDeliveryLogs)" `
            -InventoryScript {
                $s = Get-MailboxTransportService $server -ErrorAction Stop
                "Delivery={0}; PipelineTracing={1}" -f $s.MailboxDeliveryAgentLogPath, $s.PipelineTracingPath
            } `
            -PreCheckScript {
                $s = Get-MailboxTransportService $server -ErrorAction Stop
                return (([string]$s.MailboxDeliveryAgentLogPath).StartsWith($L.MailboxDeliveryLogs, 'OrdinalIgnoreCase'))
            } `
            -ActionScript {
                # Set-MailboxTransportService on Exchange 2019: there is no
                # MailboxDeliveryConnectorProtocolLogPath. The Protocol Receive/Send logs
                # can be set through ReceiveProtocolLogPath / SendProtocolLogPath.
                $base = $L.MailboxDeliveryLogs
                Set-MailboxTransportService -Identity $server `
                    -MailboxDeliveryAgentLogPath (Join-Path $base 'AgentLog') `
                    -ReceiveProtocolLogPath      (Join-Path $base 'ReceiveProtocol') `
                    -SendProtocolLogPath         (Join-Path $base 'SendProtocol') `
                    -PipelineTracingPath         (Join-Path $base 'PipelineTracing') `
                    -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
            }

        # --- POP3 / IMAP4 ---------------------------------------------------
        Invoke-Action -Step '06-LogPaths' -Target $server -Action 'Set-PopSettings logs' -Detail $L.Pop3Logs `
            -InventoryScript {
                $s = Get-PopSettings -Server $server -ErrorAction Stop
                "LogDir={0}" -f $s.LogFileLocation
            } `
            -PreCheckScript {
                $s = Get-PopSettings -Server $server -ErrorAction Stop
                return (([string]$s.LogFileLocation).StartsWith($L.Pop3Logs, 'OrdinalIgnoreCase'))
            } `
            -ActionScript {
                Set-PopSettings -Server $server -LogFileLocation $L.Pop3Logs -ProtocolLogEnabled $true `
                    -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
            }

        # Restart the POP3 services to apply the changes (Microsoft warning).
        # Logic by StartType:
        #   - Disabled  : nothing is touched (deliberate admin decision)
        #   - Manual    : switch to StartType=Automatic, then Restart
        #   - Automatic : Restart (starts the service if Stopped, otherwise stop+start)
        Invoke-Action -Step '06-LogPaths' -Target $server -Action 'Restart POP3 services' `
            -Detail 'MSExchangePop3, MSExchangePop3BE: Disabled=skip, Manual->Auto+Restart, Auto=Restart' `
            -InventoryScript {
                $lines = foreach ($n in 'MSExchangePop3','MSExchangePop3BE') {
                    $s = Get-Service -ComputerName $server -Name $n -EA SilentlyContinue
                    if ($s) { '{0}={1} (StartType={2})' -f $n, $s.Status, $s.StartType } else { '{0}=<absent>' -f $n }
                }
                $lines -join ' | '
            } `
            -PreCheckScript $null `
            -ActionScript {
                $isWhatIfSvc = ([bool]($Mode -eq 'Simulate'))
                foreach ($n in 'MSExchangePop3','MSExchangePop3BE') {
                    $s = Get-Service -ComputerName $server -Name $n -EA SilentlyContinue
                    if (-not $s) { Write-Log "${server}: ${n} not found" -Level Warning; continue }
                    $startType = [string]$s.StartType

                    if ($startType -eq 'Disabled') {
                        Write-Log "${server}: ${n} = Disabled - skipped (deliberate decision)" -Level Sub
                        continue
                    }
                    if ($isWhatIfSvc) {
                        $todo = if ($startType -eq 'Manual') { 'Auto + Restart' } else { 'Restart' }
                        Write-Log "WhatIf: ${n} on ${server} (StartType=${startType}, Status=$($s.Status)) -> $todo" -Level Sub
                        continue
                    }
                    # Set-Service in PS 5.1 has no -ComputerName: Invoke-Command is required.
                    Invoke-Command -ComputerName $server -ScriptBlock {
                        param($name, $switchToAuto)
                        $ErrorActionPreference = 'Stop'
                        if ($switchToAuto) { Set-Service -Name $name -StartupType Automatic }
                        Restart-Service -Name $name -Force
                    } -ArgumentList $n, ($startType -eq 'Manual') -ErrorAction Stop
                    if ($startType -eq 'Manual') {
                        Write-Log "${server}: ${n}: StartType Manual -> Automatic + Restart" -Level Sub
                    } else {
                        Write-Log "${server}: ${n}: Restart (StartType=${startType})" -Level Sub
                    }
                }
            }

        Invoke-Action -Step '06-LogPaths' -Target $server -Action 'Set-ImapSettings logs' -Detail $L.Imap4Logs `
            -InventoryScript {
                $s = Get-ImapSettings -Server $server -ErrorAction Stop
                "LogDir={0}" -f $s.LogFileLocation
            } `
            -PreCheckScript {
                $s = Get-ImapSettings -Server $server -ErrorAction Stop
                return (([string]$s.LogFileLocation).StartsWith($L.Imap4Logs, 'OrdinalIgnoreCase'))
            } `
            -ActionScript {
                Set-ImapSettings -Server $server -LogFileLocation $L.Imap4Logs -ProtocolLogEnabled $true `
                    -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
            }

        # Restart the IMAP4 services - same tri-state logic as POP3.
        Invoke-Action -Step '06-LogPaths' -Target $server -Action 'Restart IMAP4 services' `
            -Detail 'MSExchangeImap4, MSExchangeImap4BE: Disabled=skip, Manual->Auto+Restart, Auto=Restart' `
            -InventoryScript {
                $lines = foreach ($n in 'MSExchangeImap4','MSExchangeImap4BE') {
                    $s = Get-Service -ComputerName $server -Name $n -EA SilentlyContinue
                    if ($s) { '{0}={1} (StartType={2})' -f $n, $s.Status, $s.StartType } else { '{0}=<absent>' -f $n }
                }
                $lines -join ' | '
            } `
            -PreCheckScript $null `
            -ActionScript {
                $isWhatIfSvc = ([bool]($Mode -eq 'Simulate'))
                foreach ($n in 'MSExchangeImap4','MSExchangeImap4BE') {
                    $s = Get-Service -ComputerName $server -Name $n -EA SilentlyContinue
                    if (-not $s) { Write-Log "${server}: ${n} not found" -Level Warning; continue }
                    $startType = [string]$s.StartType

                    if ($startType -eq 'Disabled') {
                        Write-Log "${server}: ${n} = Disabled - skipped (deliberate decision)" -Level Sub
                        continue
                    }
                    if ($isWhatIfSvc) {
                        $todo = if ($startType -eq 'Manual') { 'Auto + Restart' } else { 'Restart' }
                        Write-Log "WhatIf: ${n} on ${server} (StartType=${startType}, Status=$($s.Status)) -> $todo" -Level Sub
                        continue
                    }
                    Invoke-Command -ComputerName $server -ScriptBlock {
                        param($name, $switchToAuto)
                        $ErrorActionPreference = 'Stop'
                        if ($switchToAuto) { Set-Service -Name $name -StartupType Automatic }
                        Restart-Service -Name $name -Force
                    } -ArgumentList $n, ($startType -eq 'Manual') -ErrorAction Stop
                    if ($startType -eq 'Manual') {
                        Write-Log "${server}: ${n}: StartType Manual -> Automatic + Restart" -Level Sub
                    } else {
                        Write-Log "${server}: ${n}: Restart (StartType=${startType})" -Level Sub
                    }
                }
            }

        # --- IIS (DefaultWebSite + ExchangeBackEnd) -------------------------
        # The system.applicationHost/sites section is locked at the apphost
        # level: Set-WebConfigurationProperty cannot be used at site level.
        # Set-ItemProperty is used on the site in the IIS: drive
        # (which writes directly into applicationHost.config) -- method
        # supported by WebAdministration.
        $isWhatIfIIS = ([bool]($Mode -eq 'Simulate'))
        Invoke-Action -Step '06-LogPaths' -Target $server -Action 'IIS log directories' -Detail "IIS -> $($L.IISLogsDefaultWebSite) / $($L.IISLogsBackEnd)" `
            -InventoryScript {
                Invoke-Command -ComputerName $server -ScriptBlock {
                    Import-Module WebAdministration -ErrorAction Stop
                    $sites = Get-WebSite | Where-Object { $_.Name -in 'Default Web Site','Exchange Back End' }
                    ($sites | ForEach-Object { "{0}: directory={1}" -f $_.Name, $_.LogFile.directory }) -join ' | '
                } -ErrorAction Stop
            } `
            -PreCheckScript {
                $current = Invoke-Command -ComputerName $server -ScriptBlock {
                    Import-Module WebAdministration -ErrorAction Stop
                    @{
                        Default = (Get-ItemProperty 'IIS:\Sites\Default Web Site' -Name logFile.directory).Value
                        Back    = (Get-ItemProperty 'IIS:\Sites\Exchange Back End' -Name logFile.directory).Value
                    }
                } -ErrorAction Stop
                return ($current.Default -ieq $L.IISLogsDefaultWebSite -and $current.Back -ieq $L.IISLogsBackEnd)
            } `
            -ActionScript {
                Invoke-Command -ComputerName $server -ScriptBlock {
                    param($default, $back, $whatIf)
                    Import-Module WebAdministration -ErrorAction Stop
                    foreach ($p in @($default, $back)) {
                        if (-not (Test-Path $p)) {
                            if ($whatIf) {
                                Write-Output "WhatIf: New-Item -Path $p"
                            } else {
                                New-Item -ItemType Directory -Path $p -Force | Out-Null
                            }
                        }
                    }
                    if ($whatIf) {
                        Write-Output "WhatIf: Set logFile.directory 'Default Web Site' = $default"
                        Write-Output "WhatIf: Set logFile.directory 'Exchange Back End' = $back"
                    } else {
                        Set-ItemProperty 'IIS:\Sites\Default Web Site' -Name logFile.directory -Value $default
                        Set-ItemProperty 'IIS:\Sites\Exchange Back End' -Name logFile.directory -Value $back
                    }
                } -ArgumentList $L.IISLogsDefaultWebSite, $L.IISLogsBackEnd, $isWhatIfIIS -ErrorAction Stop
            }
    }

    return 0
}
