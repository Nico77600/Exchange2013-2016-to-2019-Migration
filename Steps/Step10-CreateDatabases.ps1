<#
.SYNOPSIS
    Step 10 - Create the mailbox databases.

.DESCRIPTION
    Creates the mailbox databases.

    The databases are generated automatically from DatabaseLayout (Deployment.config.psd1):
      DatabasePrefix / DatabaseDigits / DatabaseStartIndex / DatabaseCount
      -> e.g. DB01, DB02, DB03, DB04
    EDB and Logs paths are resolved from EdbPathTemplate / LogPathTemplate.
    The active server is assigned in round-robin (detected Exchange 2019 servers or DatabaseLayout.Servers).
    Properties are applied from DatabaseLayout.DefaultProperties.
    The generated plan is exported to Reports\DatabasePlan.csv.

    Sequence:
      1) Check that the target server is an Exchange 2019 server
      2) Create the EDB and Logs folders remotely (Invoke-Command) before New-MailboxDatabase
      3) New-MailboxDatabase
      4) Set-MailboxDatabase (DefaultProperties)
      5) Mount-Database with retry (Apply only)

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

    Initialize-ExchangeShell -Credential $Credential

    # ==========================================================================
    # 1) Building the database list
    # ==========================================================================
    $layout      = $Config.DatabaseLayout
    $prefix      = $layout.DatabasePrefix
    $digits      = [int]$layout.DatabaseDigits
    $startIdx    = [int]$layout.DatabaseStartIndex
    $count       = [int]$layout.DatabaseCount
    $driveLetter = ([string]$layout.DriveLetter).TrimEnd(':').Trim()
    $edbTpl      = $layout.EdbPathTemplate
    $logTpl      = $layout.LogPathTemplate
    $explicitSrv = @($layout.Servers | Where-Object { $_ })
    $defProps    = $layout.DefaultProperties

    $servers2019   = (Get-Exchange2019Servers).Name | Sort-Object
    $targetServers = if ($explicitSrv) { $explicitSrv } else { $servers2019 }

    $padFmt    = '{0:D' + $digits + '}'
    $firstName = "$prefix" + ($padFmt -f $startIdx)
    $lastName  = "$prefix" + ($padFmt -f ($startIdx + $count - 1))

    $previewEdb = $edbTpl -replace '\{Drive\}', $driveLetter -replace '\{Name\}', $firstName
    $previewLog = $logTpl -replace '\{Drive\}', $driveLetter -replace '\{Name\}', $firstName

    Write-StepBanner -StepName '10' -Title 'Mailbox database creation' -Mode $Mode -Actions @(
        "Generated names: $firstName ... $lastName ($count database(s))",
        "EDB path preview  ($firstName): $previewEdb",
        "Logs path preview ($firstName): $previewLog",
        "Target servers (round-robin): $($targetServers -join ', ')",
        "Check that each target server is an Exchange 2019 server",
        "Create the EDB and Logs folders REMOTELY before New-MailboxDatabase",
        "Set-MailboxDatabase with DefaultProperties ($($defProps.Keys.Count) properties)",
        "Mount-Database with retry (Apply only)",
        "Cleanup of the default 'Mailbox Database NNN' databases (installed by default):",
        "  - Remove the HealthMailbox/Monitoring mailboxes + AD object (recreated automatically by Exchange)",
        "  - If user/arbitration/audit mailboxes remain: WARN + admin prompt (move required)",
        "  - If empty: Dismount + Remove-MailboxDatabase + cleanup of .edb and Logs on disk",
        "Restart-Service MSExchangeIS on each target server (Apply only)"
    )

    $dbs = @()
    for ($i = 0; $i -lt $count; $i++) {
        $dbIdx  = $startIdx + $i
        $name   = "$prefix" + ($padFmt -f $dbIdx)
        $srv    = $targetServers[$i % $targetServers.Count]

        $edbPath = $edbTpl -replace '\{Drive\}', $driveLetter -replace '\{Name\}', $name
        $logPath = $logTpl -replace '\{Drive\}', $driveLetter -replace '\{Name\}', $name

        $dbs += [PSCustomObject]@{
            Name          = $name
            Server        = $srv
            EDBFilePath   = $edbPath
            LogFolderPath = $logPath
            GC            = ($Config.PreferredGCs | Select-Object -First 1)
        }
    }

    # Export the plan to Reports\
    $planPath = Join-Path $OutputFolder 'DatabasePlan.csv'
    $dbs | Select-Object Name, Server, EDBFilePath, LogFolderPath, GC |
        Export-Csv -Path $planPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Log "Creation plan exported: $planPath" -Level Sub

    # ==========================================================================
    # 2) ScriptBlocks for the remote folder operations
    # ==========================================================================
    $checkFoldersSb = {
        param([string]$EdbDir, [string]$LogDir)
        @{
            EdbDirExists = (Test-Path $EdbDir)
            LogDirExists = (Test-Path $LogDir)
        }
    }
    $createFoldersSb = {
        param([string]$EdbDir, [string]$LogDir)
        $ErrorActionPreference = 'Stop'
        $created = @()
        foreach ($d in @($EdbDir, $LogDir) | Where-Object { $_ } | Sort-Object -Unique) {
            if (-not (Test-Path $d)) {
                New-Item -ItemType Directory -Path $d -Force | Out-Null
                $created += $d
            }
        }
        @{
            EdbDirExists = (Test-Path $EdbDir)
            LogDirExists = (Test-Path $LogDir)
            Created      = $created
        }
    }

    # ==========================================================================
    # 3) Iterating over the databases
    # ==========================================================================
    foreach ($db in $dbs) {
        $dbName  = $db.Name
        $server  = $db.Server
        $edbPath = $db.EDBFilePath
        $logPath = $db.LogFolderPath
        $gc      = $db.GC

        if (-not (Test-IsExchange2019Server -ServerName $server)) {
            Add-Report -Step '10-DBs' -Target $dbName -Action 'New-MailboxDatabase' -Status Skipped `
                       -Detail "Server $server is not an Exchange 2019 server"
            continue
        }

        $edbDir = Split-Path $edbPath -Parent
        $logDir = $logPath

        # FQDN or short name for Invoke-Command
        $remoteHost = (Get-Exchange2019Servers | Where-Object { $_.Name -eq $server } | Select-Object -First 1).Fqdn
        if (-not $remoteHost) { $remoteHost = $server }

        $invokeParams = @{
            ComputerName = $remoteHost
            ErrorAction  = 'Stop'
        }
        if ($Credential) { $invokeParams['Credential'] = $Credential }

        # In Simulate, the target drive (M:) may not exist (Step04 -WhatIf)
        # and/or the database may not exist (New-MailboxDatabase -WhatIf does not
        # create it). Any downstream action (Set-MailboxDatabase) then fails with
        # an "object not found". We detect these cases and report Simulated.
        $skipDb = $false
        $simulateSetDb = $false
        if ($Mode -eq 'Simulate') {
            $driveExists = $false
            try {
                $driveExists = Invoke-Command @invokeParams -ScriptBlock {
                    param($d) (Test-Path "${d}:\")
                } -ArgumentList $driveLetter
            } catch { $driveExists = $false }

            $dbExistsNow = [bool](Get-MailboxDatabase -Identity $dbName -ErrorAction SilentlyContinue)

            if (-not $driveExists) {
                # Drive missing: no action can be performed, even in Apply
                Add-Report -Step '10-DBs' -Target "$server : $dbName folders" -Action 'Create EDB+Log folders (remote)' `
                           -Status Simulated -BeforeValue "Drive ${driveLetter}: absent" `
                           -AfterValue '(simulation - not applied)' `
                           -Detail "EDB=$edbDir | Logs=$logDir (drive will be created by Step04 Apply)"
                Add-Report -Step '10-DBs' -Target $dbName -Action 'New-MailboxDatabase' `
                           -Status Simulated -BeforeValue '<database absent, target drive absent>' `
                           -AfterValue '(simulation - not applied)' -Detail "Server=$server"
                Add-Report -Step '10-DBs' -Target $dbName -Action 'Set-MailboxDatabase (DefaultProperties)' `
                           -Status Simulated -BeforeValue '<database absent>' `
                           -AfterValue '(simulation - not applied)' -Detail "$($defProps.Count) properties"
                Write-Log "[Simulate] Drive ${driveLetter}: absent on $server - DB '$dbName' actions deferred to the next Apply." -Level Sub
                $skipDb = $true
            }
            elseif (-not $dbExistsNow) {
                # Drive OK but DB missing -> folder + New-DB can be simulated normally,
                # but the Set that follows would fail (DB still non-existent). We only
                # pre-simulate Set-MailboxDatabase.
                $simulateSetDb = $true
            }
        }
        if ($skipDb) { continue }

        # ----- 3.a) Creating the folders UPSTREAM (remotely) -----
        Invoke-Action -Step '10-DBs' -Target "$server : $dbName folders" -Action 'Create EDB+Log folders (remote)' `
            -Detail "EDB=$edbDir | Logs=$logDir" `
            -InventoryScript {
                $info = Invoke-Command @invokeParams -ScriptBlock $checkFoldersSb -ArgumentList $edbDir, $logDir
                "EDB exists=$($info.EdbDirExists) | Logs exists=$($info.LogDirExists)"
            } `
            -PreCheckScript {
                $info = Invoke-Command @invokeParams -ScriptBlock $checkFoldersSb -ArgumentList $edbDir, $logDir
                return ($info.EdbDirExists -and $info.LogDirExists)
            } `
            -ActionScript {
                $r = Invoke-Command @invokeParams -ScriptBlock $createFoldersSb -ArgumentList $edbDir, $logDir
                if ($r.Created.Count -gt 0) {
                    Write-Log "  Created: $($r.Created -join ', ')" -Level Sub
                }
                if (-not $r.EdbDirExists) { throw "Creation failed: $edbDir" }
                if (-not $r.LogDirExists) { throw "Creation failed: $logDir" }
            }

        # ----- 3.b) New-MailboxDatabase -----
        Invoke-Action -Step '10-DBs' -Target $dbName -Action 'New-MailboxDatabase' -Detail "Server=$server" `
            -InventoryScript {
                $existing = Get-MailboxDatabase -Identity $dbName -ErrorAction SilentlyContinue
                if ($existing) {
                    "Server={0}; Edb={1}; Log={2}" -f $existing.Server, $existing.EdbFilePath, $existing.LogFolderPath
                } else { '<absent>' }
            } `
            -PreCheckScript {
                $existing = Get-MailboxDatabase -Identity $dbName -ErrorAction SilentlyContinue
                return [bool]$existing
            } `
            -ActionScript {
                $params = @{
                    Name           = $dbName
                    Server         = $server
                    EdbFilePath    = $edbPath
                    LogFolderPath  = $logPath
                    ErrorAction    = 'Stop'
                }
                if ($gc) { $params['DomainController'] = $gc }
                New-MailboxDatabase @params -WhatIf:([bool]($Mode -eq 'Simulate')) | Out-Null
            }

        # ----- 3.c) Set-MailboxDatabase (DefaultProperties) -----
        $setProps = @{}
        foreach ($k in $defProps.Keys) {
            if ($null -ne $defProps[$k]) { $setProps[$k] = $defProps[$k] }
        }
        if ($gc) { $setProps['DomainController'] = $gc }
        $setProps['ErrorAction'] = 'Stop'

        if ($simulateSetDb) {
            Add-Report -Step '10-DBs' -Target $dbName -Action 'Set-MailboxDatabase (DefaultProperties)' `
                       -Status Simulated -BeforeValue '<database absent - simulated by New>' `
                       -AfterValue '(simulation - not applied)' -Detail "$($setProps.Count) properties"
        } else {
            Invoke-Action -Step '10-DBs' -Target $dbName -Action 'Set-MailboxDatabase (DefaultProperties)' `
                -Detail "$($setProps.Count) properties" `
                -InventoryScript {
                    $b = Get-MailboxDatabase -Identity $dbName -Status -ErrorAction SilentlyContinue
                    if ($b) {
                        "Send={0}; SendRecv={1}; Warn={2}; Index={3}; Mounted={4}" -f $b.ProhibitSendQuota, $b.ProhibitSendReceiveQuota, $b.IssueWarningQuota, $b.IndexEnabled, $b.Mounted
                    } else { '<absent>' }
                } `
                -PreCheckScript $null `
                -ActionScript {
                    Set-MailboxDatabase -Identity $dbName @setProps -WhatIf:([bool]($Mode -eq 'Simulate'))
                }
        }

        # ----- 3.d) Mount (Apply only) -----
        if ($Mode -eq 'Apply') {
            Invoke-Action -Step '10-DBs' -Target $dbName -Action 'Mount-Database' -Detail "GC=$gc" `
                -InventoryScript {
                    $b = Get-MailboxDatabase -Identity $dbName -Status -ErrorAction SilentlyContinue
                    if ($b) { "Mounted=$($b.Mounted)" } else { '<database absent>' }
                } `
                -PreCheckScript {
                    $b = Get-MailboxDatabase -Identity $dbName -Status -ErrorAction Stop
                    return [bool]$b.Mounted
                } `
                -ActionScript {
                    $maxTry = 5
                    for ($i = 1; $i -le $maxTry; $i++) {
                        try {
                            if ($gc) {
                                Mount-Database -Identity $dbName -DomainController $gc -ErrorAction Stop
                            } else {
                                Mount-Database -Identity $dbName -ErrorAction Stop
                            }
                            break
                        } catch {
                            if ($i -eq $maxTry) { throw }
                            Start-Sleep -Seconds 60
                        }
                    }
                }
        }
    }

    # ==========================================================================
    # 4) Cleanup of the default "Mailbox Database NNN" databases
    # ==========================================================================
    # Strict pattern: name matching 'Mailbox Database \d+' (created by the default
    # Exchange 2019 installation). We ONLY touch databases hosted on the detected
    # Exchange 2019 servers (never on 2013/2016).
    # Sequence per database:
    #   a) Remove the Health/Monitoring mailboxes (recreated automatically by Exchange)
    #   b) Inventory of the remaining mailboxes (user + arbitration + audit)
    #   c) If not empty in Apply: colored WARN + YES/NO prompt to confirm
    #      that the admin has moved these mailboxes
    #   d) If empty: Dismount + Remove-MailboxDatabase + cleanup of .edb and Logs

    $defaultDbPattern = '^Mailbox Database \d+$'
    # Note: through Exchange remoting, $_.Server is serialized as a [string] (the
    # ToString of the ProvisioningServer), not as an object with .Name. We compare
    # the string directly. If the cmdlet were called locally, $_.Server.Name would be correct.
    $defaultDbs = @(Get-MailboxDatabase -ErrorAction SilentlyContinue |
                     Where-Object {
                         $_.Name -match $defaultDbPattern -and
                         ([string]$_.Server) -in $servers2019
                     })

    if (-not $defaultDbs -or $defaultDbs.Count -eq 0) {
        Write-Log "Cleanup: no 'Mailbox Database NNN' database on the Exchange 2019 servers - nothing to clean up." -Level Sub
    }

    foreach ($defDb in $defaultDbs) {
        $defDbName = $defDb.Name
        $defDbSrv  = [string]$defDb.Server

        # ----- 4.a) Remove Health/Monitoring Mailboxes -----------------------
        # Remove-Mailbox -Monitoring (vs Disable-Mailbox) also removes the associated
        # AD object in 'Microsoft Exchange System Objects'. The HealthMailbox
        # mailboxes are recreated automatically by the Exchange Health Manager Service
        # at its next iteration (a few minutes), so removing the AD object is safe.
        Invoke-Action -Step '10-DBs' -Target $defDbName -Action 'Cleanup HealthMailboxes' `
            -Detail "Remove the Monitoring mailboxes (HealthMailbox-*) + AD object - recreated automatically by Exchange" `
            -InventoryScript {
                try {
                    $hm = @(Get-Mailbox -Monitoring -ResultSize Unlimited -EA SilentlyContinue |
                            Where-Object { $_.Database -ieq $defDbName })
                    "Monitoring mailboxes count=$($hm.Count)"
                } catch { "Error: $($_.Exception.Message)" }
            } `
            -PreCheckScript {
                $hm = @(Get-Mailbox -Monitoring -ResultSize Unlimited -EA SilentlyContinue |
                        Where-Object { $_.Database -ieq $defDbName })
                return ($hm.Count -eq 0)
            } `
            -ActionScript {
                $hm = @(Get-Mailbox -Monitoring -ResultSize Unlimited -EA SilentlyContinue |
                        Where-Object { $_.Database -ieq $defDbName })
                foreach ($mbx in $hm) {
                    Remove-Mailbox -Identity $mbx.Identity -Monitoring -Confirm:$false `
                        -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction SilentlyContinue
                }
                if ($Mode -eq 'Apply') {
                    Write-Log "$defDbName - $($hm.Count) Monitoring mailbox(es) removed (AD cleaned up)" -Level Sub
                }
            }

        # In Inventory: we report the state of the remaining mailboxes (without touching anything).
        # In Simulate/Apply: we continue the flow (the removal follows, conditionally).
        if ($Mode -eq 'Inventory') {
            $allMbx = @()
            try {
                $allMbx += @(Get-Mailbox -Database $defDbName -ResultSize Unlimited -EA SilentlyContinue)
                $allMbx += @(Get-Mailbox -Arbitration -ResultSize Unlimited -EA SilentlyContinue |
                             Where-Object { $_.Database -ieq $defDbName })
                $allMbx += @(Get-Mailbox -AuditLog -ResultSize Unlimited -EA SilentlyContinue |
                             Where-Object { $_.Database -ieq $defDbName })
                $allMbx = @($allMbx | Sort-Object Name -Unique)
            } catch { }
            $mbxSummary = if ($allMbx.Count -eq 0) {
                'no remaining non-Monitoring mailbox'
            } else {
                $sample = ($allMbx | Select-Object -First 5 | ForEach-Object { "$($_.Name)[$($_.RecipientTypeDetails)]" }) -join ', '
                "$($allMbx.Count) non-Monitoring mailbox(es): $sample$(if ($allMbx.Count -gt 5) {' ...'})"
            }
            Add-Report -Step '10-DBs' -Target $defDbName -Action 'Remove-MailboxDatabase (default)' -Status Inventoried `
                       -BeforeValue ("Server=$defDbSrv; $mbxSummary") -Detail 'Default database to clean up after Step18/20'
            continue
        }

        # ----- 4.b) Inventory of the remaining mailboxes (after the Health cleanup) -
        $remainingMbx = @()
        try {
            $remainingMbx += @(Get-Mailbox -Database $defDbName -ResultSize Unlimited -EA SilentlyContinue)
            $remainingMbx += @(Get-Mailbox -Arbitration -ResultSize Unlimited -EA SilentlyContinue |
                               Where-Object { $_.Database -ieq $defDbName })
            $remainingMbx += @(Get-Mailbox -AuditLog -ResultSize Unlimited -EA SilentlyContinue |
                               Where-Object { $_.Database -ieq $defDbName })
            $remainingMbx = @($remainingMbx | Sort-Object Name -Unique)
        } catch {
            Write-Log "Error listing mailboxes of ${defDbName}: $($_.Exception.Message)" -Level Warning
        }

        # ----- 4.c) WARN + admin prompt if mailboxes remain ------------------
        if ($remainingMbx.Count -gt 0) {
            $mbxList = ($remainingMbx | Select-Object -First 20 |
                        ForEach-Object { "$($_.Name) [$($_.RecipientTypeDetails)]" }) -join ', '

            Write-Host ''
            Write-ExItem Warn ("REMAINING MAILBOXES on '$defDbName' ($defDbSrv): $($remainingMbx.Count)")
            foreach ($mbx in ($remainingMbx | Select-Object -First 50)) {
                Write-ExItem Dim ("{0,-50} [{1}]" -f $mbx.Name, $mbx.RecipientTypeDetails) -Indent 9
            }
            if ($remainingMbx.Count -gt 50) {
                Write-ExItem Dim ("... and $($remainingMbx.Count - 50) more") -Indent 9
            }
            Write-Host ''
            Write-ExItem Warn "Move these mailboxes to another database BEFORE deleting '$defDbName'."
            Write-ExItem Dim "System mailboxes (arbitration/audit): New-MoveRequest (see Step18 SystemMailboxMigration)" -Indent 9
            Write-ExItem Dim "User mailboxes:                       New-MoveRequest (see Step19/20 Prepare/RunMigration)" -Indent 9
            Write-Host ''

            $skipDelete = $true
            if ($Mode -eq 'Apply' -and -not [Console]::IsInputRedirected -and $Host.Name -match 'ConsoleHost') {
                $resp = Read-Host "     Have you moved these mailboxes? YES to re-check and delete / NO to skip"
                if ($resp -ieq 'YES') {
                    # Post-move re-check
                    $remainingMbx = @()
                    try {
                        $remainingMbx += @(Get-Mailbox -Database $defDbName -ResultSize Unlimited -EA SilentlyContinue)
                        $remainingMbx += @(Get-Mailbox -Arbitration -ResultSize Unlimited -EA SilentlyContinue |
                                           Where-Object { $_.Database -ieq $defDbName })
                        $remainingMbx += @(Get-Mailbox -AuditLog -ResultSize Unlimited -EA SilentlyContinue |
                                           Where-Object { $_.Database -ieq $defDbName })
                        $remainingMbx = @($remainingMbx | Sort-Object Name -Unique)
                    } catch { }

                    if ($remainingMbx.Count -eq 0) {
                        Write-Log "$defDbName - mailboxes moved, deletion possible." -Level Sub
                        $skipDelete = $false
                    } else {
                        Write-Log "$defDbName - $($remainingMbx.Count) mailbox(es) still remaining - skipped." -Level Warning
                    }
                } else {
                    Write-Log "$defDbName - admin chose NOT to delete now (response=$resp)" -Level Sub
                }
            }

            if ($skipDelete) {
                Add-Report -Step '10-DBs' -Target $defDbName -Action 'Remove-MailboxDatabase (default)' -Status Skipped `
                           -Detail "$($remainingMbx.Count) remaining non-Monitoring mailbox(es). Move required. Sample: $mbxList"
                continue
            }
        }

        # ----- 4.d) Actual removal + disk cleanup ----------------------------
        Invoke-Action -Step '10-DBs' -Target $defDbName -Action 'Remove-MailboxDatabase (default)' `
            -Detail "Default DB '$defDbName' on $defDbSrv - empty -> Dismount + Remove + file cleanup" `
            -InventoryScript {
                $b = Get-MailboxDatabase -Identity $defDbName -Status -ErrorAction SilentlyContinue
                if ($b) { "Mounted=$($b.Mounted); EDB=$($b.EdbFilePath); Logs=$($b.LogFolderPath)" } else { '<absent>' }
            } `
            -PreCheckScript {
                return -not [bool](Get-MailboxDatabase -Identity $defDbName -ErrorAction SilentlyContinue)
            } `
            -ActionScript {
                $existing = Get-MailboxDatabase -Identity $defDbName -Status -ErrorAction SilentlyContinue
                if (-not $existing) { return }

                $edbFile   = [string]$existing.EdbFilePath
                $logFolder = [string]$existing.LogFolderPath

                if ($existing.Mounted) {
                    Dismount-Database -Identity $defDbName -Confirm:$false `
                        -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
                }

                Remove-MailboxDatabase -Identity $defDbName -Confirm:$false `
                    -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop

                # Disk cleanup (Apply only) - Remove-MailboxDatabase does NOT delete
                # the .edb / .log / .chk files on the target server's filesystem.
                #
                # We delete the folders containing the files (parent of the .edb + LogFolderPath).
                # For the Exchange default databases, these two paths point to the SAME
                # "Mailbox Database NNN" folder, so a single recursive deletion is enough.
                # Safeguard: we only delete folders whose name contains the database name,
                # which rules out any risk on system parent folders.
                # Retry x3 (sleep 5s) to absorb residual MSExchangeIS locks after
                # Dismount + Remove-MailboxDatabase.
                if ($Mode -eq 'Apply') {
                    try {
                        $fqdn = (Get-Exchange2019Servers | Where-Object { $_.Name -eq $defDbSrv } | Select-Object -First 1).Fqdn
                        if (-not $fqdn) { $fqdn = $defDbSrv }
                        $cleanupResult = Invoke-Command -ComputerName $fqdn -ScriptBlock {
                            param($dbName, $edbFile, $logFolder)
                            $removed = @()
                            $kept    = @()

                            $folders = @()
                            if ($edbFile)   { $folders += (Split-Path $edbFile -Parent) }
                            if ($logFolder) { $folders += $logFolder }
                            $folders = @($folders | Where-Object { $_ } | Select-Object -Unique)

                            foreach ($f in $folders) {
                                if (-not (Test-Path $f)) { continue }

                                # Safeguard: the folder name MUST contain the database name.
                                # Protects against abnormal paths that would point to a
                                # system parent folder (e.g. "C:\Program Files\...\Mailbox\").
                                $leaf = Split-Path $f -Leaf
                                if ($leaf -notmatch ([regex]::Escape($dbName))) {
                                    $kept += "$f (folder name '$leaf' does not contain '$dbName' - safety skip)"
                                    continue
                                }

                                # Retry to absorb possible MSExchangeIS locks
                                $ok = $false
                                for ($i = 1; $i -le 3 -and -not $ok; $i++) {
                                    try {
                                        Remove-Item -Path $f -Recurse -Force -ErrorAction Stop
                                        $ok = $true
                                        $removed += $f
                                    } catch {
                                        if ($i -eq 3) {
                                            $kept += "$f (after 3 attempts: $($_.Exception.Message))"
                                        } else {
                                            Start-Sleep -Seconds 5
                                        }
                                    }
                                }
                            }
                            [PSCustomObject]@{ Removed = $removed; Kept = $kept }
                        } -ArgumentList $defDbName, $edbFile, $logFolder -ErrorAction Stop

                        if ($cleanupResult.Removed.Count -gt 0) {
                            Write-Log "$defDbName - folder(s) removed: $($cleanupResult.Removed -join ' | ')" -Level Sub
                        }
                        if ($cleanupResult.Kept.Count -gt 0) {
                            Write-Log "$defDbName - folder(s) NOT removed: $($cleanupResult.Kept -join ' | ')" -Level Warning
                        }
                        if ($cleanupResult.Removed.Count -eq 0 -and $cleanupResult.Kept.Count -eq 0) {
                            Write-Log "$defDbName - no database folder found to remove (already cleaned up?)" -Level Sub
                        }
                    } catch {
                        Write-Log "$defDbName - disk cleanup failed: $($_.Exception.Message)" -Level Warning
                    }
                }
            }
    }

    # ==========================================================================
    # 5) MSExchangeIS restart per server (Apply only)
    # ==========================================================================
    $uniqueServers = $dbs | Select-Object -ExpandProperty Server -Unique
    foreach ($srv in $uniqueServers) {
        $srvFqdn = (Get-Exchange2019Servers | Where-Object { $_.Name -eq $srv } | Select-Object -First 1).Fqdn
        if (-not $srvFqdn) { $srvFqdn = $srv }

        Invoke-Action -Step '10-DBs' -Target $srv -Action 'Restart-Service MSExchangeIS' `
            -Detail 'Restart after database creation' `
            -InventoryScript {
                $s = Get-Service -ComputerName $srvFqdn -Name 'MSExchangeIS' -ErrorAction SilentlyContinue
                if ($s) { "Status=$($s.Status)" } else { '<service not found>' }
            } `
            -PreCheckScript $null `
            -ActionScript {
                $svc = Get-Service -ComputerName $srvFqdn -Name 'MSExchangeIS' -ErrorAction Stop
                if ($svc.Status -eq 'Running') {
                    Restart-Service -InputObject $svc -Force -ErrorAction Stop
                } else {
                    Start-Service -InputObject $svc -ErrorAction Stop
                }
            }
    }

    Save-Report -StepName 'Step10-CreateDatabases'
    return 0
}
