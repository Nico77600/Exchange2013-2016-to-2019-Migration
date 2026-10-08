<#
.SYNOPSIS
    Step 11 - Create the mailbox database copies.

.DESCRIPTION
    Creates the mailbox database copies.

    The copies are generated ALGORITHMICALLY from:
      - The database list (DatabaseLayout config, same logic as Step10)
      - The Site1/Site2 distribution (DAGInfo.csv)

    Fair distribution algorithm (2 sites of N servers, 2N databases):
      For database DB_i whose active server is allServers[i % total]:
        Pref 1: active server           -> mounted copy (defined by New-MailboxDatabase)
        Pref 2: active site, next srv.  -> passive copy
        Pref 3: other site, same idx    -> passive copy
        Pref 4: other site, next idx    -> lagged copy (ReplayLagTime/ReplayLagMaxDelay)

    Guarantees:
      - Each server hosts exactly 1 active + 2 passive + 1 lagged copy
      - Each site hosts the same number of active databases

    The server order is identical to that of Step10:
      DatabaseLayout.Servers if defined, otherwise Get-Exchange2019Servers in alphabetical order.

    No source CSV is needed: the logic is entirely algorithmic.
    The generated plan is exported to Reports\CopyPlan.csv.

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
    # 1) Reading the configuration
    # ==========================================================================
    $layout   = $Config.DatabaseLayout
    $copyConf = $Config.DatabaseCopies
    $gc       = $Config.PreferredGCs | Select-Object -First 1

    $prefix   = $layout.DatabasePrefix
    $digits   = [int]$layout.DatabaseDigits
    $startIdx = [int]$layout.DatabaseStartIndex
    $count    = [int]$layout.DatabaseCount

    $replayLag = $copyConf.ReplayLagTime
    $truncLag  = $copyConf.TruncationLagTime
    $maxDelay  = $copyConf.ReplayLagMaxDelay

    # Round-robin server list (same order as Step10)
    $explicitSrv = @($layout.Servers | Where-Object { $_ })
    $allServers  = if ($explicitSrv) {
        $explicitSrv
    } else {
        @((Get-Exchange2019Servers).Name | Sort-Object)
    }

    # Sites from DAGInfo.csv (used to compute the copy distribution)
    $dagRow = Import-ConfigCsv -Path (Join-Path $CsvFolder 'DAGInfo.csv') | Select-Object -First 1
    $site1  = @($dagRow.Site1Servers -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $site2  = @($dagRow.Site2Servers -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    if (-not $site1 -or -not $site2) {
        Write-Log "Site1Servers or Site2Servers missing in DAGInfo.csv - aborting" -Level Error
        return 1
    }

    # ==========================================================================
    # 2) Generating the copy plan
    # ==========================================================================
    $padFmt = '{0:D' + $digits + '}'
    $plan   = @()

    for ($i = 0; $i -lt $count; $i++) {
        $dbIdx     = $startIdx + $i
        $dbName    = "$prefix" + ($padFmt -f $dbIdx)
        $activeSrv = $allServers[$i % $allServers.Count]

        if ($activeSrv -in $site1) { $actSite = $site1; $othSite = $site2 }
        else                       { $actSite = $site2; $othSite = $site1 }

        $aN   = $actSite.Count
        $oN   = $othSite.Count
        $sIdx = [Array]::IndexOf([string[]]$actSite, $activeSrv)

        $plan += [PSCustomObject]@{
            Database     = $dbName
            Pref1Active  = $activeSrv
            Pref2Passive = $actSite[($sIdx + 1) % $aN]
            Pref3Passive = $othSite[$sIdx       % $oN]
            Pref4Lagged  = $othSite[($sIdx + 1) % $oN]
        }
    }

    Write-StepBanner -StepName '11' -Title 'Database copy creation' -Mode $Mode -Actions @(
        "Algorithmic distribution: $($plan.Count) database(s), 3 copies each",
        "Pref2=passive (same site) | Pref3=passive (other site) | Pref4=lagged (other site)",
        "ReplayLag=$replayLag | TruncationLag=$truncLag | ReplayLagMaxDelay=$maxDelay",
        "Guarantee: 1 active + 2 passive + 1 lagged copy per server",
        "Add-MailboxDatabaseCopy for prefs 2, 3 and 4",
        "Set-MailboxDatabaseCopy (ReplayLagMaxDelay) on the lagged copy",
        "Restart-Service MSExchangeIS on each server involved (Apply only)",
        "Wait until all copies are Healthy/Mounted (10 min timeout, Apply only)",
        "Rebalance to Pref1 via Move-ActiveMailboxDatabase -ActivatePreferredOnServer:$true"
    )

    # Export the plan to Reports\
    $planPath = Join-Path $OutputFolder 'CopyPlan.csv'
    $plan | Export-Csv -Path $planPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Log "Copy plan exported: $planPath" -Level Sub

    # ==========================================================================
    # 3) Creating the copies
    # ==========================================================================
    $ex2019Names = (Get-Exchange2019Servers).Name

    foreach ($row in $plan) {
        $dbName = $row.Database

        $copyDefs = @(
            [PSCustomObject]@{ Server = $row.Pref2Passive; Pref = 2; Lagged = $false; Replay = '0.00:00:00'; Trunc = '0.00:00:00' },
            [PSCustomObject]@{ Server = $row.Pref3Passive; Pref = 3; Lagged = $false; Replay = '0.00:00:00'; Trunc = '0.00:00:00' },
            [PSCustomObject]@{ Server = $row.Pref4Lagged;  Pref = 4; Lagged = $true;  Replay = $replayLag;   Trunc = $truncLag    }
        )

        # In Simulate, the database may not exist (created only in Apply
        # by Step10). Any Add/Set-MailboxDatabaseCopy action then fails
        # with "object not found". We detect this case and report Simulated.
        $simulateAll = $false
        if ($Mode -eq 'Simulate') {
            $dbExistsNow = [bool](Get-MailboxDatabase -Identity $dbName -ErrorAction SilentlyContinue)
            if (-not $dbExistsNow) { $simulateAll = $true }
        }

        if ($simulateAll) {
            foreach ($cd in $copyDefs) {
                $srv  = $cd.Server
                $pref = $cd.Pref
                if ($srv -notin $ex2019Names) { continue }
                Add-Report -Step '11-Copies' -Target "$dbName -> $srv (pref$pref)" -Action 'Add-MailboxDatabaseCopy' `
                           -Status Simulated -BeforeValue '<database absent - will be created by Step10>' `
                           -AfterValue '(simulation - not applied)' `
                           -Detail "Pref=$pref$(if ($cd.Lagged) { ' | Lagged=True' })"
                if ($cd.Lagged -and $maxDelay -ne '0.00:00:00') {
                    Add-Report -Step '11-Copies' -Target "$dbName\$srv" -Action 'Set-MailboxDatabaseCopy' `
                               -Status Simulated -BeforeValue '<copy absent>' `
                               -AfterValue '(simulation - not applied)' `
                               -Detail "ReplayLagMaxDelay=$maxDelay"
                }
            }
            Write-Log "[Simulate] Database '$dbName' does not exist: copies deferred to the next Apply." -Level Sub
            continue
        }

        foreach ($cd in $copyDefs) {
            $srv  = $cd.Server
            $pref = $cd.Pref

            if ($srv -notin $ex2019Names) {
                Add-Report -Step '11-Copies' -Target "$dbName\$srv" -Action 'Add-MailboxDatabaseCopy' `
                           -Status Skipped -Detail 'Server is not Exchange 2019'
                continue
            }

            # ---- Add-MailboxDatabaseCopy ----------------------------------------
            Invoke-Action -Step '11-Copies' -Target "$dbName -> $srv (pref$pref)" -Action 'Add-MailboxDatabaseCopy' `
                -Detail "Pref=$pref$(if ($cd.Lagged) { ' | Lagged=True' })" `
                -InventoryScript {
                    $db = Get-MailboxDatabase -Identity $dbName -ErrorAction SilentlyContinue
                    if ($db) {
                        $ap = $db.ActivationPreference | Where-Object { $_.Key.Name -eq $srv }
                        if ($ap) { "Existing copy Pref=$($ap.Value)" } else { 'No copy' }
                    } else { '<database absent>' }
                } `
                -PreCheckScript {
                    $db = Get-MailboxDatabase -Identity $dbName -ErrorAction SilentlyContinue
                    if (-not $db) { return $false }
                    $ap = $db.ActivationPreference | Where-Object { $_.Key.Name -eq $srv }
                    return [bool]$ap
                } `
                -ActionScript {
                    $addParams = @{
                        Identity             = $dbName
                        MailboxServer        = $srv
                        ActivationPreference = $pref
                        ReplayLagTime        = $cd.Replay
                        TruncationLagTime    = $cd.Trunc
                        DomainController     = $gc
                        ErrorAction          = 'Stop'
                    }
                    Add-MailboxDatabaseCopy @addParams -WhatIf:([bool]($Mode -eq 'Simulate'))
                }

            # ---- Set-MailboxDatabaseCopy (ReplayLagMaxDelay, lagged copy) -------
            if ($cd.Lagged -and $maxDelay -ne '0.00:00:00') {
                Invoke-Action -Step '11-Copies' -Target "$dbName\$srv" -Action 'Set-MailboxDatabaseCopy' `
                    -Detail "ReplayLagMaxDelay=$maxDelay" `
                    -InventoryScript {
                        $cs = Get-MailboxDatabaseCopyStatus -Identity "$dbName\$srv" -ErrorAction SilentlyContinue
                        if ($cs) { "Status=$($cs.Status)" } else { '<copy absent>' }
                    } `
                    -PreCheckScript $null `
                    -ActionScript {
                        Set-MailboxDatabaseCopy -Identity "$dbName\$srv" `
                            -ReplayLagTime        $replayLag `
                            -TruncationLagTime    $truncLag `
                            -ReplayLagMaxDelay    $maxDelay `
                            -ActivationPreference $pref `
                            -DomainController     $gc `
                            -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                    }
            }
        }
    }

    # ==========================================================================
    # 4) MSExchangeIS restart per server (Apply only)
    # ==========================================================================
    $allSrvInvolved = @($plan | ForEach-Object {
        $_.Pref1Active; $_.Pref2Passive; $_.Pref3Passive; $_.Pref4Lagged
    } | Sort-Object -Unique | Where-Object { $_ })

    foreach ($srv in $allSrvInvolved) {
        $srvFqdn = (Get-Exchange2019Servers | Where-Object { $_.Name -eq $srv } | Select-Object -First 1).Fqdn
        if (-not $srvFqdn) { $srvFqdn = $srv }

        Invoke-Action -Step '11-Copies' -Target $srv -Action 'Restart-Service MSExchangeIS' `
            -Detail 'Restart after copy creation' `
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

    # ==========================================================================
    # 5) Wait for Healthy + Rebalance to Pref1
    # ==========================================================================
    # After the MSExchangeIS restart (Section 4), the active copy may have failed
    # over to a server other than the planned Pref1 (Managed Availability or Best
    # Copy Selection). We switch each database back to its preferred server via:
    #     Move-ActiveMailboxDatabase -Identity <db> -ActivatePreferredOnServer:$true
    #
    # Prerequisite: all copies must be Healthy/Mounted, otherwise the move
    # fails with "the source copy is not in a state that allows moving".
    # In Apply: we wait up to 10 min for the copies to stabilize.

    # ----- 5.a) Wait for Healthy (Apply only) --------------------------------
    # 2-phase strategy:
    #   Phase 1: 5 polls x 20s (= 100s) to let the copies stabilize
    #            after the MSExchangeIS restart.
    #   Phase 2: if copies are in Failed* after phase 1, we try a
    #            Suspend + Resume MailboxDatabaseCopy to unblock the
    #            replication, THEN we wait up to 10 min for everything to become
    #            Healthy/Mounted.
    if ($Mode -eq 'Apply') {
        $dbNames = @($plan | ForEach-Object { $_.Database })
        Write-Log "Waiting for the copies to be Healthy/Mounted (phase 1: 5 polls x 20s)..." -Level Info

        $allHealthy   = $false
        $lastSummary  = ''
        $lastStatuses = @()

        # ----- Phase 1: initial polling (5 x 20s = 100s) ---------------------
        for ($try = 1; $try -le 5; $try++) {
            $statuses = foreach ($db in $dbNames) {
                Get-MailboxDatabaseCopyStatus -Identity $db -ErrorAction SilentlyContinue
            }
            $lastStatuses = $statuses
            $unhealthy = @($statuses | Where-Object { $_.Status -notin 'Mounted','Healthy' })
            if ($unhealthy.Count -eq 0) {
                $allHealthy = $true
                break
            }
            $lastSummary = ($unhealthy | Select-Object -First 8 |
                            ForEach-Object { "$($_.Name)=$($_.Status)" }) -join ', '
            Write-Log ("  Phase 1 try {0}/5: {1} copy(ies) not Healthy: {2}{3}" -f `
                       $try, $unhealthy.Count, $lastSummary,
                       $(if ($unhealthy.Count -gt 8) {' ...'})) -Level Sub
            if ($try -lt 5) { Start-Sleep -Seconds 20 }
        }

        # ----- Phase 2: Suspend+Resume if Failed* + 10 min wait --------------
        if (-not $allHealthy) {
            $failedCopies = @($lastStatuses | Where-Object {
                $_.Status -in 'Failed','FailedAndSuspended','Disconnected','DisconnectedAndResynchronizing'
            })

            if ($failedCopies.Count -gt 0) {
                Write-Log ("  Phase 2: Suspend + Resume on {0} Failed/Disconnected copy(ies)" -f $failedCopies.Count) -Level Warning
                foreach ($copy in $failedCopies) {
                    $copyId = [string]$copy.Identity
                    Write-Log "    -> $copyId (status=$($copy.Status))" -Level Sub
                    try {
                        Suspend-MailboxDatabaseCopy -Identity $copyId -Confirm:$false `
                            -SuspendComment 'Step11 auto-fix: Suspend+Resume after Failed' -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 3
                        Resume-MailboxDatabaseCopy -Identity $copyId -Confirm:$false -ErrorAction SilentlyContinue
                    } catch {
                        Write-Log "    Suspend/Resume error on $copyId - $($_.Exception.Message)" -Level Warning
                    }
                }
                Add-Report -Step '11-Copies' -Target '<DAG>' -Action 'Suspend+Resume Failed copies' -Status Success `
                           -Detail ("{0} copy(ies) fixed: {1}" -f $failedCopies.Count,
                                    (($failedCopies | ForEach-Object { [string]$_.Identity }) -join ' | '))
            }

            # Wait up to 10 min for everything to become Healthy/Mounted
            Write-Log "  Phase 2: waiting for Healthy/Mounted (10 min timeout)..." -Level Info
            $deadline = (Get-Date).AddMinutes(10)
            do {
                Start-Sleep -Seconds 20
                $statuses = foreach ($db in $dbNames) {
                    Get-MailboxDatabaseCopyStatus -Identity $db -ErrorAction SilentlyContinue
                }
                $unhealthy = @($statuses | Where-Object { $_.Status -notin 'Mounted','Healthy' })
                if ($unhealthy.Count -eq 0) {
                    $allHealthy = $true
                    break
                }
                $lastSummary = ($unhealthy | Select-Object -First 8 |
                                ForEach-Object { "$($_.Name)=$($_.Status)" }) -join ', '
                Write-Log ("    Still {0} copy(ies) not Healthy: {1}{2}" -f `
                           $unhealthy.Count, $lastSummary,
                           $(if ($unhealthy.Count -gt 8) {' ...'})) -Level Sub
            } while ((Get-Date) -lt $deadline)
        }

        # ----- Summary -------------------------------------------------------
        if ($allHealthy) {
            Add-Report -Step '11-Copies' -Target '<DAG>' -Action 'WaitCopiesHealthy' -Status Success `
                       -Detail ("All copies Healthy/Mounted ({0} databases x 4 copies)" -f $dbNames.Count)
            Write-Log "All copies are Healthy/Mounted - rebalance possible." -Level Success
        } else {
            Add-Report -Step '11-Copies' -Target '<DAG>' -Action 'WaitCopiesHealthy' -Status Failed `
                       -ErrorMessage "After phase 1 + Suspend/Resume + 10 min: copies still not Healthy: $lastSummary"
            Write-Log "Rebalance CANCELLED: copies not Healthy. Rerun Step11 later." -Level Error
            Save-Report -StepName 'Step11-CreateDatabaseCopies'
            return 1
        }
    }

    # ----- 5.b) Rebalance to Pref1 (3 modes) ---------------------------------
    foreach ($row in $plan) {
        $dbName = $row.Database
        $pref1  = $row.Pref1Active

        # In Simulate, if the database does not exist yet: pre-simulate without calling Get/Move
        if ($Mode -eq 'Simulate') {
            $dbExistsNow = [bool](Get-MailboxDatabase -Identity $dbName -ErrorAction SilentlyContinue)
            if (-not $dbExistsNow) {
                Add-Report -Step '11-Copies' -Target $dbName -Action 'Move-ActiveMailboxDatabase (rebalance)' `
                           -Status Simulated -BeforeValue '<database absent - will be created by Step10>' `
                           -AfterValue '(simulation - not applied)' -Detail "To Pref1=$pref1"
                continue
            }
        }

        Invoke-Action -Step '11-Copies' -Target $dbName -Action 'Move-ActiveMailboxDatabase (rebalance)' `
            -Detail "To Pref1=$pref1" `
            -InventoryScript {
                $b = Get-MailboxDatabase -Identity $dbName -Status -ErrorAction SilentlyContinue
                if ($b) {
                    "Active={0}; Mounted={1}; Pref1Target={2}" -f ([string]$b.Server), $b.Mounted, $pref1
                } else { '<database absent>' }
            } `
            -PreCheckScript {
                $b = Get-MailboxDatabase -Identity $dbName -Status -ErrorAction Stop
                return (([string]$b.Server) -ieq $pref1 -and $b.Mounted)
            } `
            -ActionScript {
                # Move-ActiveMailboxDatabase has 2 mutually exclusive parameter sets:
                #   - Identity : -Identity <db> + -ActivateOnServer <srv>  (targets 1 database)
                #   - Server   : -Server <srv> + -ActivatePreferredOnServer (switch, per server)
                # We have $pref1 from the algorithmic plan, so we use the Identity set.
                #
                # Skip* flags: Exchange limits repeated moves via Active Manager
                # ('too many moves have happened recently'). At Step11, the databases are
                # freshly created, with no users: we bypass these protections.
                #   -SkipMoveSuppressionChecks : bypass anti-thrashing
                #   -SkipClientExperienceChecks : no users on these databases
                #   -SkipMaximumActiveDatabasesChecks : MAD limit ignored
                Move-ActiveMailboxDatabase -Identity $dbName -ActivateOnServer $pref1 `
                    -SkipMoveSuppressionChecks -SkipClientExperienceChecks `
                    -SkipMaximumActiveDatabasesChecks `
                    -MoveComment 'Step11 auto-rebalance to Pref1' `
                    -Confirm:$false -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop | Out-Null
            }
    }

    Save-Report -StepName 'Step11-CreateDatabaseCopies'
    return 0
}
