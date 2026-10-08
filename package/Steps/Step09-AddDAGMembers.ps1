<#
.SYNOPSIS
    Step 09 - Add the Exchange 2019 servers to the DAG.

.DESCRIPTION
    Adds the Exchange 2019 servers to the Database Availability Group (DAG).

    Source: DAGInfo.csv
        - Site1Servers / Site2Servers : server names (comma-separated CSV values)
        - MaximumActiveDatabasesSite1/2, MaximumPreferredActiveDatabasesSite1/2
        - AutoDatabaseMountDial

    Prerequisites:
        - The DAG must already exist (step 08)
        - Each server must be an Exchange 2019 server (defensive filter)

    Idempotent: if the server is already a member of the DAG, the action is AlreadyDone.

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

    Write-StepBanner -StepName '09' -Title 'Add the Exchange 2019 servers to the DAG' -Mode $Mode -Actions @(
        'For each DAG defined in DAGInfo.csv:',
        '  - Check that the DAG exists (otherwise error => rerun step 08)',
        '  - For each Site1/Site2 server: check that it is an Exchange 2019 server',
        '  - Configure MaximumActiveDatabases / MaximumPreferredActiveDatabases / AutoDatabaseMountDial',
        '  - Add it to the DAG via Add-DatabaseAvailabilityGroupServer if not already a member'
    )

    Initialize-ExchangeShell

    $dagsCsv = Import-ConfigCsv -Path (Join-Path $CsvFolder 'DAGInfo.csv') | Where-Object { $_.DAGName }

    foreach ($dag in $dagsCsv) {
        $dagName = $dag.DAGName
        $existing = Get-DatabaseAvailabilityGroup -Identity $dagName -ErrorAction SilentlyContinue
        if (-not $existing) {
            if ($Mode -eq 'Inventory') {
                Add-Report -Step '09-AddDAGMembers' -Target $dagName -Action 'AddMembers' `
                           -Status Inventoried -BeforeValue '<absent>' `
                           -Detail "DAG not found in AD - to be created via step 08 before the Apply"
                continue
            }
            if ($Mode -eq 'Simulate') {
                # DAG missing: it will be created by step 08 in Apply. We simulate
                # adding the members for each configured server.
                $serversCsv = @(($dag.Site1Servers -split ',') + ($dag.Site2Servers -split ','))
                $serversCsv = $serversCsv | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                foreach ($s in $serversCsv) {
                    Add-Report -Step '09-AddDAGMembers' -Target $s -Action 'Add-DatabaseAvailabilityGroupServer' `
                               -Status Simulated -BeforeValue '<DAG does not exist - simulated by Step08>' `
                               -AfterValue '(simulation - not applied)' -Detail "DAG=$dagName"
                }
                Write-Log "[Simulate] DAG '$dagName' does not exist: adding the $($serversCsv.Count) members deferred to the next Apply." -Level Sub
                continue
            }
            # Apply but DAG missing -> real error
            Add-Report -Step '09-AddDAGMembers' -Target $dagName -Action 'AddMembers' -Status Failed `
                       -Detail "DAG $dagName does not exist. Run step 08 first."
            continue
        }

        $mountDial = $dag.AutoDatabaseMountDial
        $sites = @(
            @{ Servers = $dag.Site1Servers -split ','; Mad = [int]$dag.MaximumActiveDatabasesSite1; Mpad = [int]$dag.MaximumPreferredActiveDatabasesSite1 },
            @{ Servers = $dag.Site2Servers -split ','; Mad = [int]$dag.MaximumActiveDatabasesSite2; Mpad = [int]$dag.MaximumPreferredActiveDatabasesSite2 }
        )

        foreach ($site in $sites) {
            foreach ($s in $site.Servers) {
                $server = $s.Trim()
                if (-not $server) { continue }

                if (-not (Test-IsExchange2019Server -ServerName $server)) {
                    Add-Report -Step '09-AddDAGMembers' -Target $server -Action 'AddDAGMember' -Status Skipped `
                               -Detail "$server is not an Exchange 2019 server - coexistence protection"
                    continue
                }

                $mad  = $site.Mad
                $mpad = $site.Mpad

                # MaximumActiveDatabases / Preferred
                Invoke-Action -Step '09-AddDAGMembers' -Target $server -Action 'Set-MailboxServer (limits)' -Detail "MAD=$mad MPAD=$mpad MountDial=$mountDial" `
                    -InventoryScript {
                        $sv = Get-MailboxServer $server -ErrorAction Stop
                        "MAD={0}; MPAD={1}; MountDial={2}" -f $sv.MaximumActiveDatabases, $sv.MaximumPreferredActiveDatabases, $sv.AutoDatabaseMountDial
                    } `
                    -PreCheckScript {
                        $sv = Get-MailboxServer $server -ErrorAction Stop
                        return ($sv.MaximumActiveDatabases -eq $mad -and `
                                ($mpad -eq 0 -or $sv.MaximumPreferredActiveDatabases -eq $mpad) -and `
                                $sv.AutoDatabaseMountDial -eq $mountDial)
                    } `
                    -ActionScript {
                        Set-MailboxServer -Identity $server -MaximumActiveDatabases $mad -AutoDatabaseMountDial $mountDial `
                            -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                        if ($mpad -ne 0) {
                            Set-MailboxServer -Identity $server -MaximumPreferredActiveDatabases $mpad `
                                -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                        }
                    }

                # Add to DAG
                Invoke-Action -Step '09-AddDAGMembers' -Target $server -Action 'Add-DatabaseAvailabilityGroupServer' -Detail "DAG=$dagName" `
                    -InventoryScript {
                        $d = Get-DatabaseAvailabilityGroup -Identity $dagName -ErrorAction Stop
                        if ($server -in $d.Servers.Name) { "$server is already a member" } else { "$server is not yet a member" }
                    } `
                    -PreCheckScript {
                        $d = Get-DatabaseAvailabilityGroup -Identity $dagName -ErrorAction Stop
                        return ($server -in $d.Servers.Name)
                    } `
                    -ActionScript {
                        Add-DatabaseAvailabilityGroupServer -Identity $dagName -MailboxServer $server `
                            -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                    }
            }
        }

        # --- Deferred activation of DatacenterActivationMode --------------------
        # Step08 cannot enable DACMode while the DAG is empty (Exchange refuses
        # with "fewer than two mailbox servers"). Once the members are added (above),
        # it can now be enabled. Idempotent: skipped if already in the right mode.
        $dacExpected = $Config.DAG.DACMode
        Invoke-Action -Step '09-AddDAGMembers' -Target $dagName -Action 'Set-DAG DatacenterActivationMode (post-members)' `
            -Detail "DatacenterActivationMode=$dacExpected (requires 2+ members)" `
            -InventoryScript {
                $d = Get-DatabaseAvailabilityGroup -Identity $dagName -ErrorAction Stop
                "DAC={0}; Members={1}" -f $d.DatacenterActivationMode, @($d.Servers).Count
            } `
            -PreCheckScript {
                $d = Get-DatabaseAvailabilityGroup -Identity $dagName -ErrorAction Stop
                return ($d.DatacenterActivationMode -eq $dacExpected)
            } `
            -ActionScript {
                $d = Get-DatabaseAvailabilityGroup -Identity $dagName -ErrorAction Stop
                $count = @($d.Servers).Count
                if ($count -lt 2) {
                    if ($Mode -eq 'Simulate') {
                        Write-Log "[Simulate] Set-DAG -DACMode deferred: DAG $dagName has $count member(s) (additions simulated just before)" -Level Sub
                        return
                    }
                    throw "DAG $dagName has $count member(s) - DAC cannot be enabled without at least 2 members. Check that the previous Add-DatabaseAvailabilityGroupServer actions succeeded."
                }
                Set-DatabaseAvailabilityGroup -Identity $dagName -DatacenterActivationMode $dacExpected `
                    -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
            }
    }

    return 0
}
