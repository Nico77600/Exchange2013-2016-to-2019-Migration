<#
.SYNOPSIS
    Step 08 - Create the DAG in IP-less mode.

.DESCRIPTION
    Creates the Database Availability Group (DAG) in IP-less mode.

    Source: Configs\DAGInfo.csv.
    Particularities:
        - IP-less mode: DAGIp = 255.255.255.255 -> [System.Net.IPAddress]::None is passed
        - DACMode, NetworkEncryption, NetworkCompression, SafetyNetHoldTime: read from
          Config.DAG (fixed values for all DAGs)
        - Witness server / directory, ManualDagNetworkConfiguration,
          ReplayLagManagerEnabled, ReplicationPort: read from DAGInfo.csv

    Adding the members is handled in step 09.

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

    Write-StepBanner -StepName '08' -Title 'DAG creation in IP-less mode' -Mode $Mode -Actions @(
        'Read DAGInfo.csv',
        'Check whether the DAG already exists (idempotency)',
        'If missing: New-DatabaseAvailabilityGroup in IP-less mode ([IPAddress]::None)',
        "DACMode=$($Config.DAG.DACMode) | Encryption=$($Config.DAG.NetworkEncryption) | Compression=$($Config.DAG.NetworkCompression) | SafetyNet=$($Config.DAG.SafetyNetHoldTime) day(s) (from Config.DAG)",
        'Configure Witness, ReplicationPort, ReplayLagManager (from DAGInfo.csv)',
        'No member is added in this step (see step 09)'
    )

    Initialize-ExchangeShell

    $csv = Join-Path $CsvFolder 'DAGInfo.csv'
    $dagsCsv = Import-ConfigCsv -Path $csv | Where-Object { $_.DAGName }

    foreach ($dag in $dagsCsv) {
        $name      = $dag.DAGName
        $witness   = $dag.WitnessServer
        $witDir    = $dag.WitnessDir
        $altWit    = $dag.AltWitnessServer
        $altWitDir = $dag.AltWitnessDir
        $gc        = $dag.GC

        # IP-less if 255.255.255.255
        $dagIps = if ([string]$dag.DAGIps -eq '255.255.255.255') {
            ,[System.Net.IPAddress]::None
        } else {
            $dag.DAGIps -split ',' | ForEach-Object { [System.Net.IPAddress]$_.Trim() }
        }

        # Fixed values (Config.DAG)
        $safetyNet = [int]$Config.DAG.SafetyNetHoldTime
        $dac       = $Config.DAG.DACMode
        $netEnc    = $Config.DAG.NetworkEncryption
        $netComp   = $Config.DAG.NetworkCompression
        # Per-DAG values (DAGInfo.csv)
        $manualNet = [System.Convert]::ToBoolean($dag.ManualDagnetworkConfiguration)
        $replayLag = [System.Convert]::ToBoolean($dag.ReplayLagManagerEnabled)
        $replPort  = [uint16]$dag.ReplicationPort

        # --- DAG creation -------------------------------------------------------
        Invoke-Action -Step '08-CreateDAG' -Target $name -Action 'New-DatabaseAvailabilityGroup' -Detail "IP-Less, Witness=$witness" `
            -InventoryScript {
                $existing = Get-DatabaseAvailabilityGroup -Identity $name -ErrorAction SilentlyContinue
                if ($existing) {
                    "DAG exists: Witness={0}; AltWitness={1}; DAC={2}; ManualNet={3}" -f `
                        $existing.WitnessServer, $existing.AlternateWitnessServer,
                        $existing.DatacenterActivationMode, $existing.ManualDagNetworkConfiguration
                } else { '<absent>' }
            } `
            -PreCheckScript {
                return [bool](Get-DatabaseAvailabilityGroup -Identity $name -ErrorAction SilentlyContinue)
            } `
            -ActionScript {
                # SafetyNetHoldTime first
                $tc = Get-TransportConfig -ErrorAction Stop
                if ($tc.SafetyNetHoldTime.TotalDays -lt $safetyNet) {
                    Set-TransportConfig -SafetyNetHoldTime ([TimeSpan]::FromDays($safetyNet)) -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                }

                $newParams = @{
                    Name              = $name
                    WitnessServer     = $witness
                    WitnessDirectory  = $witDir
                    DatabaseAvailabilityGroupIPAddresses = $dagIps
                    DomainController  = $gc
                    ErrorAction       = 'Stop'
                }
                New-DatabaseAvailabilityGroup @newParams -WhatIf:([bool]($Mode -eq 'Simulate')) | Out-Null
            }

        # --- Post-creation configuration ----------------------------------------
        # Run even if AlreadyDone to ensure convergence.
        # Special case in Simulate: if the DAG does not exist yet (it has just
        # been "created" via -WhatIf), Get/Set cannot be used on it -> we
        # report Simulated directly.
        $dagExists = [bool](Get-DatabaseAvailabilityGroup -Identity $name -ErrorAction SilentlyContinue)
        if (-not $dagExists -and $Mode -eq 'Simulate') {
            $detail = ("DAC={0}; Encryption={1}; Compression={2}; ReplPort={3}; ReplayLag={4}; ManualNet={5}" -f `
                $dac, $netEnc, $netComp, $replPort, $replayLag, $manualNet)
            Add-Report -Step '08-CreateDAG' -Target $name -Action 'Set-DatabaseAvailabilityGroup' `
                       -Status Simulated -BeforeValue '<DAG does not exist - simulated by New->Set>' `
                       -AfterValue '(simulation - not applied)' -Detail $detail
            if ($altWit) {
                Add-Report -Step '08-CreateDAG' -Target $name -Action 'Set-DatabaseAvailabilityGroup AltWitness' `
                           -Status Simulated -BeforeValue '<DAG does not exist - simulated>' `
                           -AfterValue '(simulation - not applied)' `
                           -Detail "AltWitness=$altWit; AltWitnessDir=$altWitDir"
            }
            Write-Log "[Simulate] DAG '$name' does not exist: Set-DAG deferred to the next Apply." -Level Sub
            continue
        }

        Invoke-Action -Step '08-CreateDAG' -Target $name -Action 'Set-DatabaseAvailabilityGroup' -Detail "DAC + Network" `
            -InventoryScript {
                $d = Get-DatabaseAvailabilityGroup -Identity $name -ErrorAction SilentlyContinue
                if (-not $d) { return '<absent>' }
                "DAC={0}; Encryption={1}; Compression={2}; ReplPort={3}; ReplayLag={4}" -f `
                    $d.DatacenterActivationMode, $d.NetworkEncryption, $d.NetworkCompression,
                    $d.ReplicationPort, $d.ReplayLagManagerEnabled
            } `
            -PreCheckScript {
                $d = Get-DatabaseAvailabilityGroup -Identity $name -ErrorAction Stop
                $memberCount = @($d.Servers).Count
                # DACMode can ONLY be checked if the DAG has >= 2 members (otherwise Exchange
                # refuses the activation). With fewer than 2 members, we consider it "ok for
                # now" - Step09 will finalize the activation after the members are added.
                $dacOk = if ($memberCount -ge 2) { $d.DatacenterActivationMode -eq $dac } else { $true }
                return ($dacOk -and `
                        $d.NetworkEncryption -eq $netEnc -and `
                        $d.NetworkCompression -eq $netComp -and `
                        $d.ReplicationPort -eq $replPort -and `
                        $d.ReplayLagManagerEnabled -eq $replayLag -and `
                        $d.ManualDagNetworkConfiguration -eq $manualNet)
            } `
            -ActionScript {
                # We read the member count to decide whether DACMode can be enabled.
                # With 0 or 1 member, Set-DAG -DatacenterActivationMode fails with
                # "DAG cannot be set into datacenter activation mode, since it contains
                # fewer than two mailbox servers". Step09 reapplies DACMode after the members are added.
                $dagObj = Get-DatabaseAvailabilityGroup -Identity $name -ErrorAction Stop
                $memberCount = @($dagObj.Servers).Count

                $setParams = @{
                    Identity                      = $name
                    ManualDagNetworkConfiguration = $manualNet
                    NetworkEncryption             = $netEnc
                    NetworkCompression            = $netComp
                    ReplayLagManagerEnabled       = $replayLag
                    ReplicationPort               = $replPort
                    ErrorAction                   = 'Stop'
                    WhatIf                        = [bool]($Mode -eq 'Simulate')
                }
                if ($memberCount -ge 2) {
                    $setParams['DatacenterActivationMode'] = $dac
                } else {
                    Write-Log "DAG $name has $memberCount member(s) (<2) - DatacenterActivationMode=$dac will be applied after the members are added (Step09)." -Level Warning
                }
                Set-DatabaseAvailabilityGroup @setParams

                if ($altWit) {
                    Set-DatabaseAvailabilityGroup -Identity $name `
                        -AlternateWitnessServer $altWit `
                        -AlternateWitnessDirectory $altWitDir `
                        -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                }
            }
    }

    return 0
}
