<#
.SYNOPSIS
    Step 25 - Autodiscover Service Connection Point (SCP) cleanup.

.DESCRIPTION
    Cleans up the Autodiscover Service Connection Point (SCP) on the
    Exchange 2013 and 2016 servers ONLY.

    WHY: after the cutover to Exchange 2019, Outlook clients joined to the
    AD domain query the SCP to discover the Autodiscover URL. If the SCPs of
    the 2013/2016 servers are still valid (AutoDiscoverServiceInternalUri
    attribute not empty), clients may land on them randomly (round-robin
    between all the SCPs found) and try to connect to a legacy server. This
    brings IIS traffic back to the old servers and prevents their
    decommissioning.

    ACTION: clears AutoDiscoverServiceInternalUri on each 2013/2016
    ClientAccessService via
    Set-ClientAccessService -AutoDiscoverServiceInternalUri $null.
    Result: their SCP is no longer returned by the client AD lookup.

    PREREQUISITES:
      - The mailbox cutover must be COMPLETED (Step21+22).
      - The 2019 Autodiscover services must be working (can be checked with
        Step26-AnalyzeProtocolLogs: Autodiscover 2019 must show user
        traffic).

    POST-CHECK: after applying this step, run Step26-AnalyzeProtocolLogs
    again a few hours later to confirm that the Autodiscover column of the
    legacy versions dropped to 0.

    Modes:
      - Inventory: lists the legacy CAS servers with their current
                   AutoDiscoverServiceInternalUri. No change is made.
      - Simulate : -WhatIf through Invoke-Action.
      - Apply    : Set-ClientAccessService -AutoDiscoverServiceInternalUri $null
                   on each legacy server detected.

    STRICT FILTER: touches ONLY Exchange 2013 (15.0.*) and 2016 (15.1.*).
    The 2019 (15.2.*) servers and the Edge servers are automatically excluded
    through Get-LegacyExchangeServers of the Common module.

    ManualOnly: this step is NOT run with -Step All -Mode Apply. It must be
    called explicitly (-Step 25 -Mode Apply -Force) and only after the
    cutover has been validated.

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

    Write-StepBanner -StepName '25' -Title 'Autodiscover SCP cleanup on Exchange 2013/2016' -Mode $Mode -Actions @(
        'Target   : Exchange 2013 (15.0) + 2016 (15.1) ONLY (never 2019)',
        'Action   : Set-ClientAccessService -AutoDiscoverServiceInternalUri $null',
        'Effect   : clients no longer land randomly on a legacy SCP',
        'Requires : mailbox cutover completed, Autodiscover 2019 working',
        'Post-check: Step 26 -Mode Inventory (legacy Autodiscover must be 0)',
        'ManualOnly: not run with -Step All; must be called explicitly.'
    )

    Reset-Report
    Initialize-ExchangeShell -Credential $Credential

    # Get-LegacyExchangeServers returns the 2013 (15.0.*) + 2016 (15.1.*) servers,
    # and automatically excludes 2019 and Edge.
    $legacy = @(Get-LegacyExchangeServers -ErrorAction SilentlyContinue)
    if (-not $legacy) {
        Write-Log "No Exchange 2013/2016 server found in the organization." -Level Warning
        Write-Log "Nothing to clean up. Exiting." -Level Sub
        Save-Report -StepName 'Step25-CleanupAutodiscoverSCP' | Out-Null
        return 0
    }

    Write-Log ('=' * 70) -Level Step
    Write-Log "Legacy Exchange servers detected:" -Level Step
    foreach ($s in $legacy | Sort-Object Name) {
        Write-Log ("  {0,-12} | {1}" -f $s.Name, $s.AdminDisplayVersion) -Level Info
    }
    Write-Log ('=' * 70) -Level Step

    # Helper for the Invoke-Action scriptblock scope (see PS 5.1 pattern)
    foreach ($srv in $legacy) {
        $srvName = [string]$srv.Name
        $srvVer  = [string]$srv.AdminDisplayVersion

        Invoke-Action -Step '25-SCP' -Target $srvName -Action 'Clear AutoDiscoverServiceInternalUri' `
            -Detail "Version=$srvVer; target=AutoDiscoverServiceInternalUri empty" `
            -InventoryScript {
                $cas = Get-ClientAccessService -Identity $srvName -ErrorAction SilentlyContinue
                if ($cas) {
                    $uri = [string]$cas.AutoDiscoverServiceInternalUri
                    if ([string]::IsNullOrEmpty($uri)) { '<empty>' } else { $uri }
                } else { '<absent>' }
            } `
            -PreCheckScript {
                $cas = Get-ClientAccessService -Identity $srvName -ErrorAction SilentlyContinue
                if (-not $cas) { return $true }
                return [string]::IsNullOrEmpty([string]$cas.AutoDiscoverServiceInternalUri)
            } `
            -ActionScript {
                Set-ClientAccessService -Identity $srvName -AutoDiscoverServiceInternalUri $null -Confirm:$false -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
            }
    }

    Save-Report -StepName 'Step25-CleanupAutodiscoverSCP' | Out-Null
    return 0
}
