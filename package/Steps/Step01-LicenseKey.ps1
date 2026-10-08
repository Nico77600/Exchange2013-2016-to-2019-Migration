<#
.SYNOPSIS
    Step 01 - Exchange 2019 product key activation.

.DESCRIPTION
    Activates the Exchange 2019 product key, then restarts the MSExchangeIS service on each
    Exchange 2019 server.

    Server source: Get-Exchange2019Servers (AD discovery via Get-ExchangeServer,
                   filter on Version 15.2* and exclusion of Edge Transport)
    Product key: Config.LicenseKey

    Logic:
      - Inventory: reads the current version (Edition, ProductID/Trial)
      - Simulate: simulation with -WhatIf
      - Apply: Set-ExchangeServer -ProductKey ... then remote Restart-Service MSExchangeIS

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

    Write-StepBanner -StepName '01' -Title 'Exchange 2019 product key activation' -Mode $Mode -Actions @(
        'Discover the Exchange 2019 servers via Get-Exchange2019Servers (AD, Edge excluded)',
        'Compare the current product key with the target key (idempotency)',
        'Apply Set-ExchangeServer -ProductKey if needed',
        'Restart the MSExchangeIS service so that the license takes effect'
    )

    Initialize-ExchangeShell

    $serversList = @(Get-Exchange2019Servers)
    if (-not $serversList -or $serversList.Count -eq 0) {
        Write-Log 'No Exchange 2019 server detected.' -Level Warning
        return 1
    }

    if ($Mode -ne 'Inventory' -and (-not $Config.LicenseKey -or $Config.LicenseKey -match 'XXXXX-XXXXX')) {
        Write-Log 'No product key defined in Deployment.config.psd1 (LicenseKey).' -Level Error
        Add-Report -Step '01-LicenseKey' -Target '<global>' -Action 'CheckConfig' -Status Failed `
                   -Detail 'LicenseKey missing or still a placeholder in the config'
        return 2
    }

    foreach ($srv in $serversList) {
        $server = $srv.Name

        Invoke-Action -Step '01-LicenseKey' -Target $server -Action 'ApplyLicense' -Detail 'Set-ExchangeServer -ProductKey' `
            -InventoryScript {
                $srv = Get-ExchangeServer $server -ErrorAction Stop
                "Edition={0}; AdminDisplayVersion={1}; IsExchangeTrialEdition={2}" -f `
                    $srv.Edition, $srv.AdminDisplayVersion, $srv.IsExchangeTrialEdition
            } `
            -PreCheckScript {
                $srv = Get-ExchangeServer $server -ErrorAction Stop
                # If the license is no longer in Trial mode -> already licensed
                return ($srv.IsExchangeTrialEdition -eq $false)
            } `
            -ActionScript {
                Set-ExchangeServer -Identity $server -ProductKey $Config.LicenseKey -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
            }

        Invoke-Action -Step '01-LicenseKey' -Target $server -Action 'Restart-Service MSExchangeIS' `
            -Detail 'Restart so that the product key takes effect' `
            -InventoryScript {
                $s = Get-Service -ComputerName $server -Name 'MSExchangeIS' -ErrorAction SilentlyContinue
                if ($s) { "Status=$($s.Status)" } else { '<service not found>' }
            } `
            -PreCheckScript $null `
            -ActionScript {
                $svc = Get-Service -ComputerName $server -Name 'MSExchangeIS' -ErrorAction Stop
                if ($svc.Status -eq 'Running') {
                    Restart-Service -InputObject $svc -Force -ErrorAction Stop
                } else {
                    Start-Service -InputObject $svc -ErrorAction Stop
                }
            }
    }

    return 0
}
