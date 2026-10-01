<#
.SYNOPSIS
    Step 12 - Add the X-Forwarded-For field to the IIS logs.

.DESCRIPTION
    Context: behind a load balancer or a reverse proxy (HAProxy, KEMP, F5, ...),
    the source IP seen by IIS is the IP of the load balancer. The real client IP
    is passed in the X-Forwarded-For HTTP header. This step adds that header as an
    IIS custom log field on the two Exchange sites: Default Web Site and
    Exchange Back End.

    Reference: https://docs.microsoft.com/en-us/iis/get-started/whats-new-in-iis-85/enhanced-logging-for-iis85

    Idempotent: if the X-Forwarded-For field already exists on the site,
    AlreadyDone is returned.

    The action is performed remotely through Invoke-Command on each
    Exchange 2019 server (WebAdministration module / appcmd). No action is
    taken on the Exchange 2013/2016 servers.

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

    Write-StepBanner -StepName '12' -Title 'Add X-Forwarded-For to the IIS logs' -Mode $Mode -Actions @(
        'On each Exchange 2019 server:',
        '  - Add the X-Forwarded-For custom field on Default Web Site',
        '  - Add the X-Forwarded-For custom field on Exchange Back End',
        '  - Source: RequestHeader X-Forwarded-For',
        'No action on the Exchange 2013/2016 servers'
    )

    Reset-Report
    Initialize-ExchangeShell -Credential $Credential
    $servers = Get-Exchange2019Servers
    if (-not $servers) {
        Write-Log "No Exchange 2019 server detected." -Level Warning
        Save-Report -StepName 'Step12-IISXForwardedFor'
        return 0
    }

    $sites = @('Default Web Site', 'Exchange Back End')

    foreach ($srv in $servers) {
        $icArgs = @{ ComputerName = $srv.Fqdn }
        if ($Credential) { $icArgs['Credential'] = $Credential }

        foreach ($site in $sites) {
            Invoke-Action -Step '12-XFF' -Target "$($srv.Name)\$site" -Action 'Add custom log field X-Forwarded-For' `
                -InventoryScript {
                    $sb = {
                        param($siteName)
                        Import-Module WebAdministration -ErrorAction Stop
                        $path = "/system.applicationHost/sites/site[@name='$siteName']/logFile/customFields"
                        try {
                            $cf = Get-WebConfigurationProperty -Filter $path -PSPath 'IIS:\' -Name '.' -ErrorAction Stop
                            if ($cf -and $cf.Collection) {
                                ($cf.Collection | ForEach-Object { "$($_.sourceType):$($_.sourceName)" }) -join '; '
                            } else { 'No custom field' }
                        } catch { 'No custom field' }
                    }
                    Invoke-Command @icArgs -ScriptBlock $sb -ArgumentList $site
                } `
                -PreCheckScript {
                    $sb = {
                        param($siteName)
                        Import-Module WebAdministration -ErrorAction Stop
                        $path = "/system.applicationHost/sites/site[@name='$siteName']/logFile/customFields"
                        try {
                            $cf = Get-WebConfigurationProperty -Filter $path -PSPath 'IIS:\' -Name '.' -ErrorAction SilentlyContinue
                            if ($cf -and $cf.Collection) {
                                return [bool]($cf.Collection | Where-Object { $_.sourceName -eq 'X-Forwarded-For' -and $_.sourceType -eq 'RequestHeader' })
                            }
                        } catch { }
                        return $false
                    }
                    Invoke-Command @icArgs -ScriptBlock $sb -ArgumentList $site
                } `
                -ActionScript {
                    $sb = {
                        param($siteName)
                        Import-Module WebAdministration -ErrorAction Stop
                        $filter = "/system.applicationHost/sites/site[@name='$siteName']/logFile/customFields"
                        Add-WebConfigurationProperty -Filter $filter -PSPath 'IIS:\' -Name '.' -Value @{
                            logFieldName = 'X-Forwarded-For'
                            sourceName   = 'X-Forwarded-For'
                            sourceType   = 'RequestHeader'
                        } -ErrorAction Stop
                    }
                    Invoke-Command @icArgs -ScriptBlock $sb -ArgumentList $site -ErrorAction Stop
                }
        }
    }

    Save-Report -StepName 'Step12-IISXForwardedFor' | Out-Null
    return 0
}
