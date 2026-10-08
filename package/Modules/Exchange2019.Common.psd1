#
#  Exchange 2013/2016 to 2019 Migration - module manifest
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : see ModuleVersion
#
#  Loaded by Deploy-Exchange2019.ps1 (Import-Module by path). The 26 steps are dot-sourced by the
#  orchestrator and call the exported functions below.
#
@{
    RootModule        = 'Exchange2019.Common.psm1'
    ModuleVersion     = '2.0.0'
    GUID              = '5b0f7a3e-4f1c-4c39-9b8e-2f6d1c7a9e41'
    Author            = 'Nicolas Fabert'
    Description       = 'Exchange 2013/2016 to 2019 Migration: shared functions of the deployment framework (console, run log, reports, modes, Exchange helpers, resumable state).'
    PowerShellVersion = '5.1'

    # Functions called by the orchestrator, the steps, the tools and the tests. Add a function here
    # only when one of them calls it; the other functions stay internal to the module.
    FunctionsToExport = @(
        # Console and run log
        'Write-Log', 'Write-StepBanner', 'Write-ExHost', 'Write-ExBanner', 'Write-ExStep', 'Write-ExItem', 'Write-ExSummary', 'Write-ExRule'
        'Start-ExRun', 'Stop-ExRun', 'Write-ExLog', 'Set-ExStepContext', 'Format-ExDuration', 'Format-ExNumber', 'Get-ExIcon', 'Get-ExToolInfo'
        'Get-ExIconSet', 'Get-ExFrameSet', 'Enable-ExConsoleCapture', 'Get-ExConsoleCapture'
        # Step catalogue
        'Get-DeploymentStepCatalog'
        # Session and reports
        'Initialize-DeploymentSession', 'Stop-DeploymentSession'
        'Add-Report', 'Save-Report', 'Reset-Report', 'Get-ExRunSummary'
        'Invoke-Action'
        # HTML theme shared by every report
        'ConvertTo-SafeHtml', 'Get-ExHtmlHead', 'Get-ExHtmlHero', 'Get-ExHtmlTile', 'Get-ExHtmlFooter', 'Get-ExStatusColor', 'New-HtmlReport'
        # Exchange helpers
        'Set-DefaultExchangeServer', 'Get-ExchangeRemoteSession'
        'Test-ExchangeShellLoaded', 'Initialize-ExchangeShell'
        'Get-Exchange2019Servers', 'Get-Exchange2013Servers', 'Get-Exchange2016Servers', 'Get-LegacyExchangeServers', 'Test-IsExchange2019Server'
        # Configuration files and resumable state
        'Import-ConfigCsv', 'Get-DeploymentState', 'Set-StepState'
    )
}
