<#
.SYNOPSIS
    Exchange 2013/2016 to 2019 Migration - single entry point: deploys Exchange 2019 next to
    Exchange 2013/2016, migrates the mailboxes in batches and validates the cut-over (26 steps).

.DESCRIPTION
    Runs the 26 steps of the catalogue in their order. Every step supports three modes:

        Inventory   Reads the current state. Changes nothing. (default)
        Simulate    Runs the changes with -WhatIf. Changes nothing.
        Apply       Applies the changes (confirmation YES, unless -Force).

    Common properties of the steps:

        Idempotence      A step run again reports AlreadyDone when the target state is in place.
        Exchange 2019    Targets come from Get-Exchange2019Servers (AdminDisplayVersion 'Version 15.2*',
        only             Edge Transport excluded): Exchange 2013/2016 and Edge servers are never
                         changed, except by Step 15 (Kerberos ASA on every CAS) and Step 25 (SCP of
                         the legacy servers), which say so in their header.
        Reports          One CSV per step, Report_GLOBAL.csv and the HTML run report, in
                         <OutputFolder>\<yyyyMMdd_HHmmss>\, with the run log and the transcripts.
        Resume           Status of every step in <OutputFolder>\DeploymentState.json; -Resume runs
                         again only the Pending, Failed or InProgress steps.
        Manual steps     Steps 18 to 26 (migration and validation) are ManualOnly: skipped by
                         -Step All -Mode Apply, run only when selected by number or name.

    Everything specific to an environment is in Configs\ (Deployment.config.psd1, DiskLayout.csv,
    DAGInfo.csv). Start with:  .\Deploy-Exchange2019.ps1 -Step List

.PARAMETER Step
    Required. Step number(s), name(s), All or List.
      -Step 1                  step 1 only
      -Step 1,2,3              steps 1, 2 and 3, in this order
      -Step PrepareMigration   by name (Step 19)
      -Step All                every step (Apply mode: without the ManualOnly steps 18-26)
      -Step List               shows the catalogue of the 26 steps and stops

.PARAMETER Mode
    Inventory (default) | Simulate | Apply.

.PARAMETER OutputFolder
    Reports, run log, transcripts and DeploymentState.json. Default: <script folder>\Reports

.PARAMETER CsvFolder
    Folder of the configuration CSV files (DiskLayout.csv, DAGInfo.csv, ...). Default: <script folder>\Configs

.PARAMETER ConfigFile
    Global configuration (product key, URLs, log paths, database layout, filters...).
    Default: <script folder>\Configs\Deployment.config.psd1

.PARAMETER Resume
    Runs again the steps that are Pending, Failed or InProgress in DeploymentState.json.
    Combine with -Step All to resume the whole deployment, or with -Step 1,2,3 to limit it.

.PARAMETER ExchangeServer
    Exchange 2019 server used for PowerShell remoting when the local Exchange snap-in is not
    loaded (run from a standard PowerShell console). Default: the local computer.

.PARAMETER Credential
    Credentials for the remoting to the Exchange 2019 servers. Default: the current account.

.PARAMETER Force
    Skips the interactive YES confirmation of Apply mode (unattended runs, wrappers).

.PARAMETER RemoteDiskFormat
    Step 04 only. Formats the disks through remoting without interactive confirmation.

.PARAMETER Follow
    Step 20 only. Follow-up mode: submits nothing, changes nothing; builds MigrationStatus.html
    and rebuilds it every -FollowInterval minutes (the browser reloads it automatically).
    Accepted only when step 20 is selected.

.PARAMETER FollowInterval
    Step 20 (and Step 21 with -FollowAfter). Minutes between two refreshes. Default 60.
    0 = build the report once, no loop.

.PARAMETER PfxPassword
    Step 02 only. Password of the PFX file (SecureString): no interactive prompt (unattended
    runs, Posh-ACME: (Get-PACertificate).PfxPass). Accepted only when step 2 is selected.

.PARAMETER BatchNumber
    Step 21 only. Batch number(s) to complete: '01', '06','07','08', or 'ALL' (every batch
    whose move requests are ready).

.PARAMETER ScheduledCompletionTime
    Step 21 only. Completion time, '2026-06-15 23:00' or '23:00'. The step sets CompleteAfter on
    every move request and resumes it, then ends: the Mailbox Replication service completes at
    that time, the console can be closed. Complete-MigrationBatch is not called in this mode
    (the batch stays Synced in the admin center; Step 22 cleans up).

.PARAMETER FollowAfter
    Step 21 only. After the completion order, continues with Step 20 -Follow (Inventory) in the
    same console to watch Synced -> Completing -> Completed. Stops with Ctrl+C. Works with
    -ScheduledCompletionTime.

.EXAMPLE
    .\Deploy-Exchange2019.ps1 -Step List
    The catalogue of the 26 steps. Changes nothing.

.EXAMPLE
    .\Deploy-Exchange2019.ps1 -Step All -Mode Inventory
    Current state of every step, on every Exchange 2019 server. Changes nothing.

.EXAMPLE
    .\Deploy-Exchange2019.ps1 -Step 1,2,3 -Mode Simulate
    Steps 1 to 3 with -WhatIf, to review the planned changes.

.EXAMPLE
    .\Deploy-Exchange2019.ps1 -Step All -Mode Apply
    Applies steps 1 to 17 (the ManualOnly steps 18-26 are skipped). Asks for YES first.

.EXAMPLE
    .\Deploy-Exchange2019.ps1 -Resume -Step All -Mode Apply
    Resumes an interrupted deployment: only the Pending, Failed and InProgress steps run.

.EXAMPLE
    .\Deploy-Exchange2019.ps1 -Step 19 -Mode Apply
    A ManualOnly step, selected explicitly (Step 19 - PrepareMigration).

.EXAMPLE
    .\Deploy-Exchange2019.ps1 -Step 21 -Mode Apply -BatchNumber 02 -ScheduledCompletionTime '23:00'
    Batch02 completes at 23:00, handled by Exchange; the step ends immediately.

.EXAMPLE
    .\Deploy-Exchange2019.ps1 -Step 21 -Mode Apply -BatchNumber 01 -FollowAfter -FollowInterval 2
    Completes Batch01, then follows it in MigrationStatus.html every 2 minutes (Ctrl+C to stop).

.EXAMPLE
    .\Deploy-Exchange2019.ps1 -Step 4 -Mode Apply -ExchangeServer EXCH201902 -Credential (Get-Credential)
    Step 04 from a standard PowerShell console, through remoting with explicit credentials.

.EXAMPLE
    $cert = Get-PACertificate ; .\Deploy-Exchange2019.ps1 -Step 2 -Mode Apply -PfxPassword $cert.PfxPass -Force
    Step 02 without any prompt, with a PFX from Posh-ACME.

.EXAMPLE
    .\Deploy-Exchange2019.ps1 -Step 20 -Follow -FollowInterval 15
    Follow-up of the migration batches, report refreshed every 15 minutes (Ctrl+C to stop).

.NOTES
    Author        : Nicolas Fabert
    Version       : 2.0.0
    Requirements  : Windows PowerShell 5.1, Exchange Management Shell 2019 (local or WinRM),
                    Organization Management, local administrator on the Exchange 2019 servers.
    Exit codes    : 0 = every selected step completed (or was skipped as ManualOnly)
                    1 = at least one step failed or is missing
                    2 = every step completed, but some actions failed (see the run report)
                    10 = shared module not found        11 = unknown step
                    12 = -PfxPassword without step 2     13 = -Follow without step 20 (or 21 + -FollowAfter)
                    14 = -FollowAfter without step 21
    Console       : EXM_ICONS = Emoji | Symbols | Ascii forces the icon style; NO_COLOR removes the colours.
    Documentation : Docs\Exchange2019Migration-Guide.html (or .md). History: CHANGELOG.md
#>

[CmdletBinding(DefaultParameterSetName='Standard')]
param(
    [Parameter(Mandatory, ParameterSetName='Standard')]
    [Parameter(Mandatory, ParameterSetName='CompleteMigration')]
    [string[]]$Step,

    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [ValidateSet('Inventory','Simulate','Apply')]
    [string]$Mode = 'Inventory',

    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [string]$OutputFolder = (Join-Path $PSScriptRoot 'Reports'),

    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [string]$CsvFolder = (Join-Path $PSScriptRoot 'Configs'),

    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [string]$ConfigFile = (Join-Path $PSScriptRoot 'Configs\Deployment.config.psd1'),

    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [switch]$Resume,

    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [switch]$RemoteDiskFormat,

    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [string]$ExchangeServer = $env:COMPUTERNAME,

    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [System.Management.Automation.PSCredential]$Credential,

    # Skips the interactive YES confirmation of Apply mode (unattended sessions).
    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [switch]$Force,

    # --- Step 02 (Certificate) only -------------------------------------------------------------
    # Password of the PFX file (SecureString), used by Step 02:
    #   - AUTO mode   : passed to Export-ExchangeCertificate then Import-ExchangeCertificate
    #   - MANUAL mode : checks the PFX placed beforehand and reads its thumbprint
    # When given, no Read-Host prompt (unattended runs, Posh-ACME wrapper).
    # Accepted only when step 2 is selected.
    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [System.Security.SecureString]$PfxPassword,

    # --- Step 20 (RunMigration, follow-up) only --------------------------------------------------
    # -Follow         : follow-up mode, submits no migration. Builds MigrationStatus.html and
    #                   rebuilds it every -FollowInterval minutes (the browser reloads it through
    #                   meta http-equiv refresh).
    # -FollowInterval : minutes between two refreshes (default 60, 0 = once, no loop). The meta
    #                   refresh of the browser uses the same value.
    # Accepted only when step 20 is selected (or step 21 with -FollowAfter).
    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [switch]$Follow,

    [Parameter(ParameterSetName='Standard')]
    [Parameter(ParameterSetName='CompleteMigration')]
    [int]$FollowInterval = 60,

    # --- Step 21 (CompleteMigration) only -------------------------------------------------------
    # Batch number(s):
    #   '01','02',...      completion of one batch
    #   '06','07','08'     completion of several batches, in sequence
    #   'ALL'              completion of every Synced batch
    [Parameter(ParameterSetName='CompleteMigration')]
    [string[]]$BatchNumber,

    # Scheduled completion time ('2026-06-15 23:00' or '23:00'). Exchange handles the timing
    # (CompleteAfter); the step ends as soon as the order is set.
    [Parameter(ParameterSetName='CompleteMigration')]
    [string]$ScheduledCompletionTime,

    # -FollowAfter: after Step 21, continues with Step 20 -Follow (Inventory) in the same session
    # to watch the batches go Synced -> Completing -> Completed in the HTML report.
    # Refresh interval: -FollowInterval (default 60 minutes). Accepted only with step 21.
    [Parameter(ParameterSetName='CompleteMigration')]
    [switch]$FollowAfter
)

$ErrorActionPreference = 'Stop'

#region -- Shared module ------------------------------------------------------------------------

$modulePath = Join-Path $PSScriptRoot 'Modules\Exchange2019.Common.psd1'
if (-not (Test-Path $modulePath)) {
    Write-Host "Shared module not found: $modulePath" -ForegroundColor Red
    exit 10
}
Import-Module $modulePath -Force -DisableNameChecking

# Default Exchange server for remoting (used by Initialize-ExchangeShell when the local snap-in is
# not available, i.e. from a standard PowerShell console).
Set-DefaultExchangeServer -Server $ExchangeServer -Credential $Credential

$Script:StepCatalog = Get-DeploymentStepCatalog
$tool = Get-ExToolInfo
$mid = [string][char]0x00B7

#endregion
#region -- Step catalogue ------------------------------------------------------------------------

function Get-StepEntry {
    param([Parameter(Mandatory)] [string]$Identifier)
    if ($Script:StepCatalog.Contains($Identifier)) {
        return $Script:StepCatalog[$Identifier]
    }
    foreach ($k in $Script:StepCatalog.Keys) {
        if ($Script:StepCatalog[$k].Name -eq $Identifier) {
            return $Script:StepCatalog[$k]
        }
    }
    return $null
}

function Show-StepList {
    # Catalogue grouped by phase: number, name, title; ManualOnly steps are marked.
    $phases = [ordered]@{
        Platform         = 'Phase 1 ' + $mid + ' Prepare the platform'
        HighAvailability = 'Phase 2 ' + $mid + ' Build high availability'
        Runtime          = 'Phase 3 ' + $mid + ' Configure runtime services'
        Migration        = 'Phase 4 ' + $mid + ' Migrate, validate and clean up (manual steps)'
    }
    Write-ExBanner -Title $tool.Name -Subtitle 'Catalogue of the 26 steps' -Details ([ordered]@{
        Usage = @('Info', '.\Deploy-Exchange2019.ps1 -Step <number|name|All> -Mode <Inventory|Simulate|Apply>')
    })
    foreach ($phase in $phases.Keys) {
        Write-ExRule $phases[$phase]
        foreach ($k in $Script:StepCatalog.Keys) {
            $e = $Script:StepCatalog[$k]
            if ($e.Phase -ne $phase) { continue }
            Write-ExItem Info ('{0,2}  {1,-24} {2}' -f $k, $e.Name, $e.Title) -Icon $e.Icon
        }
    }
    Write-ExHost ''
    Write-ExItem Dim 'Steps 18 to 26 run only when selected explicitly: -Step All -Mode Apply skips them.' -Icon 'Sub'
    Write-ExHost ''
}

#endregion
#region -- Selection of the steps ----------------------------------------------------------------

if ($Step.Count -eq 1 -and $Step[0] -ieq 'List') {
    Show-StepList
    return
}

# -Resume: the Pending/Failed/InProgress steps of the state file, limited to -Step when given.
$selectedKeys = @()
if ($Resume) {
    Initialize-DeploymentSession -Mode $Mode -OutputFolder $OutputFolder -StepName 'Resume'
    $state = Get-DeploymentState
    foreach ($k in $Script:StepCatalog.Keys) {
        $name = 'Step{0}-{1}' -f $k.PadLeft(2,'0'), $Script:StepCatalog[$k].Name
        if (-not $state.ContainsKey($name) -or $state[$name].Status -in 'Pending','Failed','InProgress') {
            $selectedKeys += $k
        }
    }
    if ($Step.Count -eq 1 -and $Step[0] -ieq 'All') { } else {
        # Limit to the -Step argument.
        $requested = @()
        foreach ($s in $Step) {
            $entry = Get-StepEntry -Identifier $s
            if ($entry) {
                $key = ($Script:StepCatalog.GetEnumerator() | Where-Object { $_.Value -eq $entry } | Select-Object -First 1).Name
                $requested += $key
            }
        }
        if ($requested.Count -gt 0) {
            $selectedKeys = $selectedKeys | Where-Object { $_ -in $requested }
        }
    }
}
elseif ($Step.Count -eq 1 -and $Step[0] -ieq 'All') {
    $selectedKeys    = $Script:StepCatalog.Keys
    $selectedFromAll = $true
}
else {
    foreach ($s in $Step) {
        $entry = Get-StepEntry -Identifier $s
        if ($null -eq $entry) {
            Write-ExItem Fail "Unknown step: $s"
            Show-StepList
            exit 11
        }
        $key = ($Script:StepCatalog.GetEnumerator() | Where-Object { $_.Value -eq $entry } | Select-Object -First 1).Name
        $selectedKeys += $key
    }
}

if (-not $selectedKeys -or @($selectedKeys).Count -eq 0) {
    Write-ExItem Warn 'No step to run.'
    return
}

# -PfxPassword is valid only when step 2 (Certificate) is selected.
if ($PfxPassword -and ($selectedKeys -notcontains '2')) {
    Write-ExItem Fail ("-PfxPassword is reserved for step 2 (Certificate), which is not selected: {0}" -f ($selectedKeys -join ', '))
    exit 12
}

# -Follow / -FollowInterval are valid only for step 20 (RunMigration), or for step 21 with
# -FollowAfter (which uses FollowInterval to drive Step 20).
if (($Follow -or $PSBoundParameters.ContainsKey('FollowInterval')) `
        -and ($selectedKeys -notcontains '20') `
        -and -not ($FollowAfter -and $selectedKeys -contains '21')) {
    Write-ExItem Fail ("-Follow / -FollowInterval are reserved for step 20 (RunMigration), or step 21 with -FollowAfter. Selection: {0}" -f ($selectedKeys -join ', '))
    exit 13
}

# -FollowAfter is reserved for step 21 (CompleteMigration).
if ($FollowAfter -and ($selectedKeys -notcontains '21')) {
    Write-ExItem Fail ("-FollowAfter is reserved for step 21 (CompleteMigration), which is not selected: {0}" -f ($selectedKeys -join ', '))
    exit 14
}

#endregion
#region -- Configuration -------------------------------------------------------------------------

$Script:Config = $null
$configText = 'not found - only the steps that need no configuration can run'
if (Test-Path $ConfigFile) {
    $Script:Config = Import-PowerShellDataFile -Path $ConfigFile
    $configText = $ConfigFile
}

#endregion
#region -- Title card and confirmation ----------------------------------------------------------

$clock = [Diagnostics.Stopwatch]::StartNew()
$run = Start-ExRun -OutputFolder $OutputFolder

# With -Step All in Apply mode the ManualOnly steps (18-26) are skipped by the loop: they are
# also left out of the displayed selection, so that nobody expects them to run.
$displayKeys = if ($Mode -eq 'Apply' -and $selectedFromAll) {
    @($selectedKeys | Where-Object { -not $Script:StepCatalog[$_].ManualOnly })
} else {
    @($selectedKeys)
}
$skippedCount = @($selectedKeys).Count - $displayKeys.Count
$stepsLine = $displayKeys -join ', '
if ($skippedCount -gt 0) { $stepsLine += "   (+$skippedCount manual-only steps skipped)" }
$modeText = @{ Inventory = 'Inventory - reads the current state, changes nothing'; Simulate = 'Simulate - runs the changes with -WhatIf, changes nothing'; Apply = 'Apply - changes the servers' }[$Mode]

Write-ExBanner -Title $tool.Name -Subtitle 'Deployment, mailbox migration and validation' -Details ([ordered]@{
    Mode    = @('Mode', $modeText)
    Steps   = @('Plan', $stepsLine)
    Server  = @('Server', ('{0}{1}' -f $ExchangeServer, $(if ($Credential) { "  ($($Credential.UserName))" } else { '' })))
    Config  = @('File', $configText)
    Reports = @('Folder', $run.Folder)
    Log     = @('Log', $run.Log)
})
if ($Resume) { Write-ExItem Info 'Resume: only the Pending, Failed and InProgress steps of DeploymentState.json run.' -Icon 'Sub' }
if (-not $Script:Config) { Write-ExItem Warn "Configuration file not found: $ConfigFile" }

if ($Mode -eq 'Apply') {
    Write-ExHost ''
    # Step 25 targets the Exchange 2013/2016 Client Access servers on purpose
    # (Set-ClientAccessService -AutoDiscoverServiceInternalUri $null): say so.
    if ($selectedKeys -contains '25') {
        Write-ExItem Warn 'APPLY MODE - real changes on Exchange 2013/2016: Step 25 clears AutoDiscoverServiceInternalUri on the legacy servers.'
    } else {
        Write-ExItem Warn 'APPLY MODE - real changes on the Exchange 2019 servers. Exchange 2013/2016 servers are not changed.'
    }
    if ($Force) {
        Write-ExItem Skip '-Force: interactive confirmation skipped.'
    } else {
        $confirm = Read-Host "      Type YES to run in Apply mode"
        if ($confirm -ne 'YES') {
            Write-ExItem Warn 'Cancelled by the user.'
            Stop-ExRun
            exit 0
        }
    }
}

#endregion
#region -- Step loop -----------------------------------------------------------------------------

$globalSummary = @()
$runKeys = @($selectedKeys | Where-Object { -not ($Mode -eq 'Apply' -and $selectedFromAll -and $Script:StepCatalog[$_].ManualOnly) })
$index = 0

foreach ($key in $selectedKeys) {
    $entry = $Script:StepCatalog[$key]
    $stepName = 'Step{0}-{1}' -f $key.PadLeft(2,'0'), $entry.Name
    $stepFile = Join-Path $PSScriptRoot ('Steps\' + $entry.File)

    # Manual steps: skipped with -Step All in Apply mode (not in Inventory/Simulate, not when selected explicitly).
    if ($Mode -eq 'Apply' -and $selectedFromAll -and $entry.ManualOnly) {
        Write-ExItem Skip ('{0,2}  {1,-24} skipped (ManualOnly - select it explicitly to run it)' -f $key, $entry.Name)
        $globalSummary += [PSCustomObject]@{ Step = $stepName; Status = 'ManualOnly'; Mode = $Mode; Report = '-' }
        continue
    }

    if (-not (Test-Path $stepFile)) {
        Write-ExItem Fail "Step file not found: $stepFile"
        $globalSummary += [PSCustomObject]@{ Step = $stepName; Status = 'MissingFile'; Mode = $Mode; Report = '-' }
        continue
    }

    $index++
    Set-ExStepContext -Index $index -Total $runKeys.Count -Icon $entry.Icon
    Initialize-DeploymentSession -Mode $Mode -OutputFolder $OutputFolder -StepName $stepName
    Set-StepState -StepName $stepName -Status InProgress
    $stepClock = [Diagnostics.Stopwatch]::StartNew()

    try {
        # Dot-source the step (it defines Invoke-Step), then call it.
        . $stepFile

        # Global variables read by Step 21 (CompleteMigration).
        if ($BatchNumber)             { $Global:BatchNumber             = $BatchNumber }             else { Remove-Variable -Scope Global -Name BatchNumber             -ErrorAction SilentlyContinue }
        if ($ScheduledCompletionTime) { $Global:ScheduledCompletionTime = $ScheduledCompletionTime } else { Remove-Variable -Scope Global -Name ScheduledCompletionTime -ErrorAction SilentlyContinue }

        # Global variable read by Step 02 (Certificate): no PFX prompt.
        if ($PfxPassword)             { $Global:PfxPassword             = $PfxPassword }             else { Remove-Variable -Scope Global -Name PfxPassword             -ErrorAction SilentlyContinue }

        # Global variables read by Step 20 (RunMigration): follow-up mode.
        if ($Follow) {
            $Global:FollowMigration         = $true
            $Global:FollowMigrationInterval = [int]$FollowInterval
        } else {
            Remove-Variable -Scope Global -Name FollowMigration         -ErrorAction SilentlyContinue
            Remove-Variable -Scope Global -Name FollowMigrationInterval -ErrorAction SilentlyContinue
        }

        # Step 04 (Disks): remote formatting.
        $Global:RemoteDiskFormat = [bool]$RemoteDiskFormat

        $stepParams = @{
            Mode         = $Mode
            OutputFolder = $OutputFolder
            CsvFolder    = $CsvFolder
            Config       = $Script:Config
            Credential   = $Credential
        }

        $rc = Invoke-Step @stepParams
        # Invoke-Action writes its status ('Inventoried', 'Skipped', ...) to the pipeline, so $rc
        # may be an array. The real return code is always the last element (return N).
        $actualRc = $rc | Select-Object -Last 1

        $reportCsv = Save-Report -StepName $stepName

        if ($actualRc -eq 0) {
            Set-StepState -StepName $stepName -Status Completed -Detail $reportCsv
            Write-ExItem Ok ('{0} completed in {1}' -f $stepName, (Format-ExDuration $stepClock.Elapsed.TotalSeconds)) -Icon 'Clock'
            $globalSummary += [PSCustomObject]@{ Step = $stepName; Status = 'OK'; Mode = $Mode; Report = $reportCsv }
        } else {
            Set-StepState -StepName $stepName -Status Failed -Detail "ReturnCode=$actualRc"
            Write-ExItem Fail ('{0} ended with return code {1} after {2}' -f $stepName, $actualRc, (Format-ExDuration $stepClock.Elapsed.TotalSeconds))
            $globalSummary += [PSCustomObject]@{ Step = $stepName; Status = 'Failed'; Mode = $Mode; Report = $reportCsv }
        }
    }
    catch {
        Write-Log "Fatal error in step $stepName - $($_.Exception.Message)" -Level Error
        Set-StepState -StepName $stepName -Status Failed -Detail $_.Exception.Message
        $globalSummary += [PSCustomObject]@{ Step = $stepName; Status = 'Failed'; Mode = $Mode; Report = '-' }
    }
    finally {
        Stop-DeploymentSession
        Reset-Report
    }
}

# --- -FollowAfter: continue with Step 20 -Follow after a successful Step 21 --------------------
# Step 20 loops (while $true) and keeps the console until Ctrl+C, on purpose: the operator
# watches Complete-MigrationBatch until the batches are Completed.
if ($FollowAfter) {
    $step21Result = $globalSummary | Where-Object { $_.Step -like 'Step21-*' } | Select-Object -First 1
    if ($step21Result -and $step21Result.Status -eq 'OK') {
        Write-ExStep -Pill 'Follow' -Title ('Step 20 ' + $mid + ' Migration follow-up after the completion (-FollowAfter)') -Icon 'Search' -Note 'Inventory'

        $step20File = Join-Path $PSScriptRoot 'Steps\Step20-RunMigration.ps1'
        if (Test-Path $step20File) {
            $Global:FollowMigration         = $true
            $Global:FollowMigrationInterval = [int]$FollowInterval

            # Focus on one or several batches: every numeric -BatchNumber becomes 'BatchNN' for
            # Step 20. 'ALL' or nothing = no filter.
            $focusBatches = @()
            if ($BatchNumber) {
                $focusBatches = @($BatchNumber |
                    Where-Object { $_ -and $_ -ne 'ALL' -and $_ -match '^\d+$' } |
                    ForEach-Object { 'Batch{0:D2}' -f [int]$_ })
            }
            if ($focusBatches.Count -eq 1) {
                $Global:FollowBatchFilter = $focusBatches[0]
                Write-ExItem Info "Focus on one batch: $Global:FollowBatchFilter" -Icon 'Target'
            } elseif ($focusBatches.Count -gt 1) {
                $Global:FollowBatchFilter = $focusBatches
                Write-ExItem Info ("Focus on {0} batches: {1}" -f $focusBatches.Count, ($focusBatches -join ', ')) -Icon 'Target'
            } else {
                Remove-Variable -Scope Global -Name FollowBatchFilter -ErrorAction SilentlyContinue
            }
            $followStep = 'Step20-RunMigration'

            Initialize-DeploymentSession -Mode Inventory -OutputFolder $OutputFolder -StepName $followStep
            try {
                . $step20File
                $followParams = @{
                    Mode         = 'Inventory'
                    OutputFolder = $OutputFolder
                    CsvFolder    = $CsvFolder
                    Config       = $Script:Config
                    Credential   = $Credential
                }
                Invoke-Step @followParams | Out-Null
            } catch {
                Write-Log "Error in the follow-up after the completion: $($_.Exception.Message)" -Level Error
            } finally {
                Stop-DeploymentSession
                Reset-Report
                Remove-Variable -Scope Global -Name FollowBatchFilter -ErrorAction SilentlyContinue
            }
        } else {
            Write-ExItem Warn "Step 20 file not found: $step20File"
        }
    } else {
        Write-ExItem Warn ("-FollowAfter: Step 21 did not succeed (Status={0}); no follow-up." -f $step21Result.Status)
    }
}

#endregion
#region -- Final summary -------------------------------------------------------------------------

$ok        = @($globalSummary | Where-Object Status -eq 'OK').Count
$failed    = @($globalSummary | Where-Object { $_.Status -in 'Failed', 'MissingFile' }).Count
$manual    = @($globalSummary | Where-Object Status -eq 'ManualOnly').Count
$summary   = Get-ExRunSummary

$stepsText = '{0} completed {1} {2} failed' -f $ok, $mid, $failed
if ($manual) { $stepsText += (' {0} {1} manual-only skipped' -f $mid, $manual) }
$failedSteps = @($globalSummary | Where-Object { $_.Status -in 'Failed', 'MissingFile' } | ForEach-Object Step)
$actionsText = foreach ($s in 'Success', 'AlreadyDone', 'Inventoried', 'Simulated', 'Skipped', 'Failed') {
    if ($summary.Counts[$s]) { '{0} {1}' -f $summary.Counts[$s], @{ Success = 'changed'; AlreadyDone = 'already done'; Inventoried = 'inventoried'; Simulated = 'simulated'; Skipped = 'skipped'; Failed = 'failed' }[$s] }
}

$failedActions = [int]$summary.Counts['Failed']
$runStatus = if ($failed) { 'Fail' } elseif ($failedActions) { 'Warn' } else { 'Ok' }
$values = [ordered]@{ Steps = @($runStatus, $stepsText) }
if ($failedSteps) { $values['Failed'] = @('Fail', ($failedSteps -join ', ')) }
$values['Actions']  = @('Report', $(if ($actionsText) { $actionsText -join (" $mid ") } else { 'none' }))
$values['Mode']     = @('Mode', $Mode)
$values['Report']   = @('File', $(if ($summary.Entries) { $summary.Html } else { 'no action recorded' }))
$values['Folder']   = @('Folder', $run.Folder)
$values['Duration'] = @('Clock', (Format-ExDuration $clock.Elapsed.TotalSeconds))
$values['Log']      = @('Log', $run.Log)

$title = @{ Fail = 'Run finished with errors'; Warn = 'Run complete, some actions failed'; Ok = 'Run complete' }[$runStatus]
Write-ExSummary -Title $title -Values $values -Status $runStatus
Stop-ExRun

if ($failed) { exit 1 }
if ($failedActions) { exit 2 }
exit 0

#endregion
