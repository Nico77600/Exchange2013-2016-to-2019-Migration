#Requires -Version 5.1
<#
.SYNOPSIS
    Regenerates the screenshots of the guide (package\Docs\images\*.png) from synthetic data.

.DESCRIPTION
    No Exchange server is needed. The tool:

      1. builds the HTML reports with anonymized sample data (contoso, EXCH2019xx), in Windows
         PowerShell 5.1 - the shell of the framework - with the real report functions of the
         module and of the steps;
      2. records the console of "-Step List" and of a short demonstration run (the real console
         functions of the module, emoji style as in Windows Terminal) through EXM_CAPTURE_FILE,
         and renders them as a terminal window;
      3. takes the screenshots with Microsoft Edge in headless mode;
      4. renders the graphics of the GitHub README - banner, design principles, how it works,
         the 26 steps, migration runbook - with the CSS and the icons of the HTML guide, in a
         light and a dark version (2x resolution). The README cannot show the custom blocks of
         the guide (cards, flow): it shows these images instead, chosen by <picture>.

    Guide images: console-list.png, console-run.png, report-run.png,
    report-migration-status.png, report-protocol-logs.png.
    README images: readme-<name>-light.png and readme-<name>-dark.png.

.PARAMETER OutputFolder
    Default: package\Docs\images next to the tools folder.

.PARAMETER Images
    All (default), Guide (screenshots of the guide only) or Readme (README graphics only).
    The README graphics read Docs\Exchange2019Migration-Guide.md (flow and cards blocks),
    the built Docs\Exchange2019Migration-Guide.html (CSS) and the step catalogue of the module.

.PARAMETER KeepWork
    Keeps the work folder (sample HTML files, console captures) and shows its path.

.EXAMPLE
    .\tools\New-DocumentationImages.ps1
    Then .\tools\Build-Documentation.ps1 to embed the new images in the HTML guide.

.EXAMPLE
    .\tools\New-DocumentationImages.ps1 -Images Readme
    Only the README graphics, after a change of the guide or of the step catalogue.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    Part of : Exchange 2013/2016 to 2019 Migration (repository tool, not in the package)
#>
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [ValidateSet('All', 'Guide', 'Readme')][string]$Images = 'All',
    [switch]$KeepWork,
    # Internal: the parts that run in a separate Windows PowerShell 5.1 process.
    [ValidateSet('', 'Samples', 'ConsoleRun')][string]$Part = '',
    [string]$Work
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$packageRoot = Join-Path $root 'package'

#region Part 1 - sample reports (Windows PowerShell 5.1) --------------------------------------------
function Invoke-SampleReports {
    param([string]$Root, [string]$Out)
    New-Item -ItemType Directory -Path $Out -Force | Out-Null
    Import-Module (Join-Path (Join-Path $Root 'package') 'Modules\Exchange2019.Common.psd1') -Force -DisableNameChecking
    $mod = Get-Module Exchange2019.Common
    # ---- Run report ----------------------------------------------------------------------------------
    & $mod { $Script:Mode = 'Apply'; $Script:OutputFolder = 'D:\Tools\Deploy-Exchange2019\Reports\20261001_101500'; $Script:TimestampRun = '20261001_101500'; $Script:RunStart = [datetime]'2026-10-01 10:15:00'; $Script:DefaultExchangeServer = 'EXCH201901'; Enable-ExConsoleCapture -Quiet }
    $servers = 'EXCH201901', 'EXCH201902', 'EXCH201903', 'EXCH201904'
    foreach ($s in $servers) { Add-Report -Step '01-LicenseKey' -Target $s -Action 'ApplyLicense' -Status AlreadyDone -BeforeValue 'Edition=Standard; AdminDisplayVersion=Version 15.2 (Build 1544.4); IsExchangeTrialEdition=False' -Detail 'Set-ExchangeServer -ProductKey' }
    foreach ($s in $servers) {
        foreach ($v in 'Autodiscover', 'OWA', 'ECP', 'EWS', 'ActiveSync', 'OAB', 'MAPI', 'OutlookAnywhere') {
            $st = if ($s -eq 'EXCH201904' -and $v -eq 'MAPI') { 'Failed' } elseif ($v -in 'OWA', 'ECP') { 'AlreadyDone' } else { 'Success' }
            Add-Report -Step '03-VirtualDirectories' -Target $s -Action "Configure $v" -Status $st -Phase 'After' -BeforeValue "InternalUrl=https://$($s.ToLower()).contoso.local/$v" -AfterValue "InternalUrl=https://mail.contoso.com/$v" -Detail 'URLs and authentication copied from EXCH201601' -ErrorMessage $(if ($st -eq 'Failed') { 'The WinRM client cannot process the request (timeout).' } else { '' })
        }
    }
    foreach ($db in 'DB01', 'DB02', 'DB03', 'DB04') { Add-Report -Step '10-CreateDatabases' -Target $db -Action 'New-MailboxDatabase' -Status Success -Phase 'After' -BeforeValue '<absent>' -AfterValue "EdbFilePath=M:\Databases\$db\$db.edb" -Detail "Server=EXCH20190$($db.Substring(3))" }
    foreach ($s in $servers) { Add-Report -Step '14-EventLogSize' -Target $s -Action 'Application log size' -Status Success -Phase 'After' -BeforeValue 'MaximumSizeInBytes=20971520' -AfterValue 'MaximumSizeInBytes=1073741824' }
    Add-Report -Step '15-KerberosASA' -Target 'EXCH201601' -Action 'ASA credential' -Status Skipped -Detail 'Legacy server unreachable, retry later'
    New-HtmlReport -Path (Join-Path $Out 'Report_GLOBAL_Apply.html')

    # ---- Protocol-log analysis (Step 26) -------------------------------------------------------------
    . (Join-Path (Join-Path $Root 'package') 'Steps\Step26-AnalyzeProtocolLogs.ps1')
    $sv = [ordered]@{ EXCH201601 = '2016'; EXCH201602 = '2016'; EXCH201901 = '2019'; EXCH201902 = '2019'; EXCH201903 = '2019'; EXCH201904 = '2019' }
    $vd = @('Mapi', 'OWA', 'EWS', 'OAB', 'ECP', 'RPC', 'Autodiscover', 'Microsoft-Server-ActiveSync', 'PowerShell', '<Other>')
    $iis = foreach ($n in $sv.Keys) {
        $legacy = $sv[$n] -eq '2016'
        $by = @{ Mapi = $(if ($legacy) { 12 } else { 18450 }); OWA = $(if ($legacy) { 0 } else { 2210 }); EWS = $(if ($legacy) { 3 } else { 6120 }); Autodiscover = $(if ($legacy) { 41 } else { 3380 }); 'Microsoft-Server-ActiveSync' = $(if ($legacy) { 0 } else { 1975 }) }
        $tot = ($by.Values | Measure-Object -Sum).Sum
        [pscustomobject]@{ Server = $n; Version = $sv[$n]; Path = 'D:\Logs\IIS\DefaultWebSite'; TotalRequests = $tot; TotalRaw = $tot * 3; ExcludedUri = 120; ExcludedUser = 840; ExcludedUA = 1530; ExcludedStatus = 760
            ByVdir = $by; ByVdirIPs = @{ Mapi = @{ '10.10.20.31' = 410; '10.10.20.47' = 388; '10.10.21.12' = 120 }; Autodiscover = @{ '10.10.20.31' = 30; '10.10.22.8' = 11 } }; TopClientIps = '10.10.20.31=440 ; 10.10.20.47=388'; TopStatus = '200=9120 ; 401=760'; Note = '' }
    }
    $smtp = foreach ($n in $sv.Keys) {
        $legacy = $sv[$n] -eq '2016'
        foreach ($k in 'HubReceive', 'FERecv', 'HubSend') {
            $valid = if ($legacy) { 0 } else { @{ HubReceive = 512; FERecv = 1404; HubSend = 388 }[$k] }
            [pscustomobject]@{ Server = $n; Version = $sv[$n]; Kind = $k; Path = 'D:\Logs\Transport'; Sessions = $valid; SessionsTotal = $valid + 2876; ExcludedSystem = 410; ExcludedNoMailFrom = 2466; UniqueRemoteIps = 6; TopRemoteIps = '10.10.30.5=210 ; 10.10.30.6=180'; ByConnector = @{ "$n\Default Frontend $n" = $valid; "$n\Relay $n" = [int]($valid / 4) }; Note = '' }
        }
    }
    $matrix = @{}; foreach ($v in $vd) { $matrix[$v] = @{ '2016' = 0; '2019' = 0 } }
    foreach ($s in $iis) { foreach ($kv in $s.ByVdir.GetEnumerator()) { $matrix[$kv.Key][$s.Version] += $kv.Value } }
    $iisByVer = @{ '2016' = (($iis | Where-Object Version -eq '2016').TotalRequests | Measure-Object -Sum).Sum; '2019' = (($iis | Where-Object Version -eq '2019').TotalRequests | Measure-Object -Sum).Sum }
    $smtpMatrix = @{ Receive = @{ '2016' = 0; '2019' = 3 * (512 + 1404) + (512 + 1404) }; Send = @{ '2016' = 0; '2019' = 4 * 388 } }
    $smtpByVer = @{ '2016' = 0; '2019' = $smtpMatrix.Receive['2019'] + $smtpMatrix.Send['2019'] }
    New-ProtocolLogsHtml -OutputPath (Join-Path $Out 'ProtocolLogsAnalysis.html') -Hours 24 -Since ([datetime]'2026-09-30 10:15') -ServerVersion $sv -VDirsOrder $vd -VersionsPresent @('2016', '2019') `
        -IisStats $iis -SmtpStats $smtp -IisMatrix $matrix -SmtpMatrix $smtpMatrix -IisByVer $iisByVer -SmtpByVer $smtpByVer `
        -IisLegacy $iisByVer['2016'] -Iis2019 $iisByVer['2019'] -SmtpLegacy 0 -Smtp2019 $smtpByVer['2019'] -AutodLegacy 82 -Autod2019 13520 -LegacyLabel '2016'

    # ---- Nested report functions (steps 18, 19, 20): extracted from the step files ------------------
    function Import-NestedFunction([string]$File, [string]$Name) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($File, [ref]$null, [ref]$null)
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
        . ([scriptblock]::Create($fn.Extent.Text.Replace("function $Name", "function global:$Name")))
    }
    function global:Get-ExchangeServer { param($Identity) [pscustomobject]@{ AdminDisplayVersion = $(if ($Identity -like 'EXCH2016*') { 'Version 15.1 (Build 2507.6)' } else { 'Version 15.2 (Build 1544.4)' }) } }
    function global:Get-MoveRequest { param($Identity) if ($Identity -like '*Audit*') { [pscustomobject]@{ Status = 'InProgress' } } }

    Import-NestedFunction (Join-Path (Join-Path $Root 'package') 'Steps\Step18-SystemMailboxMigration.ps1') 'New-SystemMailboxPlanHtml'
    $sys = @()
    $i = 0
    foreach ($t in 'ArbitrationMailbox', 'ArbitrationMailbox', 'ArbitrationMailbox', 'ArbitrationMailbox', 'ArbitrationMailbox', 'AuditLogMailbox', 'AuxAuditLogMailbox', 'DiscoveryMailbox') {
        $i++; $srv = if ($i -le 3) { 'EXCH201901' } else { 'EXCH201601' }
        $sys += [pscustomobject]@{ Name = "SystemMailbox{$([guid]::NewGuid())}"; DisplayName = @('Microsoft Exchange', 'Microsoft Exchange Approval Assistant', 'Microsoft Exchange Federation Mailbox', 'Migration', 'E4E Encryption Store', 'Microsoft Exchange Audit', 'Microsoft Exchange Aux Audit', 'Discovery Search Mailbox')[$i - 1]
            RecipientTypeDetails = $t; ServerName = $srv; Database = $(if ($srv -eq 'EXCH201901') { 'DB01' } else { 'DB-LEG-01' }); Identity = "$t-$i" }
    }
    New-SystemMailboxPlanHtml -Mailboxes $sys -TargetDbCount 4 -OutputPath (Join-Path $Out 'SystemMailboxPlan.html') -ReportMode 'Simulate'

    Import-NestedFunction (Join-Path (Join-Path $Root 'package') 'Steps\Step19-PrepareMigration.ps1') 'New-MigrationPlanHtml'
    $names = 'Adele Vance', 'Alex Wilber', 'Diego Siciliani', 'Grady Archie', 'Henrietta Mueller', 'Isaiah Langer', 'Johanna Lorenz', 'Joni Sherman', 'Lee Gu', 'Lidia Holloway', 'Lynne Robbins', 'Megan Bowen', 'Miriam Graham', 'Nestor Wilke', 'Patti Fernandez', 'Pradeep Gupta', 'Accounting', 'Room A', 'Projector'
    $data = New-Object System.Collections.Generic.List[object]; $batches = @{}
    for ($k = 0; $k -lt $names.Count; $k++) {
        $bn = ($k % 4) + 1
        $type = if ($names[$k] -eq 'Accounting') { 'SharedMailbox' } elseif ($names[$k] -eq 'Room A') { 'RoomMailbox' } elseif ($names[$k] -eq 'Projector') { 'EquipmentMailbox' } else { 'UserMailbox' }
        $size = 800 + ($k * 397) % 5200
        $alias = ($names[$k] -replace ' ', '.').ToLower()
        $data.Add([pscustomobject]@{ DisplayName = $names[$k]; Alias = $alias; RecipientTypeDetails = $type; SourceDatabase = "DB-LEG-0$((($k % 2) + 1))"; SizeMB = $size; Type = 'Primary'; HasArchive = ($k % 3 -eq 0); Batch = $bn })
        if ($k % 3 -eq 0) { $data.Add([pscustomobject]@{ DisplayName = "$($names[$k]) (Archive)"; Alias = $alias; RecipientTypeDetails = 'Archive'; SourceDatabase = 'DB-ARCH-01'; SizeMB = [int]($size / 2); Type = 'Archive'; HasArchive = $false; Batch = $bn }) }
    }
    foreach ($bn in 1..4) { $items = @($data | Where-Object Batch -eq $bn); $batches[$bn] = [pscustomobject]@{ Number = $bn; SizeMB = ($items | Measure-Object SizeMB -Sum).Sum; Count = @($items | Where-Object Type -eq 'Primary').Count } }
    New-MigrationPlanHtml -Data $data -Batches $batches -OutputPath (Join-Path $Out 'MigrationPlan.html') -ReportMode 'Apply' -BatchCount 4

    Import-NestedFunction (Join-Path (Join-Path $Root 'package') 'Steps\Step20-RunMigration.ps1') 'New-MigrationStatusHtml'
    $stats = foreach ($bn in 1..4) {
        $users = foreach ($d in ($data | Where-Object { $_.Batch -eq $bn -and $_.Type -eq 'Primary' })) {
            $st = switch ($bn) { 1 { 'Completed' } 2 { 'Synced' } 3 { if ($d.DisplayName -like 'L*') { 'Failed' } else { 'Syncing' } } default { 'Provisioning' } }
            $pct = switch ($st) { 'Completed' { 100 } 'Synced' { 95 } 'Syncing' { 40 + ($d.SizeMB % 40) } 'Failed' { 5 } default { 0 } }
            $sd = if ($bn -eq 2 -and $d.DisplayName -like 'H*') { 'StalledDueToMRS_Quarantined' } else { '' }
            [pscustomobject]@{ Identity = "$($d.Alias)@contoso.com"; Status = $st; StatusDetail = $sd; PercentageComplete = $pct; TargetDatabase = "DB0$bn"; LastUpdateTimestamp = [datetime]'2026-10-01 09:42'
                FailureType = $(if ($st -eq 'Failed') { 'TooManyBadItemsPermanentException' } else { '' }); Error = $(if ($st -eq 'Failed') { 'The job encountered too many bad items (21). It has been stopped.' } elseif ($sd) { 'The move request is quarantined by the Mailbox Replication service.' } else { '' }); SkippedItemCount = 0 }
        }
        $bs = switch ($bn) { 1 { 'Completed' } 2 { 'Synced' } 3 { 'Syncing' } default { 'Created' } }
        [pscustomobject]@{ Name = ('Batch{0:D2}' -f $bn); Status = $bs; Users = @($users) }
    }
    New-MigrationStatusHtml -Stats $stats -OutputPath (Join-Path $Out 'MigrationStatus.html') -ReportMode 'Inventory' -Title 'Migration status' -AutoRefreshSeconds 900

    . (Join-Path (Join-Path $Root 'package') 'Steps\Step24-HealthCheckerReport.ps1')
    $issues = @{
        EXCH201901 = @([pscustomobject]@{ Severity = 'Warning'; CategoryName = 'Operating System Information'; CategoryOrder = 2; Name = 'Visual C++ 2012 x64'; DisplayValue = 'Redistributable is outdated' }
                       [pscustomobject]@{ Severity = 'Error'; CategoryName = 'Security Settings'; CategoryOrder = 5; Name = 'TLS 1.2 - SystemDefaultTlsVersions'; DisplayValue = 'Error: SystemDefaultTlsVersions is not set to the recommended value' })
        EXCH201902 = @([pscustomobject]@{ Severity = 'Warning'; CategoryName = 'Frequent Configuration Issues'; CategoryOrder = 4; Name = 'Open Relay Wild Card Domain'; DisplayValue = 'Error: Found an accepted domain with a wild card' })
        EXCH201903 = @()
        EXCH201904 = @([pscustomobject]@{ Severity = 'Warning'; CategoryName = 'Processor/Hardware Information'; CategoryOrder = 3; Name = 'Physical Memory'; DisplayValue = '64 GB - more than 128 GB is recommended for 4 databases' })
    }
    New-HealthCheckerCustomHtmlReport -ServerIssues $issues -Path (Join-Path $Out 'HealthChecker-Issues.html') -OfficialHtmlPath 'D:\Reports\ExchangeAllServersReport.html' -XmlDirectory 'D:\Reports\Step24-HealthChecker\20261001_101500'
}
#endregion

#region Part 2 - demonstration run for the console screenshot (Windows PowerShell 5.1) -----------------
function Invoke-ConsoleRun {
    param([string]$Root, [string]$Out)
    Import-Module (Join-Path (Join-Path $Root 'package') 'Modules\Exchange2019.Common.psd1') -Force -DisableNameChecking
    $mod = Get-Module Exchange2019.Common
    $tool = Get-ExToolInfo
    New-Item -ItemType Directory -Path $Out -Force | Out-Null
    $base = 'D:\Tools\Deploy-Exchange2019'
    & $mod { param($o) $Script:Mode = 'Inventory'; $Script:OutputFolder = $o; $Script:DefaultExchangeServer = 'EXCH201901' } $Out

    Write-ExHost 'PS D:\Tools\Deploy-Exchange2019> ' -Color 'Cyan' -NoNewline
    Write-ExHost '.\Deploy-Exchange2019.ps1 -Step 3,10 -Mode Inventory'
    Write-ExBanner -Title $tool.Name -Subtitle 'Deployment, mailbox migration and validation' -Details ([ordered]@{
        Mode    = @('Mode', 'Inventory - reads the current state, changes nothing')
        Steps   = @('Plan', '3, 10')
        Server  = @('Server', 'EXCH201901')
        Config  = @('File', "$base\Configs\Deployment.config.psd1")
        Reports = @('Folder', "$base\Reports\20261001_101500")
        Log     = @('Log', "$base\Reports\20261001_101500\Deploy-Exchange2019.log")
    })

    $servers = 'EXCH201901', 'EXCH201902', 'EXCH201903', 'EXCH201904'
    Set-ExStepContext -Index 1 -Total 2 -Icon 'Globe'
    Write-StepBanner -StepName '03' -Title 'Virtual directories' -Mode 'Inventory' -Actions @(
        'Read the URLs and authentication of the source server (EXCH201601)',
        'Compare every virtual directory of the Exchange 2019 servers with the target values',
        'Apply mode only: Set-*VirtualDirectory, then one IIS reset per server')
    Write-Log 'Reference URLs (source: EXCH201601 2016)' -Level Info
    foreach ($s in $servers[0..1]) {
        foreach ($v in 'OWA', 'EWS', 'MAPI') {
            Add-Report -Step '03-VirtualDirectories' -Target $s -Action "$v virtual directory" -Status Inventoried -Phase 'Before' -BeforeValue "InternalUrl=https://mail.contoso.com/$($v.ToLower())" -Detail 'target https://mail.contoso.com'
        }
    }
    Add-Report -Step '03-VirtualDirectories' -Target 'EXCH201903' -Action 'MAPI virtual directory' -Status Failed -ErrorMessage 'The WinRM client cannot process the request (timeout).'
    [void](Save-Report -StepName 'Step03-VirtualDirectories')
    Write-ExItem Ok 'Step03-VirtualDirectories completed in 14.2 s' -Icon 'Clock'
    Reset-Report

    Set-ExStepContext -Index 2 -Total 2 -Icon 'Database'
    Write-StepBanner -StepName '10' -Title 'Mailbox databases' -Mode 'Inventory' -Actions @(
        'Generate DB01 to DB04 from DatabaseLayout (round-robin on the Exchange 2019 servers)',
        'Create the folders M:\Databases\<DB>\Logs, then New-MailboxDatabase and the default properties')
    foreach ($i in 1..4) {
        Add-Report -Step '10-CreateDatabases' -Target ('DB0{0}' -f $i) -Action 'New-MailboxDatabase' -Status Inventoried -Phase 'Before' -BeforeValue '<absent>' -Detail ('Server=EXCH20190{0}, M:\Databases\DB0{0}\DB0{0}.edb' -f $i)
    }
    [void](Save-Report -StepName 'Step10-CreateDatabases')
    Write-ExItem Ok 'Step10-CreateDatabases completed in 6.8 s' -Icon 'Clock'

    $mid = [string][char]0x00B7
    Write-ExSummary -Title 'Run complete, some actions failed' -Status Warn -Values ([ordered]@{
        Steps    = @('Warn', "2 completed $mid 0 failed")
        Actions  = @('Report', "10 inventoried $mid 1 failed")
        Mode     = @('Mode', 'Inventory')
        Report   = @('File', "$base\Reports\20261001_101500\Report_GLOBAL_Inventory.html")
        Duration = @('Clock', '21.4 s')
        Log      = @('Log', "$base\Reports\20261001_101500\Deploy-Exchange2019.log")
    })
}
#endregion

#region Part 3 - README graphics -------------------------------------------------------------------
# GitHub renders Markdown only: the custom blocks of the guide (cards, flow) and its theme are
# lost. They are rendered here as images, with the CSS of the built HTML guide and the icons of
# Build-Documentation.ps1, so that the README and the guide always look the same.

function ConvertTo-ReadmeInline([string]$Text) {
    # Inline Markdown of a guide block (code, bold, italic) -> HTML.
    $h = [System.Net.WebUtility]::HtmlEncode($Text.Trim())
    $h = [regex]::Replace($h, '`([^`]+)`', '<code>$1</code>')
    $h = [regex]::Replace($h, '\*\*([^*]+)\*\*', '<strong>$1</strong>')
    return [regex]::Replace($h, '(?<![\w*])\*([^*\s][^*]*)\*(?![\w*])', '<em>$1</em>')
}

function Get-ReadmeAssets {
    param([string]$Root)
    $builder = Join-Path $Root 'tools\Build-Documentation.ps1'
    $guideHtml = Join-Path (Join-Path $Root 'package') 'Docs\Exchange2019Migration-Guide.html'
    $guideMd = Join-Path (Join-Path $Root 'package') 'Docs\Exchange2019Migration-Guide.md'
    if (-not (Test-Path $guideHtml)) { throw 'package\Docs\Exchange2019Migration-Guide.html not found: run tools\Build-Documentation.ps1 first (it holds the CSS of the graphics).' }
    # Icons: the $Icons table of the documentation builder, read without running the builder.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($builder, [ref]$null, [ref]$null)
    $assign = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$Icons' }, $true)
    if (-not $assign) { throw "Icon table not found in $builder." }
    $md = [IO.File]::ReadAllText($guideMd) -replace "`r`n", "`n"
    $blocks = foreach ($m in [regex]::Matches($md, '(?s)```(flow|cards)\n(.*?)\n```')) {
        [pscustomobject]@{ Kind = $m.Groups[1].Value; Lines = @($m.Groups[2].Value -split "`n" | Where-Object { $_.Trim() }) }
    }
    [pscustomobject]@{
        Icons   = & ([scriptblock]::Create($assign.Right.Extent.Text))
        Css     = [regex]::Match([IO.File]::ReadAllText($guideHtml), '(?s)<style>(.*?)</style>').Groups[1].Value
        Version = [regex]::Match($md, '(?m)^version:\s*(\S+)').Groups[1].Value
        Flows   = @($blocks | Where-Object Kind -eq 'flow')
        Cards   = @($blocks | Where-Object Kind -eq 'cards')
    }
}

function Get-ReadmeIcon([string]$Name, [string]$Class = 'icon') {
    $path = $assets.Icons[$Name]; if (-not $path) { $path = $assets.Icons['info'] }
    "<svg class=""$Class"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round"" stroke-linejoin=""round"">$path</svg>"
}

function ConvertTo-ReadmeFlow([string[]]$Lines, [switch]$Vertical) {
    # Vertical: the nodes are stacked, icon on the left, with a downward arrow and its label.
    $items = foreach ($l in $Lines) {
        $icon, $title, $sub = $l.Split('|', 3).ForEach({ $_.Trim() })
        $title = [System.Net.WebUtility]::HtmlEncode($title); $sub = [System.Net.WebUtility]::HtmlEncode($sub)
        if ($Vertical) {
            if ($icon -eq 'arrow') {
                $note = if ($sub) { "<span class=""flow-sub"">$sub</span>" } else { '' }
                "<div class=""rb-varrow""><svg viewBox=""0 0 12 30""><path d=""M6 1v26M1 21l5 6 5-6"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-label"">$title</span>$note</div>"
            } else {
                "<div class=""rb-vnode""><div class=""flow-icon"">$(Get-ReadmeIcon $icon)</div><div><div class=""flow-title"">$title</div><div class=""flow-text"">$sub</div></div></div>"
            }
        } elseif ($icon -eq 'arrow') {
            $class = if ($title -or $sub) { 'flow-arrow' } else { 'flow-arrow rb-bare' }
            "<div class=""$class""><span class=""flow-label"">$title</span><svg viewBox=""0 0 40 12""><path d=""M0 6h36M31 1l6 5-6 5"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-sub"">$sub</span></div>"
        } else {
            "<div class=""flow-node""><div class=""flow-icon"">$(Get-ReadmeIcon $icon)</div><div class=""flow-title"">$title</div><div class=""flow-text"">$sub</div></div>"
        }
    }
    $class = if ($Vertical) { 'flow rb-vflow' } else { 'flow rb-flow' }
    "<div class=""$class"">$($items -join '')</div>"
}

function ConvertTo-ReadmeCards([string[]]$Lines, [string]$Class = '') {
    $items = foreach ($l in $Lines) {
        $icon, $title, $text = $l.Split('|', 3).ForEach({ $_.Trim() })
        "<div class=""card-item""><div class=""card-icon"">$(Get-ReadmeIcon $icon)</div><div><div class=""card-title"">$(ConvertTo-ReadmeInline $title)</div><div class=""card-text"">$(ConvertTo-ReadmeInline $text)</div></div></div>"
    }
    "<div class=""cards $Class"">$($items -join '')</div>"
}

function Get-ReadmePill([string]$Text, [string]$Tone) { "<span class=""rb-pill"" style=""--tone: var(--cp-$Tone)"">$Text</span>" }

$Script:ReadmeCss = @'
html, body { background: #ffffff; }
html[data-theme="dark"], html[data-theme="dark"] body { background: #0d1117; }
:root { --cp-info: #0078d4; --cp-violet: #7c3aed; --cp-teal: #0d9488; }
html[data-theme="dark"] { --cp-info: #4da6ff; --cp-violet: #a78bfa; --cp-teal: #2dd4bf; }
body { display: block; margin: 0; padding: 0; }
.canvas { padding: 6px; }
.rb-pill { display: inline-block; padding: 1px 10px; margin: 8px 6px 0 0; border-radius: 999px; font-size: 11.5px; font-weight: 600; line-height: 1.6;
  color: var(--tone); background: color-mix(in srgb, var(--tone) 11%, transparent); border: 1px solid color-mix(in srgb, var(--tone) 38%, transparent); }
.rb-caption { font-size: 11.5px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: var(--cp-accent); margin: 0 0 8px 4px; }
.rb-caption span { color: var(--cp-text-muted); font-weight: 600; letter-spacing: 0.04em; text-transform: none; font-size: 12.5px; }
/* Banner */
.rb-hero { margin: 0; padding: 32px 36px 30px; }
.rb-hero-grid { position: relative; display: grid; grid-template-columns: minmax(0, 1fr) 230px; gap: 34px; align-items: center; }
.rb-hero h1 { font-size: 35px; }
.rb-hero .lead { margin: 18px 0 0; font-size: 17px; max-width: none; }
.rb-hero .badges { margin: 20px 0 0; }
.rb-stats { position: relative; display: grid; gap: 10px; }
.rb-stat { display: flex; align-items: center; gap: 14px; padding: 12px 16px; border-radius: 14px; background: var(--cp-panel-strong); border: 1px solid var(--cp-border); box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08); }
.rb-stat b { font-size: 30px; line-height: 1; color: var(--cp-accent); font-weight: 750; min-width: 40px; text-align: center; }
.rb-stat span { font-size: 13px; color: var(--cp-text-muted); line-height: 1.35; }
.rb-stat strong { display: block; color: var(--cp-text); font-size: 14px; }
/* Cards and flows */
.cards { margin: 0; }
.rb-cards2 { grid-template-columns: 1fr 1fr; }
.rb-flow { margin: 0; flex-wrap: nowrap; padding: 18px; gap: 4px; }
.rb-flow .flow-node { flex: 1 1 0; min-width: 0; padding: 14px 10px; }
.rb-flow .flow-title { font-size: 13.5px; overflow-wrap: anywhere; }
.rb-flow .flow-arrow { min-width: 0; width: 88px; flex: 0 0 88px; }
.rb-flow .flow-arrow.rb-bare { width: 46px; flex-basis: 46px; }
.rb-flow .flow-sub { max-width: 88px; }
.rb-space { height: 18px; }
/* How it works: vertical pipeline and the three modes */
.rb-hiw { display: grid; grid-template-columns: minmax(0, 1.08fr) minmax(0, 1fr); gap: 16px; align-items: stretch; }
.rb-col { display: flex; flex-direction: column; }
.rb-vflow { flex: 1; flex-direction: column; flex-wrap: nowrap; align-items: stretch; justify-content: center; gap: 0; margin: 0; padding: 16px 18px; }
.rb-vnode { display: flex; align-items: center; gap: 14px; padding: 11px 16px; border-radius: 12px; background: var(--cp-surface); border: 1px solid var(--cp-border); }
.rb-vnode .flow-icon { margin: 0; flex-shrink: 0; }
.rb-vnode .flow-text { margin-top: 1px; }
.rb-varrow { display: flex; align-items: center; gap: 10px; min-height: 36px; padding-left: 31px; }
.rb-varrow svg { width: 12px; height: 28px; color: var(--cp-accent); flex-shrink: 0; }
.rb-varrow .flow-sub { max-width: none; font-size: 12px; }
.rb-modes { flex: 1; display: flex; flex-direction: column; gap: 10px; }
.rb-modes .card-item { flex: 1; align-items: center; }
.rb-modes .card-title { display: flex; align-items: center; gap: 8px; }
.rb-chip { font-size: 11px; font-weight: 600; padding: 0 8px; border-radius: 999px; border: 1px solid var(--cp-border); color: var(--cp-text-muted); }
.rb-chip.hot { color: var(--cp-accent-fg); background: var(--cp-accent); border-color: var(--cp-accent); }
/* Phases */
.rb-phases { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; }
.rb-phase { display: flex; flex-direction: column; border: 1px solid var(--cp-border); border-top: 3px solid var(--cp-accent); border-radius: 14px; background: var(--cp-surface); padding: 14px 14px 12px; }
.rb-phase-head { display: flex; gap: 12px; align-items: center; min-height: 76px; padding-bottom: 12px; margin-bottom: 8px; border-bottom: 1px solid var(--cp-border); }
.rb-phase-eyebrow { font-size: 11px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: var(--cp-accent); }
.rb-phase-title { font-weight: 700; font-size: 14.5px; line-height: 1.25; }
.rb-steplist { list-style: none; margin: 0; padding: 0; }
.rb-steplist li { display: flex; align-items: center; gap: 9px; margin: 0; padding: 3px 0; }
.rb-num { flex-shrink: 0; width: 28px; text-align: center; font: 700 11.5px/20px Consolas, monospace; border-radius: 6px; background: var(--cp-accent-soft); color: var(--cp-accent); }
.rb-name { font: 12.5px/1.4 Consolas, "Courier New", monospace; color: var(--cp-text); }
.rb-phase-foot { margin-top: auto; padding-top: 10px; font-size: 12px; color: var(--cp-text-muted); }
.rb-phase-foot .rb-pill { margin: 0 6px 0 0; }
'@

function New-ReadmeGraphic {
    # One graphic, light and dark: HTML page -> height measured by Edge -> 2x screenshot.
    param([string]$Name, [string]$Body, [int]$Width)
    $pages = @{}
    foreach ($theme in 'light', 'dark') {
        $html = "<!doctype html><html lang=""en"" data-theme=""$theme""><head><meta charset=""utf-8""><style>$($assets.Css)`n$($Script:ReadmeCss)</style></head>" +
            "<body><div class=""canvas"" style=""width:$($Width)px"">$Body</div><script>document.body.setAttribute('data-h', Math.ceil(document.querySelector('.canvas').getBoundingClientRect().height));</script></body></html>"
        $pages[$theme] = Join-Path $work "readme-$Name-$theme.html"
        [IO.File]::WriteAllText($pages[$theme], $html, (New-Object Text.UTF8Encoding($false)))
    }
    $height = Get-PageHeight $pages['light'] $Width
    foreach ($theme in 'light', 'dark') { Save-Screenshot $pages[$theme] (Join-Path $OutputFolder "readme-$Name-$theme.png") $Width $height 2 }
}

function Invoke-ReadmeGraphics {
    param([string]$Root)
    $Script:assets = Get-ReadmeAssets -Root $Root
    if ($assets.Flows.Count -lt 3 -or $assets.Cards.Count -lt 2) { throw 'The guide must hold at least 3 flow blocks (how it works, runbook, batch lifecycle) and 2 cards blocks (overview, design principles).' }
    $mid = '&middot;'

    # Banner: the hero of the guide, with the key figures of the catalogue.
    $catalog = & {
        Import-Module (Join-Path (Join-Path $Root 'package') 'Modules\Exchange2019.Common.psd1') -Force -DisableNameChecking
        Get-DeploymentStepCatalog
    }
    $steps = @($catalog.Keys)
    $badges = @(
        "<span class=""badge badge-accent"">Version $($assets.Version)</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'terminal' 'icon-sm')Windows PowerShell 5.1</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'shield' 'icon-sm')Exchange 2019 only, legacy protected</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'tag' 'icon-sm')MIT license</span>"
    ) -join ''
    $banner = "<header class=""hero rb-hero""><div class=""rb-hero-grid""><div>" +
        "<div class=""hero-top""><div class=""hero-logo"">$(Get-ReadmeIcon 'refresh')</div><div><div class=""eyebrow"">Exchange Server $mid PowerShell framework</div><h1>Exchange 2013/2016 &rarr; 2019 Migration</h1></div></div>" +
        "<p class=""lead"">Deploys <strong>Exchange Server 2019</strong> next to the legacy servers, moves every mailbox in <strong>controlled batches</strong> and <strong>proves the cut-over</strong> &mdash; one entry point, readable reports.</p>" +
        "<div class=""badges"">$badges</div></div>" +
        "<div class=""rb-stats"">" +
        "<div class=""rb-stat""><b>$($steps.Count)</b><span><strong>resumable steps</strong>one file each</span></div>" +
        "<div class=""rb-stat""><b>4</b><span><strong>phases</strong>platform to cut-over</span></div>" +
        "<div class=""rb-stat""><b>3</b><span><strong>modes</strong>Inventory, Simulate, Apply</span></div>" +
        "</div></div></header>"
    New-ReadmeGraphic -Name 'banner' -Body $banner -Width 1080

    # Design principles: second cards block of the guide (chapter 1).
    New-ReadmeGraphic -Name 'principles' -Body (ConvertTo-ReadmeCards $assets.Cards[1].Lines 'rb-cards2') -Width 1080

    # How it works: first flow of the guide (chapter 2), as a vertical pipeline, and the three modes.
    $modes = @(
        [pscustomobject]@{ Icon = 'search'; Name = 'Inventory'; Chip = '<span class="rb-chip">read only</span>'; Text = 'Reads the current state of every target: the audit before any action.'; Pills = (Get-ReadmePill 'Inventoried' 'info') }
        [pscustomobject]@{ Icon = 'beaker'; Name = 'Simulate'; Chip = '<span class="rb-chip">no change</span>'; Text = 'Runs the pre-check, then the action with <code>-WhatIf</code>: the planned sequence, validated.'; Pills = (Get-ReadmePill 'Simulated' 'violet') + (Get-ReadmePill 'AlreadyDone' 'teal') }
        [pscustomobject]@{ Icon = 'play'; Name = 'Apply'; Chip = '<span class="rb-chip hot">changes</span>'; Text = 'Pre-check, action, then the target is read again. Typed confirmation unless <code>-Force</code>.'; Pills = (Get-ReadmePill 'Success' 'success') + (Get-ReadmePill 'AlreadyDone' 'teal') + (Get-ReadmePill 'Failed' 'danger') }
    )
    $modeHtml = ($modes | ForEach-Object { "<div class=""card-item""><div class=""card-icon"">$(Get-ReadmeIcon $_.Icon)</div><div><div class=""card-title"">$($_.Name) $($_.Chip)</div><div class=""card-text"">$($_.Text)</div><div>$($_.Pills)</div></div></div>" }) -join ''
    $howItWorks = "<div class=""rb-hiw""><div class=""rb-col""><div class=""rb-caption"">One entry point <span>$mid from the configuration to the reports</span></div>$(ConvertTo-ReadmeFlow $assets.Flows[0].Lines -Vertical)</div>" +
        "<div class=""rb-col""><div class=""rb-caption"">Three modes <span>$mid the same actions, from read-only to real</span></div><div class=""rb-modes"">$modeHtml</div></div></div>"
    New-ReadmeGraphic -Name 'how-it-works' -Body $howItWorks -Width 1080

    # The 26 steps by phase, from the catalogue of the module (labels of "-Step List").
    $phaseInfo = [ordered]@{
        Platform         = @{ N = 1; Title = 'Prepare the platform'; Icon = 'settings'; Foot = 'Run by <code>-Step All</code>' }
        HighAvailability = @{ N = 2; Title = 'Build high availability'; Icon = 'database'; Foot = 'Run by <code>-Step All</code>' }
        Runtime          = @{ N = 3; Title = 'Configure runtime services'; Icon = 'gear'; Foot = 'Run by <code>-Step All</code>' }
        Migration        = @{ N = 4; Title = 'Migrate, validate and clean up'; Icon = 'refresh'; Foot = (Get-ReadmePill 'Manual only' 'accent') + 'called one by one' }
    }
    $columns = foreach ($phase in $phaseInfo.Keys) {
        $info = $phaseInfo[$phase]
        $numbers = @($steps | Where-Object { $catalog[$_].Phase -eq $phase })
        $items = ($numbers | ForEach-Object { "<li><span class=""rb-num"">$('{0:D2}' -f [int]$_)</span><span class=""rb-name"">$($catalog[$_].Name)</span></li>" }) -join ''
        $range = '{0:D2}&ndash;{1:D2}' -f [int]$numbers[0], [int]$numbers[-1]
        "<div class=""rb-phase""><div class=""rb-phase-head""><div class=""card-icon"">$(Get-ReadmeIcon $info.Icon)</div><div><div class=""rb-phase-eyebrow"">Phase $($info.N) $mid Steps $range</div><div class=""rb-phase-title"">$($info.Title)</div></div></div>" +
            "<ul class=""rb-steplist"">$items</ul><div class=""rb-phase-foot"">$($info.Foot)</div></div>"
    }
    New-ReadmeGraphic -Name 'steps' -Body "<div class=""rb-phases"">$($columns -join '')</div>" -Width 1080

    # Migration runbook: second and third flows of the guide (chapter 9).
    $runbook = "<div class=""rb-caption"">Migration steps <span>$mid one command at a time, the reports read in between</span></div>" + (ConvertTo-ReadmeFlow $assets.Flows[1].Lines) +
        "<div class=""rb-space""></div><div class=""rb-caption"">Batch lifecycle <span>$mid followed live in MigrationStatus.html</span></div>" + (ConvertTo-ReadmeFlow $assets.Flows[2].Lines)
    New-ReadmeGraphic -Name 'runbook' -Body $runbook -Width 1080
}
#endregion

if ($Part -eq 'Samples') { Invoke-SampleReports -Root $root -Out $Work; return }
if ($Part -eq 'ConsoleRun') { Invoke-ConsoleRun -Root $root -Out $Work; return }

#region Main ---------------------------------------------------------------------------------------
if (-not $OutputFolder) { $OutputFolder = Join-Path $packageRoot 'Docs\images' }
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found: it takes the screenshots (headless mode).' }
$work = Join-Path ([IO.Path]::GetTempPath()) ('exm-doc-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work, $OutputFolder -Force | Out-Null
$ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

# Console capture -> terminal window (Windows Terminal "Campbell" colours).
$palette = @{ Black = '#0C0C0C'; DarkBlue = '#0037DA'; DarkGreen = '#13A10E'; DarkCyan = '#3A96DD'; DarkRed = '#C50F1F'; DarkMagenta = '#881798'; DarkYellow = '#C19C00'; Gray = '#CCCCCC'
    DarkGray = '#767676'; Blue = '#3B78FF'; Green = '#16C60C'; Cyan = '#61D6D6'; Red = '#E74856'; Magenta = '#B4009E'; Yellow = '#F9F1A5'; White = '#F2F2F2' }
function ConvertTo-TerminalHtml([string]$Capture, [string]$Title) {
    $lines = New-Object System.Collections.Generic.List[string]; $current = New-Object System.Text.StringBuilder
    foreach ($json in [IO.File]::ReadAllLines($Capture)) {
        if (-not $json.Trim()) { continue }
        $seg = $json | ConvertFrom-Json
        $text = [System.Net.WebUtility]::HtmlEncode([string]$seg.t)
        # Emoji are two columns wide in the terminal: give them exactly two character cells.
        $text = [regex]::Replace($text, '(?:[\uD800-\uDBFF][\uDC00-\uDFFF]|[\u2600-\u27BF\u23E9-\u23FA])\uFE0F?', { param($m) "<span class='emo'>$($m.Value)</span>" })
        $style = @()
        if ($seg.c -and $palette[$seg.c]) { $style += "color:$($palette[$seg.c])" }
        if ($seg.b -and $palette[$seg.b]) { $style += "background:$($palette[$seg.b])" }
        if ($seg.c -eq 'White') { $style += 'font-weight:600' }
        [void]$current.Append($(if ($style) { "<span style='$($style -join ';')'>$text</span>" } else { $text }))
        if ($seg.n) { $lines.Add($current.ToString()); [void]$current.Clear() }
    }
    if ($current.Length) { $lines.Add($current.ToString()) }
    $body = ($lines | ForEach-Object { "<div class='l'>$_&#8203;</div>" }) -join ''
    $height = 92 + 20 * $lines.Count
    $html = @"
<!doctype html><html><head><meta charset="utf-8"><style>
body{margin:0;background:#f7f4ef;padding:24px;font-family:'Segoe UI',sans-serif}
.win{width:1130px;border-radius:10px;overflow:hidden;box-shadow:0 18px 48px rgba(0,0,0,.25);background:#0C0C0C}
.bar{height:34px;background:#2b2b2b;display:flex;align-items:center;gap:8px;padding:0 14px;color:#cfcfcf;font-size:12px}
.dot{width:12px;height:12px;border-radius:50%;display:inline-block}
.term{padding:14px 18px 18px;color:#CCCCCC;font:14px/20px 'Cascadia Mono','Cascadia Code',Consolas,monospace;white-space:pre}
.l{height:20px}.emo{display:inline-block;width:2ch;font-family:'Segoe UI Emoji';font-size:13px;line-height:20px;text-align:left;overflow:visible}
</style></head><body><div class="win"><div class="bar"><span class="dot" style="background:#ff5f57"></span><span class="dot" style="background:#febc2e"></span><span class="dot" style="background:#28c840"></span><span style="margin-left:8px">$Title</span></div>
<div class="term">$body</div></div></body></html>
"@
    return [pscustomobject]@{ Html = $html; Height = $height + 48 }
}

function Save-Screenshot([string]$Html, [string]$Png, [int]$Width, [int]$Height, [int]$Scale = 1) {
    $url = 'file:///' + ($Html -replace '\\', '/') + '?scoutTheme=light'
    $profile = Join-Path $work 'edge-profile'
    if (Test-Path $Png) { Remove-Item $Png -Force }
    # Start-Process, not &: an Edge helper process can keep the output pipe open after the capture.
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$profile`"", "--window-size=$Width,$Height", "--force-device-scale-factor=$Scale", "--screenshot=`"$Png`"", "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddSeconds(45)
    while (-not (Test-Path $Png) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
    if (-not $proc.WaitForExit(10000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if (-not (Test-Path $Png)) { throw "Screenshot not written: $Png" }
}

function Get-PageHeight([string]$Html, [int]$Width) {
    # Height of the .canvas element: the page writes it in body[data-h], read with --dump-dom.
    $url = 'file:///' + ($Html -replace '\\', '/')
    $dom = Join-Path $work ('dom-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.html')
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,2000", '--dump-dom', "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden -RedirectStandardOutput $dom
    if (-not $proc.WaitForExit(45000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    # Edge helper processes inherit the output handle: read in shared mode, retry until written.
    $m = $null
    for ($i = 0; $i -lt 20 -and -not ($m -and $m.Success); $i++) {
        $stream = [IO.File]::Open($dom, 'Open', 'Read', 'ReadWrite')
        try { $text = (New-Object IO.StreamReader($stream)).ReadToEnd() } finally { $stream.Dispose() }
        $m = [regex]::Match($text, 'data-h="(\d+)"')
        if (-not $m.Success) { Start-Sleep -Milliseconds 250 }
    }
    if (-not $m.Success) { throw "Height not measured: $Html" }
    return [int]$m.Groups[1].Value
}

if ($Images -ne 'Readme') {
    Write-Host 'Building the sample reports (Windows PowerShell 5.1)...'
    & $ps51 -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -Part Samples -Work $work | Out-Null
    if ($LASTEXITCODE) { throw 'Sample reports failed.' }

    Write-Host 'Recording the console...'
    $captures = @{ 'console-list' = (Join-Path $work 'console-list.jsonl'); 'console-run' = (Join-Path $work 'console-run.jsonl') }
    $env:EXM_ICONS = 'Emoji'
    try {
        $env:EXM_CAPTURE_FILE = $captures['console-list']
        [IO.File]::AppendAllText($env:EXM_CAPTURE_FILE, ('{"t":"PS D:\\Tools\\Deploy-Exchange2019> ","c":"Cyan","b":null,"n":false}' + [Environment]::NewLine + '{"t":".\\Deploy-Exchange2019.ps1 -Step List","c":null,"b":null,"n":true}' + [Environment]::NewLine))
        & $ps51 -NoProfile -ExecutionPolicy Bypass -Command "& '$(Join-Path $root 'Deploy-Exchange2019.ps1')' -Step List" | Out-Null
        $env:EXM_CAPTURE_FILE = $captures['console-run']
        & $ps51 -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -Part ConsoleRun -Work (Join-Path $work 'run') | Out-Null
    } finally {
        Remove-Item Env:EXM_CAPTURE_FILE, Env:EXM_ICONS -ErrorAction SilentlyContinue
    }

    Write-Host 'Taking the screenshots (Microsoft Edge, headless)...'
    foreach ($name in 'console-list', 'console-run') {
        $page = ConvertTo-TerminalHtml $captures[$name] 'Windows PowerShell - D:\Tools\Deploy-Exchange2019'
        $file = Join-Path $work "$name.html"; [IO.File]::WriteAllText($file, $page.Html, (New-Object Text.UTF8Encoding($false)))
        Save-Screenshot $file (Join-Path $OutputFolder "$name.png") 1180 $page.Height
    }
    Save-Screenshot (Join-Path $work 'Report_GLOBAL_Apply.html') (Join-Path $OutputFolder 'report-run.png') 1440 1180
    Save-Screenshot (Join-Path $work 'MigrationStatus.html') (Join-Path $OutputFolder 'report-migration-status.png') 1440 1240
    Save-Screenshot (Join-Path $work 'ProtocolLogsAnalysis.html') (Join-Path $OutputFolder 'report-protocol-logs.png') 1440 1150
}
if ($Images -ne 'Guide') {
    Write-Host 'Rendering the README graphics (light and dark, 2x)...'
    Invoke-ReadmeGraphics -Root $root
}

Get-ChildItem $OutputFolder -Filter *.png | Select-Object Name, @{ n = 'KB'; e = { [math]::Round($_.Length / 1KB) } } | Format-Table -AutoSize | Out-String | Write-Host
# Edge helper processes of the temporary profile, if any are left.
Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*$work*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
if ($KeepWork) { Write-Host "Work folder: $work" } else { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
#endregion
