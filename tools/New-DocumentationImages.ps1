#Requires -Version 5.1
<#
.SYNOPSIS
    Regenerates the screenshots of the guide (Docs\images\*.png) from synthetic data.

.DESCRIPTION
    No Exchange server is needed. The tool:

      1. builds the HTML reports with anonymized sample data (contoso, EXCH2019xx), in Windows
         PowerShell 5.1 - the shell of the framework - with the real report functions of the
         module and of the steps;
      2. records the console of "-Step List" and of a short demonstration run (the real console
         functions of the module, emoji style as in Windows Terminal) through EXM_CAPTURE_FILE,
         and renders them as a terminal window;
      3. takes the screenshots with Microsoft Edge in headless mode.

    Images written: console-list.png, console-run.png, report-run.png,
    report-migration-status.png, report-protocol-logs.png.

.PARAMETER OutputFolder
    Default: Docs\images next to the tools folder.

.PARAMETER KeepWork
    Keeps the work folder (sample HTML files, console captures) and shows its path.

.EXAMPLE
    .\tools\New-DocumentationImages.ps1
    Then .\tools\Build-Documentation.ps1 to embed the new images in the HTML guide.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    Part of : Exchange 2013/2016 to 2019 Migration (repository tool, not in the package)
#>
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [switch]$KeepWork,
    # Internal: the parts that run in a separate Windows PowerShell 5.1 process.
    [ValidateSet('', 'Samples', 'ConsoleRun')][string]$Part = '',
    [string]$Work
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

#region Part 1 - sample reports (Windows PowerShell 5.1) --------------------------------------------
function Invoke-SampleReports {
    param([string]$Root, [string]$Out)
    New-Item -ItemType Directory -Path $Out -Force | Out-Null
    Import-Module (Join-Path $Root 'Modules\Exchange2019.Common.psd1') -Force -DisableNameChecking
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
    . (Join-Path $Root 'Steps\Step26-AnalyzeProtocolLogs.ps1')
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

    Import-NestedFunction (Join-Path $Root 'Steps\Step18-SystemMailboxMigration.ps1') 'New-SystemMailboxPlanHtml'
    $sys = @()
    $i = 0
    foreach ($t in 'ArbitrationMailbox', 'ArbitrationMailbox', 'ArbitrationMailbox', 'ArbitrationMailbox', 'ArbitrationMailbox', 'AuditLogMailbox', 'AuxAuditLogMailbox', 'DiscoveryMailbox') {
        $i++; $srv = if ($i -le 3) { 'EXCH201901' } else { 'EXCH201601' }
        $sys += [pscustomobject]@{ Name = "SystemMailbox{$([guid]::NewGuid())}"; DisplayName = @('Microsoft Exchange', 'Microsoft Exchange Approval Assistant', 'Microsoft Exchange Federation Mailbox', 'Migration', 'E4E Encryption Store', 'Microsoft Exchange Audit', 'Microsoft Exchange Aux Audit', 'Discovery Search Mailbox')[$i - 1]
            RecipientTypeDetails = $t; ServerName = $srv; Database = $(if ($srv -eq 'EXCH201901') { 'DB01' } else { 'DB-LEG-01' }); Identity = "$t-$i" }
    }
    New-SystemMailboxPlanHtml -Mailboxes $sys -TargetDbCount 4 -OutputPath (Join-Path $Out 'SystemMailboxPlan.html') -ReportMode 'Simulate'

    Import-NestedFunction (Join-Path $Root 'Steps\Step19-PrepareMigration.ps1') 'New-MigrationPlanHtml'
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

    Import-NestedFunction (Join-Path $Root 'Steps\Step20-RunMigration.ps1') 'New-MigrationStatusHtml'
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

    . (Join-Path $Root 'Steps\Step24-HealthCheckerReport.ps1')
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
    Import-Module (Join-Path $Root 'Modules\Exchange2019.Common.psd1') -Force -DisableNameChecking
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

if ($Part -eq 'Samples') { Invoke-SampleReports -Root $root -Out $Work; return }
if ($Part -eq 'ConsoleRun') { Invoke-ConsoleRun -Root $root -Out $Work; return }

#region Main ---------------------------------------------------------------------------------------
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'Docs\images' }
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found: it takes the screenshots (headless mode).' }
$work = Join-Path ([IO.Path]::GetTempPath()) ('exm-doc-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work, $OutputFolder -Force | Out-Null
$ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

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

function Save-Screenshot([string]$Html, [string]$Png, [int]$Width, [int]$Height) {
    $url = 'file:///' + ($Html -replace '\\', '/') + '?scoutTheme=light'
    $profile = Join-Path $work 'edge-profile'
    if (Test-Path $Png) { Remove-Item $Png -Force }
    # Start-Process, not &: an Edge helper process can keep the output pipe open after the capture.
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$profile`"", "--window-size=$Width,$Height", "--screenshot=`"$Png`"", "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddSeconds(45)
    while (-not (Test-Path $Png) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
    if (-not $proc.WaitForExit(10000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if (-not (Test-Path $Png)) { throw "Screenshot not written: $Png" }
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

Get-ChildItem $OutputFolder -Filter *.png | Select-Object Name, @{ n = 'KB'; e = { [math]::Round($_.Length / 1KB) } } | Format-Table -AutoSize | Out-String | Write-Host
# Edge helper processes of the temporary profile, if any are left.
Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*$work*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
if ($KeepWork) { Write-Host "Work folder: $work" } else { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
#endregion
