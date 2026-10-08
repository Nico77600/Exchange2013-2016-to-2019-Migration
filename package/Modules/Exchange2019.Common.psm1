<#
.SYNOPSIS
    Exchange 2013/2016 to 2019 Migration - shared module.

.DESCRIPTION
    Functions shared by the orchestrator (Deploy-Exchange2019.ps1) and the 26 steps. The module
    is organised in regions, in the order of an execution:

        1. Console and run log     Write-Ex* functions, Write-Log, Write-StepBanner (what the operator sees)
        2. Step catalogue          Get-DeploymentStepCatalog (the 26 steps, phases and icons)
        3. Session and reports     Initialize-DeploymentSession, Add-Report, Save-Report (CSV per step)
        4. Execution modes         Invoke-Action (Inventory / Simulate / Apply, idempotence)
        5. Exchange helpers        Initialize-ExchangeShell, Get-Exchange2019Servers, ...
        6. Configuration and state Import-ConfigCsv, Get-DeploymentState, Set-StepState
        7. HTML reports            shared theme (Get-ExHtml*) and the run report (New-HtmlReport)

    Designed for Windows PowerShell 5.1 (Exchange Management Shell 2019). The source file is
    plain ASCII: console symbols are built from their code points at run time, so the file is
    read correctly whatever the code page of the server.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    History : see CHANGELOG.md
#>

#region 0. Module state -------------------------------------------------------------------------

$Script:ToolVersion           = '2.0.0'
$Script:ToolAuthor            = 'Nicolas Fabert'
$Script:ReportEntries         = New-Object System.Collections.Generic.List[object]
$Script:GlobalReportEntries   = New-Object System.Collections.Generic.List[object]
$Script:CurrentStep           = $null
$Script:Mode                  = 'Inventory'
$Script:OutputFolder          = $null
$Script:OutputBaseFolder      = $null
$Script:TimestampRun          = (Get-Date -Format 'yyyyMMdd_HHmmss')
$Script:RunStart              = Get-Date
$Script:DefaultExchangeServer = $null
$Script:DefaultCredential     = $null
$Script:ExchangeRemoteSession = $null
$Script:ReportSaved           = $false
$Script:LogWriter             = $null
$Script:LogPath               = $null
$Script:StepIndex             = 0
$Script:StepTotal             = 0
$Script:StepIcon              = 'Step'
$Script:Capture               = $null
# Documentation screenshots only (tools\New-DocumentationImages.ps1): every console segment is also
# appended, as one JSON line, to the file named by EXM_CAPTURE_FILE.
$Script:CaptureFile           = $env:EXM_CAPTURE_FILE

#endregion
#region 1. Console and run log -------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# Console theme: the same layout as the sister project Purview DLP Report (title card, numbered
# step pills, one line per result, framed summary card).
#   - Colours use the console colours (Write-Host -ForegroundColor), never ANSI sequences: the
#     transcripts of Start-Transcript stay clean and the classic console of Windows Server shows
#     them without any setting. NO_COLOR (any value) disables the colours.
#   - Icons: emoji in Windows Terminal and VS Code, symbols of the classic console fonts elsewhere
#     (conhost has no font fallback: a character missing from Consolas or Lucida Console would be
#     shown as an empty box, so the symbol set uses only code page 437 and Latin-1 characters).
#     Force a style with the environment variable EXM_ICONS = Emoji | Symbols | Ascii.
#   - Frames: rounded corners, present in Consolas; square corners with Lucida Console and the
#     raster font, which have no rounded corners (the console font is read once at run time).
# -------------------------------------------------------------------------------------------------

$Script:UseColor = -not $env:NO_COLOR
$Script:IconStyle = if ($env:EXM_ICONS -in 'Emoji', 'Symbols', 'Ascii') { $env:EXM_ICONS }
    elseif ([Console]::IsOutputRedirected) { 'Symbols' }
    elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') { 'Emoji' }
    else { 'Symbols' }

# Console colours of the theme. Accent = the crimson of the reports.
$Script:Theme = @{
    Accent = 'Red'; AccentBack = 'DarkRed'; AccentText = 'White'
    Ok = 'Green'; Warn = 'Yellow'; Fail = 'Red'; Info = 'Cyan'; Sim = 'Magenta'; Done = 'DarkCyan'
    Dim = 'DarkGray'; Text = 'Gray'; Strong = 'White'
}

function Get-ExIconSet {
    <# Icons of one console style. Symbols: only characters of the classic console fonts. #>
    param([Parameter(Mandatory)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)
    $u = { param([int]$Code) [char]::ConvertFromUtf32($Code) }
    switch ($Style) {
        'Emoji' { return @{
                Logo = & $u 0x1F4E7; Ok = & $u 0x2705; Warn = (& $u 0x26A0) + [char]0xFE0F; Fail = & $u 0x274C
                Info = & $u 0x1F539; Skip = & $u 0x23E9; Sim = & $u 0x1F9EA; Dot = [string][char]0x00B7; Sub = [string][char]0x2192
                Section = & $u 0x1F538; Step = & $u 0x1F527; Mode = & $u 0x1F539; Plan = & $u 0x1F4CB; Server = & $u 0x1F3E2
                File = & $u 0x1F4C4; Folder = & $u 0x1F4C1; Log = & $u 0x1F4DD; Clock = & $u 0x23F3; Report = & $u 0x1F4CA
                Done = & $u 0x1F389; Target = & $u 0x1F3AF; Calendar = & $u 0x1F4C5
                Key = & $u 0x1F511; Certificate = & $u 0x1F4DC; Globe = & $u 0x1F310; Disk = & $u 0x1F4BD; Database = & $u 0x1F4BE
                Queue = & $u 0x1F4E6; Plug = & $u 0x1F50C; Link = & $u 0x1F517; Copy = & $u 0x1F4CB; Shield = & $u 0x1F512
                Mail = & $u 0x1F4E8; Compass = & $u 0x1F9ED; Rocket = & $u 0x1F680; Flag = & $u 0x1F3C1; Broom = & $u 0x1F9F9
                Chart = & $u 0x1F4CA; Health = & $u 0x1FA7A; Search = & $u 0x1F50E
            } }
        'Symbols' { return @{
                Logo = & $u 0x2666; Ok = & $u 0x221A; Warn = & $u 0x25B2; Fail = & $u 0x00D7
                Info = & $u 0x2022; Skip = & $u 0x00BB; Sim = & $u 0x2248; Dot = & $u 0x00B7; Sub = & $u 0x2192
                Section = & $u 0x25BA; Step = & $u 0x263C; Mode = & $u 0x2022; Plan = & $u 0x25BA; Server = & $u 0x2302
                File = & $u 0x25AC; Folder = & $u 0x2302; Log = & $u 0x00B6; Clock = & $u 0x25CB; Report = & $u 0x2261
                Done = & $u 0x221A; Target = & $u 0x25D9; Calendar = & $u 0x263C
                Key = & $u 0x2194; Certificate = & $u 0x00A7; Globe = & $u 0x25CB; Disk = & $u 0x25D8; Database = & $u 0x25A0
                Queue = & $u 0x25AC; Plug = & $u 0x2195; Link = & $u 0x221E; Copy = & $u 0x2261; Shield = & $u 0x00A4
                Mail = '@'; Compass = & $u 0x25BA; Rocket = & $u 0x25BA; Flag = & $u 0x221A; Broom = & $u 0x00D7
                Chart = & $u 0x2261; Health = & $u 0x2665; Search = & $u 0x25BA
            } }
        default { return @{
                Logo = '*'; Ok = '+'; Warn = '!'; Fail = 'x'; Info = '-'; Skip = '>'; Sim = '~'; Dot = '-'; Sub = '>'
                Section = '>'; Step = '*'; Mode = '-'; Plan = '>'; Server = '#'; File = '-'; Folder = '>'; Log = '='; Clock = '~'
                Report = '='; Done = '*'; Target = 'o'; Calendar = ':'; Key = '@'; Certificate = '$'; Globe = 'o'; Disk = '#'
                Database = '#'; Queue = '='; Plug = '|'; Link = '&'; Copy = '='; Shield = '+'; Mail = '@'; Compass = '>'
                Rocket = '>'; Flag = '+'; Broom = 'x'; Chart = '='; Health = '+'; Search = '?'
            } }
    }
}

function Get-ExFrameSet {
    <#
    .SYNOPSIS
        Frame characters. Rounded corners, except with the Ascii style and with the console fonts
        that have no rounded corners (Lucida Console, raster font 'Terminal'): square corners.
    #>
    param([Parameter(Mandatory)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style, [AllowNull()][string]$FontName)
    if ($Style -eq 'Ascii') { return @{ TopLeft = [char]'+'; TopRight = [char]'+'; BottomLeft = [char]'+'; BottomRight = [char]'+'; Horizontal = [char]'-'; Vertical = [char]'|' } }
    if ($Style -eq 'Symbols' -and $FontName -in 'Lucida Console', 'Terminal') {
        return @{ TopLeft = [char]0x250C; TopRight = [char]0x2510; BottomLeft = [char]0x2514; BottomRight = [char]0x2518; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
    }
    return @{ TopLeft = [char]0x256D; TopRight = [char]0x256E; BottomLeft = [char]0x2570; BottomRight = [char]0x256F; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
}

function Get-ExConsoleFontName {
    <# Face name of the classic console font (GetCurrentConsoleFontEx), or $null when it cannot be read. #>
    if ([Console]::IsOutputRedirected) { return $null }
    try {
        if (-not ('ExMigration.ConsoleFont' -as [type])) {
            Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace ExMigration {
    public static class ConsoleFont {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct FontInfoEx {
            public uint Size; public uint Font; public short Width; public short Height;
            public int Family; public int Weight;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string FaceName;
        }
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr GetStdHandle(int handle);
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool GetCurrentConsoleFontEx(IntPtr output, bool maximumWindow, ref FontInfoEx info);
        public static string FaceName() {
            try {
                FontInfoEx info = new FontInfoEx();
                info.Size = (uint)Marshal.SizeOf(info);
                return GetCurrentConsoleFontEx(GetStdHandle(-11), false, ref info) ? info.FaceName : null;
            } catch { return null; }
        }
    }
}
'@
        }
        return [ExMigration.ConsoleFont]::FaceName()
    } catch { return $null }
}

$Script:Icons   = Get-ExIconSet $Script:IconStyle
$Script:Frame   = Get-ExFrameSet $Script:IconStyle $(if ($Script:IconStyle -eq 'Symbols') { Get-ExConsoleFontName })
# Emoji are two columns wide in the console; symbols one: pad symbols so that text stays aligned.
$Script:IconPad = if ($Script:IconStyle -eq 'Emoji') { ' ' } else { '  ' }
$Script:IconWidth = if ($Script:IconStyle -eq 'Emoji') { 2 } else { 1 }
$Script:Arrow   = if ($Script:IconStyle -eq 'Ascii') { '->' } else { [string][char]0x2192 }
$Script:Mid     = if ($Script:IconStyle -eq 'Ascii') { '-' } else { [string][char]0x00B7 }

function Get-ExIcon {
    <# One icon of the current style, followed by its padding. Unknown names fall back to Info. #>
    param([Parameter(Mandatory)][string]$Name)
    $icon = $Script:Icons[$Name]
    if (-not $icon) { $icon = $Script:Icons['Info'] }
    return $icon + $Script:IconPad
}

function Get-ExToolInfo {
    <# Name, version and author shown by the title card, the reports and the documentation. #>
    [pscustomobject]@{
        Name    = 'Exchange 2013/2016 ' + $Script:Arrow + ' 2019 Migration'
        Version = $Script:ToolVersion
        Author  = $Script:ToolAuthor
    }
}

function Enable-ExConsoleCapture {
    <#
    .SYNOPSIS
        Records every console segment (text and colours) in memory, for the documentation
        screenshots (tools\New-DocumentationImages.ps1). -Quiet also stops the console output.
    #>
    param([switch]$Quiet)
    $Script:Capture = New-Object System.Collections.Generic.List[object]
    $Script:CaptureQuiet = [bool]$Quiet
}

function Get-ExConsoleCapture { return $Script:Capture }

function Write-ExHost {
    <# The only function that writes to the console: colours are applied here. #>
    param([AllowEmptyString()][string]$Text = '', [string]$Color, [string]$Background, [switch]$NoNewline)
    if ($Script:CaptureFile) {
        [IO.File]::AppendAllText($Script:CaptureFile, (([pscustomobject]@{ t = $Text; c = $Color; b = $Background; n = -not $NoNewline } | ConvertTo-Json -Compress) + [Environment]::NewLine))
    }
    if ($null -ne $Script:Capture) {
        $Script:Capture.Add([pscustomobject]@{ Text = $Text; Color = $Color; Background = $Background; NewLine = -not $NoNewline })
        if ($Script:CaptureQuiet) { return }
    }
    $params = @{ Object = $Text; NoNewline = [bool]$NoNewline }
    if ($Script:UseColor -and $Color) { $params['ForegroundColor'] = $Color }
    if ($Script:UseColor -and $Background) { $params['BackgroundColor'] = $Background }
    Write-Host @params
}

function Format-ExDuration {
    param([Parameter(Mandatory)][double]$Seconds)
    # 0.0 (not 0): with an integer first argument PowerShell picks Math.Max(int, int) and drops the decimals.
    $t = [TimeSpan]::FromTicks([long]([Math]::Max(0.0, $Seconds) * 10000000))
    $c = [Globalization.CultureInfo]::InvariantCulture
    if ($t.TotalDays -ge 2) { return [string]::Format($c, '{0} d {1:00} h', [int][Math]::Floor($t.TotalDays), $t.Hours) }
    if ($t.TotalHours -ge 1) { return [string]::Format($c, '{0} h {1:00} min', [int][Math]::Floor($t.TotalHours), $t.Minutes) }
    if ($t.TotalMinutes -ge 1) { return [string]::Format($c, '{0} min {1:00} s', $t.Minutes, $t.Seconds) }
    return [string]::Format($c, '{0:0.0} s', $t.TotalSeconds)
}

function Format-ExNumber {
    <# Number with thousands separators, always 1,234 (en-US), whatever the regional settings. #>
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '-' }
    return ([double]$Value).ToString('N0', [Globalization.CultureInfo]::GetCultureInfo('en-US'))
}

function Start-ExRun {
    <#
    .SYNOPSIS
        Creates the run folder (<OutputFolder>\<yyyyMMdd_HHmmss>) and opens the run log
        Deploy-Exchange2019.log in it. Returns the folder and the log path.
    .DESCRIPTION
        The run log has one line per message, with a timestamp and a level, and never contains
        colours or icons. It completes the transcripts written for each step.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$OutputFolder)
    $runFolder = Join-Path $OutputFolder $Script:TimestampRun
    if (-not (Test-Path -LiteralPath $runFolder)) { New-Item -ItemType Directory -Path $runFolder -Force | Out-Null }
    $Script:OutputBaseFolder = $OutputFolder
    $Script:OutputFolder = $runFolder
    if (-not $Script:LogWriter) {
        $Script:LogPath = Join-Path $runFolder 'Deploy-Exchange2019.log'
        $stream = New-Object IO.FileStream($Script:LogPath, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
        $Script:LogWriter = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
        $Script:LogWriter.AutoFlush = $true
    }
    [pscustomobject]@{ Folder = $runFolder; Log = $Script:LogPath }
}

function Stop-ExRun {
    if ($Script:LogWriter) { $Script:LogWriter.Dispose(); $Script:LogWriter = $null }
}

function Write-ExLog {
    <# Writes one line to the run log only (never to the console). #>
    param([ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP', 'DEBUG')][string]$Level = 'INFO', [AllowEmptyString()][string]$Message = '')
    if ($Script:LogWriter) {
        $Script:LogWriter.WriteLine(('{0} [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss.fffzzz'), $Level, $Message))
    }
}

function Write-ExBanner {
    <#
    .SYNOPSIS
        Title card at the start of a run:

          +--------------------------------------------------------------------------+
          |  *  Exchange 2013/2016 -> 2019 Migration          v2.0.0 . Nicolas Fabert |
          |     Deployment, mailbox migration and validation                         |
          +--------------------------------------------------------------------------+
             -  Mode      Inventory
    .PARAMETER Details
        Ordered list of rows: key = label, value = @(IconName, Text) or plain text.
    #>
    param([Parameter(Mandatory)][string]$Title, [string]$Subtitle, [System.Collections.Specialized.OrderedDictionary]$Details)
    $F = $Script:Frame; $T = $Script:Theme; $width = 74
    $right = 'v{0} {1} {2}' -f $Script:ToolVersion, $Script:Mid, $Script:ToolAuthor
    $left = '  {0}  {1}' -f $Script:Icons['Logo'], $Title
    $leftWidth = $left.Length - $Script:Icons['Logo'].Length + $Script:IconWidth
    $gap = [Math]::Max(1, $width - $leftWidth - $right.Length - 2)
    Write-ExHost ''
    Write-ExHost ('  ' + $F.TopLeft + [string]::new($F.Horizontal, $width) + $F.TopRight) -Color $T.Accent
    Write-ExHost '  ' -NoNewline; Write-ExHost ([string]$F.Vertical) -Color $T.Accent -NoNewline
    Write-ExHost $left -Color $T.Strong -NoNewline
    Write-ExHost ([string]::new(' ', $gap)) -NoNewline
    Write-ExHost ($right + '  ') -Color $T.Dim -NoNewline
    Write-ExHost ([string]$F.Vertical) -Color $T.Accent
    if ($Subtitle) {
        Write-ExHost '  ' -NoNewline; Write-ExHost ([string]$F.Vertical) -Color $T.Accent -NoNewline
        Write-ExHost ('     ' + $Subtitle).PadRight($width) -Color $T.Dim -NoNewline
        Write-ExHost ([string]$F.Vertical) -Color $T.Accent
    }
    Write-ExHost ('  ' + $F.BottomLeft + [string]::new($F.Horizontal, $width) + $F.BottomRight) -Color $T.Accent
    Write-ExLog 'STEP' ('=== {0} v{1} ===' -f $Title, $Script:ToolVersion)
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $value = $Details[$key]
            if ($value -is [array]) { $icon = Get-ExIcon $value[0]; $text = [string]$value[1] } else { $icon = '   '; $text = [string]$value }
            Write-ExHost '     ' -NoNewline
            Write-ExHost $icon -Color $T.Info -NoNewline
            Write-ExHost ('{0,-10}' -f $key) -Color $T.Dim -NoNewline
            Write-ExHost (' ' + $text)
            Write-ExLog 'INFO' ('{0}: {1}' -f $key, $text)
        }
    }
}

function Write-ExStep {
    <#
    .SYNOPSIS
        Step header with a coloured number pill and an icon, e.g.

          [ 2/5 ] (icon)  Step 03 . Virtual directories        Inventory
    #>
    param([Parameter(Mandatory)][string]$Pill, [Parameter(Mandatory)][string]$Title, [string]$Icon = 'Step', [string]$Note)
    $T = $Script:Theme
    Write-ExHost ''
    Write-ExHost '  ' -NoNewline
    Write-ExHost (' {0} ' -f $Pill) -Color $T.AccentText -Background $T.AccentBack -NoNewline
    Write-ExHost (' ' + (Get-ExIcon $Icon)) -NoNewline
    if ($Note) {
        Write-ExHost $Title -Color $T.Strong -NoNewline
        Write-ExHost ('   ' + $Note) -Color $T.Dim
    } else {
        Write-ExHost $Title -Color $T.Strong
    }
    Write-ExLog 'STEP' ('[{0}] {1}{2}' -f $Pill, $Title, $(if ($Note) { " ($Note)" } else { '' }))
}

function Write-ExItem {
    <# One indented result line with a status icon, also written to the run log. #>
    param(
        [ValidateSet('Ok', 'Warn', 'Fail', 'Info', 'Skip', 'Dim', 'Sim')][string]$Status = 'Info',
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        [string]$Icon,
        [int]$Indent = 6
    )
    $T = $Script:Theme
    $iconColor = @{ Ok = $T.Ok; Warn = $T.Warn; Fail = $T.Fail; Info = $T.Info; Skip = $T.Dim; Dim = $T.Dim; Sim = $T.Sim }[$Status]
    $textColor = @{ Ok = $null; Warn = $T.Warn; Fail = $T.Fail; Info = $null; Skip = $T.Dim; Dim = $T.Dim; Sim = $null }[$Status]
    $level = @{ Ok = 'OK'; Warn = 'WARN'; Fail = 'ERROR'; Info = 'INFO'; Skip = 'INFO'; Dim = 'DEBUG'; Sim = 'INFO' }[$Status]
    $iconName = if ($Icon) { $Icon } elseif ($Status -eq 'Dim') { 'Dot' } else { $Status }
    Write-ExHost ([string]::new(' ', $Indent)) -NoNewline
    Write-ExHost (Get-ExIcon $iconName) -Color $iconColor -NoNewline
    Write-ExHost $Text -Color $textColor
    Write-ExLog $level $Text
}

function Write-ExRule {
    <# Sub-heading inside a step (a group of results), e.g. "> SMTP protocol logs". #>
    param([Parameter(Mandatory)][string]$Text)
    Write-ExHost ''
    Write-ExHost '    ' -NoNewline
    Write-ExHost (Get-ExIcon 'Section') -Color $Script:Theme.Accent -NoNewline
    Write-ExHost $Text -Color $Script:Theme.Strong
    Write-ExLog 'STEP' ('-- {0}' -f $Text)
}

function Write-ExSummary {
    <#
    .SYNOPSIS
        Final summary card:

          +- (icon)  Run complete -------------------------------------------------+
            (icon)  Steps      5 completed . 0 failed
          +-----------------------------------------------------------------------+
    .PARAMETER Values
        Ordered list: key = label, value = @(IconName, Text) or plain text.
    #>
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][System.Collections.Specialized.OrderedDictionary]$Values, [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok')
    $F = $Script:Frame; $T = $Script:Theme; $width = 74
    $color = @{ Ok = $T.Ok; Warn = $T.Warn; Fail = $T.Fail }[$Status]
    $icon = $Script:Icons[@{ Ok = 'Done'; Warn = 'Warn'; Fail = 'Fail' }[$Status]]
    $head = ' {0}  {1} ' -f $icon, $Title
    $rest = [Math]::Max(2, $width - 1 - ($head.Length - $icon.Length + $Script:IconWidth))
    Write-ExHost ''
    Write-ExHost ('  ' + $F.TopLeft + $F.Horizontal) -Color $color -NoNewline
    Write-ExHost $head -Color $color -NoNewline
    Write-ExHost ([string]::new($F.Horizontal, $rest) + $F.TopRight) -Color $color
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        if ($value -is [array]) { $rowIcon = Get-ExIcon $value[0]; $text = [string]$value[1] } else { $rowIcon = '   '; $text = [string]$value }
        Write-ExHost '    ' -NoNewline
        Write-ExHost $rowIcon -Color $T.Info -NoNewline
        Write-ExHost ('{0,-10}' -f $key) -Color $T.Dim -NoNewline
        Write-ExHost (' ' + $text)
        Write-ExLog 'INFO' ('Summary - {0}: {1}' -f $key, $text)
    }
    Write-ExHost ('  ' + $F.BottomLeft + [string]::new($F.Horizontal, $width) + $F.BottomRight) -Color $color
    Write-ExHost ''
}

function Set-ExStepContext {
    <# Position of the step in the selection (pill "2/5") and its icon, set by the orchestrator. #>
    param([int]$Index, [int]$Total, [string]$Icon = 'Step')
    $Script:StepIndex = $Index; $Script:StepTotal = $Total; $Script:StepIcon = $Icon
}

function Write-Log {
    <#
    .SYNOPSIS
        Console message of a step, written with the theme and copied to the run log.
    .DESCRIPTION
        Levels used by the steps:
          Info     neutral information           Success  result OK (green)
          Warning  attention needed (yellow)      Error    failure (red)
          Debug    technical detail (dimmed)      Sub      secondary detail (dimmed)
          Step     sub-heading inside a step; decorations made of '=' and '-' are removed,
                   a message made only of decorations prints an empty line.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()]
        [string]$Message,

        [ValidateSet('Info','Success','Warning','Error','Debug','Step','Sub')]
        [string]$Level = 'Info'
    )
    switch ($Level) {
        'Step' {
            $text = ($Message -replace '^[\s=\-]+', '' -replace '[\s=\-]+$', '')
            if ($text) { Write-ExRule $text } else { Write-ExHost '' }
        }
        'Success' { Write-ExItem Ok $Message }
        'Warning' { Write-ExItem Warn $Message }
        'Error'   { Write-ExItem Fail $Message }
        'Debug'   { Write-ExItem Dim $Message }
        'Sub'     { Write-ExItem Dim $Message -Icon 'Sub' }
        default   { Write-ExItem Info $Message }
    }
}

function Write-StepBanner {
    <#
    .SYNOPSIS
        Header of a step: numbered pill, icon and title, the planned actions and the scope.
    .PARAMETER IncludeLegacyServers
        The step ALSO changes the Exchange 2013/2016 servers (Step 15 Kerberos ASA, which must be
        deployed on every Client Access server): a warning replaces the "Exchange 2019 only" line.
    .PARAMETER LegacyImpactNote
        Reason of the change on the legacy servers, shown under the warning.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$StepName,
        [Parameter(Mandatory)] [string]$Title,
        [Parameter(Mandatory)] [string[]]$Actions,
        [string]$Mode,
        [switch]$IncludeLegacyServers,
        [string]$LegacyImpactNote
    )
    $pill = if ($Script:StepTotal -gt 0) { '{0}/{1}' -f $Script:StepIndex, $Script:StepTotal } else { 'Step' }
    $number = if ($StepName -match '^\d+$') { 'Step {0:D2}' -f [int]$StepName } else { "Step $StepName" }
    Write-ExStep -Pill $pill -Title ('{0} {1} {2}' -f $number, $Script:Mid, $Title) -Icon $Script:StepIcon -Note $Mode
    Write-ExHost '      ' -NoNewline
    Write-ExHost 'Planned actions' -Color $Script:Theme.Dim
    foreach ($a in $Actions) {
        Write-ExHost '        ' -NoNewline
        Write-ExHost (Get-ExIcon 'Dot') -Color $Script:Theme.Dim -NoNewline
        Write-ExHost $a -Color $Script:Theme.Text
        Write-ExLog 'INFO' ('Planned: {0}' -f $a)
    }
    if ($IncludeLegacyServers) {
        Write-ExItem Warn 'This step ALSO changes the Exchange 2013/2016 servers.'
        if ($LegacyImpactNote) { Write-ExItem Dim $LegacyImpactNote -Icon 'Sub' -Indent 9 }
    } else {
        Write-ExItem Info 'Scope: Exchange 2019 servers only. Exchange 2013/2016 servers in coexistence are never changed.' -Icon 'Target'
    }
}

#endregion

#region 2. Step catalogue -----------------------------------------------------------------------

function Get-DeploymentStepCatalog {
    <#
    .SYNOPSIS
        The 26 steps in execution order: number -> Name, File, Title, Phase, Icon, ManualOnly.
    .DESCRIPTION
        Steps 18 to 26 are ManualOnly: they are skipped by "-Step All -Mode Apply" and run only
        when they are selected explicitly, by number or by name.
    #>
    $catalog = [ordered]@{}
    $rows = @(
        @('1',  'LicenseKey',             'Step01-LicenseKey.ps1',             'Exchange 2019 product key',                                                         'Platform', 'Key',         $false),
        @('2',  'Certificate',            'Step02-Certificate.ps1',            'Certificate exported from legacy Exchange and imported on Exchange 2019',            'Platform', 'Certificate', $false),
        @('3',  'VirtualDirectories',     'Step03-VirtualDirectories.ps1',     'Virtual directories (Autodiscover, OWA, OAB, ECP, EWS, ActiveSync, OA, MAPI)',       'Platform', 'Globe',       $false),
        @('4',  'Disks',                  'Step04-Disks.ps1',                  'Disk preparation and formatting',                                                   'Platform', 'Disk',        $false),
        @('5',  'TransportQueue',         'Step05-TransportQueue.ps1',         'Transport queue database moved',                                                    'Platform', 'Queue',       $false),
        @('6',  'LogPaths',               'Step06-LogPaths.ps1',               'Log paths (Frontend, Hub, Mailbox Delivery, IIS)',                                   'Platform', 'Log',         $false),
        @('7',  'ReceiveConnectors',      'Step07-ReceiveConnectors.ps1',      'Receive connectors copied from legacy Exchange (with AD permissions)',               'Platform', 'Plug',        $false),
        @('8',  'CreateDAG',              'Step08-CreateDAG.ps1',              'IP-less database availability group',                                               'HighAvailability', 'Link', $false),
        @('9',  'AddDAGMembers',          'Step09-AddDAGMembers.ps1',          'Exchange 2019 servers added to the DAG',                                            'HighAvailability', 'Link', $false),
        @('10', 'CreateDatabases',        'Step10-CreateDatabases.ps1',        'Mailbox databases',                                                                 'HighAvailability', 'Database', $false),
        @('11', 'CreateDatabaseCopies',   'Step11-CreateDatabaseCopies.ps1',   'Database copies (passive and lagged)',                                              'HighAvailability', 'Copy', $false),
        @('12', 'IISXForwardedFor',       'Step12-IISXForwardedFor.ps1',       'X-Forwarded-For field in the IIS logs',                                             'Runtime',  'Log',         $false),
        @('13', 'CheckAntimalware',       'Step13-CheckAntimalware.ps1',       'FIPFS anti-malware engine check',                                                   'Runtime',  'Shield',      $false),
        @('14', 'EventLogSize',           'Step14-EventLogSize.ps1',           'Application and System event logs set to 1 GB',                                     'Runtime',  'Log',         $false),
        @('15', 'KerberosASA',            'Step15-KerberosASA.ps1',            'Kerberos alternate service account and password replication',                       'Runtime',  'Key',         $false),
        @('16', 'MAPIOverHTTPS',          'Step16-MAPIOverHTTPS.ps1',          'MAPI over HTTP enabled only for the users on Exchange 2019',                         'Runtime',  'Globe',       $false),
        @('17', 'ManageIISLogs',          'Step17-ManageIISLogs.ps1',          'IIS log management: scheduled compression and purge',                               'Runtime',  'Broom',       $false),
        @('18', 'SystemMailboxMigration', 'Step18-SystemMailboxMigration.ps1', 'System mailboxes (Arbitration, AuditLog, AuxAuditLog) moved to Exchange 2019',       'Migration', 'Mail',       $true),
        @('19', 'PrepareMigration',       'Step19-PrepareMigration.ps1',       'Migration batches prepared (balanced move requests)',                               'Migration', 'Compass',    $true),
        @('20', 'RunMigration',           'Step20-RunMigration.ps1',           'Migration batches started and followed (suspended before completion)',              'Migration', 'Rocket',     $true),
        @('21', 'CompleteMigration',      'Step21-CompleteMigration.ps1',      'Migration batches completed (now or at a scheduled time)',                          'Migration', 'Flag',       $true),
        @('22', 'CleanupMigration',       'Step22-CleanupMigration.ps1',       'Leftover move requests, migration batches and migration users removed',             'Migration', 'Broom',      $true),
        @('23', 'ResetDatabaseQuotas',    'Step23-ResetDatabaseQuotas.ps1',    'Mailbox quotas set back after the migration (4 / 4.5 / 5 GB)',                       'Migration', 'Chart',      $true),
        @('24', 'HealthCheckerReport',    'Step24-HealthCheckerReport.ps1',    'HealthChecker on every Exchange 2019 server and warnings/errors report',             'Migration', 'Health',     $true),
        @('25', 'CleanupAutodiscoverSCP', 'Step25-CleanupAutodiscoverSCP.ps1', 'AutoDiscoverServiceInternalUri cleared on Exchange 2013/2016 (SCP)',                  'Migration', 'Broom',      $true),
        @('26', 'AnalyzeProtocolLogs',    'Step26-AnalyzeProtocolLogs.ps1',    'IIS and SMTP protocol logs: real-user traffic, legacy versus 2019',                  'Migration', 'Search',     $true)
    )
    foreach ($r in $rows) {
        $entry = @{ Name = $r[1]; File = $r[2]; Title = $r[3]; Phase = $r[4]; Icon = $r[5] }
        if ($r[6]) { $entry['ManualOnly'] = $true }
        $catalog[$r[0]] = $entry
    }
    return $catalog
}

#endregion
#region 3. Session and reports --------------------------------------------------------------------

function Initialize-DeploymentSession {
    <#
    .SYNOPSIS
        Prepares one step: run folder, mode, empty step report and the step transcript.
    .DESCRIPTION
        <OutputFolder> keeps DeploymentState.json (shared by all the runs, used by -Resume).
        Every run writes into its own sub-folder <OutputFolder>\<yyyyMMdd_HHmmss>, so that the
        reports of two runs never mix.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Inventory','Simulate','Apply')]
        [string]$Mode,

        [Parameter(Mandatory)]
        [string]$OutputFolder,

        [string]$StepName = 'Init'
    )

    $run = Start-ExRun -OutputFolder $OutputFolder

    $Script:Mode          = $Mode
    $Script:CurrentStep   = $StepName
    $Script:ReportEntries = New-Object System.Collections.Generic.List[object]

    $transcriptPath = Join-Path $run.Folder ("Transcript_{0}.log" -f $StepName)
    try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch { }
    Start-Transcript -Path $transcriptPath -Append | Out-Null

    Write-ExLog 'DEBUG' ("Session ready. Mode = $Mode. Step = $StepName. Output = $($run.Folder)")
}

function Stop-DeploymentSession {
    [CmdletBinding()] param()
    try { Stop-Transcript | Out-Null } catch { }
}

function Add-Report {
    <#
    .SYNOPSIS
        Adds one line to the report of the current step (and of the run) and prints it.
    .DESCRIPTION
        Status:
          Success      action done                     Failed       error
          AlreadyDone  target state already in place   Skipped      not processed (e.g. legacy server)
          Simulated    simulation done (-WhatIf)       Inventoried  current state read only
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Step,
        [Parameter(Mandatory)] [string]$Target,
        [Parameter(Mandatory)] [string]$Action,
        [Parameter(Mandatory)]
        [ValidateSet('Success','Failed','AlreadyDone','Skipped','Simulated','Inventoried')]
        [string]$Status,
        [string]$Phase = '',          # Before / After / ''
        [string]$BeforeValue = '',
        [string]$AfterValue  = '',
        [string]$Detail = '',
        [string]$ErrorMessage = ''
    )

    $entry = [PSCustomObject]@{
        Timestamp    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Step         = $Step
        Phase        = $Phase
        Target       = $Target
        Action       = $Action
        Status       = $Status
        BeforeValue  = $BeforeValue
        AfterValue   = $AfterValue
        Detail       = $Detail
        ErrorMessage = $ErrorMessage
        Mode         = $Script:Mode
    }
    $Script:ReportEntries.Add($entry)       | Out-Null
    $Script:GlobalReportEntries.Add($entry) | Out-Null

    # Console: status icon and tag, target, action, detail; the current value on a second dimmed line.
    $T = $Script:Theme
    $look = @{
        Success     = @('Ok',   'OK',     $T.Ok)
        AlreadyDone = @('Ok',   'DONE',   $T.Done)
        Inventoried = @('Info', 'INV',    $T.Info)
        Simulated   = @('Sim',  'SIM',    $T.Sim)
        Skipped     = @('Skip', 'SKIP',   $T.Warn)
        Failed      = @('Fail', 'FAILED', $T.Fail)
    }[$Status]
    $detailText = if ($ErrorMessage) { $ErrorMessage } elseif ($Detail) { $Detail } elseif ($BeforeValue) { $BeforeValue } else { '' }
    $detailText = ($detailText -replace '\r?\n', ' ').Trim()
    if ($detailText.Length -gt 110) { $detailText = $detailText.Substring(0, 107) + '...' }
    $targetText = if ($Target.Length -gt 22) { $Target.Substring(0, 21) + '~' } else { $Target }

    Write-ExHost '      ' -NoNewline
    Write-ExHost (Get-ExIcon $look[0]) -Color $look[2] -NoNewline
    Write-ExHost ('{0,-7}' -f $look[1]) -Color $look[2] -NoNewline
    Write-ExHost (' {0,-22} ' -f $targetText) -Color $T.Strong -NoNewline
    if ($detailText) {
        Write-ExHost $Action -NoNewline
        Write-ExHost ('  {0} {1}' -f $Script:Mid, $detailText) -Color $(if ($ErrorMessage) { $T.Fail } else { $T.Dim })
    } else {
        Write-ExHost $Action
    }
    if ($BeforeValue -and ($Detail -or $ErrorMessage)) {
        $bv = ($BeforeValue -replace '\r?\n', ' ').Trim()
        if ($bv.Length -gt 120) { $bv = $bv.Substring(0, 117) + '...' }
        Write-ExHost ([string]::new(' ', 6 + $Script:IconWidth + $Script:IconPad.Length + 8)) -NoNewline
        Write-ExHost ('{0} {1}' -f $Script:Arrow, $bv) -Color $T.Dim
    }
    $level = @{ Success = 'OK'; AlreadyDone = 'OK'; Inventoried = 'INFO'; Simulated = 'INFO'; Skipped = 'WARN'; Failed = 'ERROR' }[$Status]
    Write-ExLog $level ('{0} | {1} | {2} | {3} | {4}{5}' -f $Step, $Status, $Target, $Action, $Detail, $(if ($ErrorMessage) { " | Error: $ErrorMessage" } else { '' }))
}

function Save-Report {
    <#
    .SYNOPSIS
        Writes the CSV of the step, appends it to Report_GLOBAL.csv, refreshes the HTML run
        report and prints the step result line. Returns the path of the step CSV.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$StepName
    )

    $stepCsv = Join-Path $Script:OutputFolder ("Report_{0}_{1}.csv" -f $StepName, $Script:Mode)

    # Guard: when the step already saved its report, the orchestrator calls Save-Report again to
    # get the CSV path - return it without exporting or printing a second time.
    if ($Script:ReportSaved) { return $stepCsv }

    if ($Script:ReportEntries.Count -eq 0) {
        Write-Log 'No report entry to export.' -Level Debug
        return $null
    }

    $globalCsv = Join-Path $Script:OutputFolder 'Report_GLOBAL.csv'

    $Script:ReportEntries | Export-Csv -Path $stepCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'

    if (Test-Path $globalCsv) {
        $Script:ReportEntries | Export-Csv -Path $globalCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Append
    } else {
        $Script:ReportEntries | Export-Csv -Path $globalCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    }

    # One result line for the step: count per status, in a fixed order.
    $order  = 'Success', 'AlreadyDone', 'Inventoried', 'Simulated', 'Skipped', 'Failed'
    $labels = @{ Success = 'OK'; AlreadyDone = 'already done'; Inventoried = 'inventoried'; Simulated = 'simulated'; Skipped = 'skipped'; Failed = 'failed' }
    $groups = @{}
    foreach ($g in ($Script:ReportEntries | Group-Object Status)) { $groups[$g.Name] = $g.Count }
    $parts = foreach ($s in $order) { if ($groups[$s]) { '{0} {1}' -f $groups[$s], $labels[$s] } }
    $failed = [int]$groups['Failed']
    $status = if ($failed) { 'Fail' } elseif ($groups['Skipped']) { 'Warn' } else { 'Ok' }
    Write-ExHost ''
    Write-ExItem $status ('Result: {0}' -f ($parts -join (' {0} ' -f $Script:Mid))) -Icon 'Report'
    Write-ExItem Dim (Split-Path $stepCsv -Leaf) -Icon 'File'

    $htmlPath = Join-Path $Script:OutputFolder ("Report_GLOBAL_{0}.html" -f $Script:Mode)
    New-HtmlReport -Path $htmlPath

    $Script:ReportSaved = $true
    return $stepCsv
}

function Reset-Report {
    [CmdletBinding()] param()
    $Script:ReportEntries = New-Object System.Collections.Generic.List[object]
    $Script:ReportSaved   = $false
}

function Get-ExRunSummary {
    <# Counts of the run by status, the run folder, the HTML run report and the run log. #>
    $counts = @{}
    foreach ($g in ($Script:GlobalReportEntries | Group-Object Status)) { $counts[$g.Name] = $g.Count }
    [pscustomobject]@{
        Entries = $Script:GlobalReportEntries.Count
        Counts  = $counts
        Folder  = $Script:OutputFolder
        Html    = $(if ($Script:OutputFolder) { Join-Path $Script:OutputFolder ("Report_GLOBAL_{0}.html" -f $Script:Mode) })
        Log     = $Script:LogPath
        Start   = $Script:RunStart
    }
}

#endregion
#region 4. Execution modes -----------------------------------------------------------------------

function Invoke-Action {
    <#
    .SYNOPSIS
        Single wrapper for one action of a step, according to the current mode.
    .PARAMETER PreCheckScript
        Returns $true when the target state is already in place: the action is reported as
        AlreadyDone and is not run (idempotence).
    .PARAMETER ActionScript
        The change itself. In Simulate mode it runs with $WhatIfPreference = $true.
    .PARAMETER InventoryScript
        Returns the current state, without side effects (it runs up to twice in Apply mode:
        before and after the action).
    .OUTPUTS
        The status written to the report: Success, Failed, AlreadyDone, Skipped, Simulated or Inventoried.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Step,
        [Parameter(Mandatory)] [string]$Target,
        [Parameter(Mandatory)] [string]$Action,
        [scriptblock]$PreCheckScript = $null,
        [scriptblock]$ActionScript   = $null,
        [scriptblock]$InventoryScript= $null,
        [string]$Detail = ''
    )

    # Inventory mode: read the current state, change nothing.
    if ($Script:Mode -eq 'Inventory') {
        if ($InventoryScript) {
            try {
                $existing = & $InventoryScript
                $value = if ($existing -is [string]) { $existing } else { ($existing | Out-String).Trim() }
                Add-Report -Step $Step -Target $Target -Action $Action -Status Inventoried `
                           -Phase 'Before' -BeforeValue $value -Detail $Detail
                return 'Inventoried'
            } catch {
                Add-Report -Step $Step -Target $Target -Action $Action -Status Failed `
                           -Phase 'Before' -ErrorMessage $_.Exception.Message -Detail $Detail
                return 'Failed'
            }
        } else {
            Add-Report -Step $Step -Target $Target -Action $Action -Status Skipped `
                       -Detail 'No inventory script for this action'
            return 'Skipped'
        }
    }

    # State before the action (Apply / Simulate), from the inventory script when there is one.
    $snapBefore = ''
    if ($InventoryScript) {
        try {
            $snap = & $InventoryScript
            $snapBefore = if ($snap -is [string]) { $snap } else { ($snap | Out-String).Trim() }
        } catch {
            $snapBefore = "(state before not available: $($_.Exception.Message))"
        }
    }

    # Pre-check: is the target state already in place?
    if ($PreCheckScript) {
        try {
            $alreadyOk = & $PreCheckScript
            if ($alreadyOk) {
                Add-Report -Step $Step -Target $Target -Action $Action -Status AlreadyDone `
                           -BeforeValue $snapBefore -AfterValue $snapBefore -Detail $Detail
                return 'AlreadyDone'
            }
        } catch {
            Write-Log "Pre-check failed on $Target - $($_.Exception.Message)" -Level Debug
        }
    }

    # Simulate mode: -WhatIf when the script block supports it, otherwise the intent is reported.
    if ($Script:Mode -eq 'Simulate') {
        if ($ActionScript) {
            try {
                $WhatIfPreference = $true
                & $ActionScript
                $WhatIfPreference = $false
                Add-Report -Step $Step -Target $Target -Action $Action -Status Simulated `
                           -BeforeValue $snapBefore -AfterValue '(simulation - not applied)' `
                           -Detail $Detail
                return 'Simulated'
            } catch {
                $WhatIfPreference = $false
                Add-Report -Step $Step -Target $Target -Action $Action -Status Failed `
                           -BeforeValue $snapBefore -ErrorMessage $_.Exception.Message -Detail $Detail
                return 'Failed'
            }
        }
    }

    # Apply mode: real change.
    if ($Script:Mode -eq 'Apply') {
        if ($ActionScript) {
            try {
                & $ActionScript
                # State after the change.
                $snapAfter = ''
                if ($InventoryScript) {
                    try {
                        $snap = & $InventoryScript
                        $snapAfter = if ($snap -is [string]) { $snap } else { ($snap | Out-String).Trim() }
                    } catch {
                        $snapAfter = "(state after not available: $($_.Exception.Message))"
                    }
                }
                Add-Report -Step $Step -Target $Target -Action $Action -Status Success `
                           -Phase 'After' -BeforeValue $snapBefore -AfterValue $snapAfter -Detail $Detail
                return 'Success'
            } catch {
                Add-Report -Step $Step -Target $Target -Action $Action -Status Failed `
                           -BeforeValue $snapBefore -ErrorMessage $_.Exception.Message -Detail $Detail
                return 'Failed'
            }
        }
    }

    Add-Report -Step $Step -Target $Target -Action $Action -Status Skipped -Detail 'No action to run'
    return 'Skipped'
}

#endregion
#region 5. Exchange helpers ----------------------------------------------------------------------

function Set-DefaultExchangeServer {
    <#
    .SYNOPSIS
        Stores the Exchange server and the credentials used by Initialize-ExchangeShell when the
        local snap-in is not available. Called by the orchestrator before the step loop.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Server,
        [System.Management.Automation.PSCredential]$Credential
    )
    $Script:DefaultExchangeServer  = $Server
    $Script:DefaultCredential      = $Credential
    $Script:ExchangeRemoteSession  = $null
    Write-ExLog 'DEBUG' "Default Exchange server for remoting: $Server"
}

function Get-ExchangeRemoteSession {
    [CmdletBinding()] param()
    return $Script:ExchangeRemoteSession
}

function Test-ExchangeShellLoaded {
    [CmdletBinding()] param()
    # The local snap-in (Exchange 2019 after CU1) only provides Get-ExchangeServer.
    # Get-DatabaseAvailabilityGroup is also required, to confirm that the full remoting session
    # with all the Exchange cmdlets is loaded.
    return [bool](Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue) -and
           [bool](Get-Command Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue)
}

function Initialize-ExchangeShell {
    <#
    .SYNOPSIS
        Loads the Exchange cmdlets: remoting session to an Exchange 2019 server, or the local snap-in.
    .DESCRIPTION
        Without -ConnectToServer, uses the server stored by Set-DefaultExchangeServer (the local
        computer or the server given with -ExchangeServer to the orchestrator).
    #>
    [CmdletBinding()]
    param(
        [string]$ConnectToServer,
        [System.Management.Automation.PSCredential]$Credential
    )

    if (Test-ExchangeShellLoaded) {
        Write-Log 'Exchange cmdlets already loaded' -Level Debug
        Invoke-ExchangeCmdletsGlobalExport
        return
    }

    # Target server: parameter, otherwise the default stored by Set-DefaultExchangeServer.
    $targetServer  = if ($ConnectToServer)                  { $ConnectToServer }
                     elseif ($Script:DefaultExchangeServer) { $Script:DefaultExchangeServer }
                     else { $null }

    $effectiveCred = if ($Credential)                   { $Credential }
                     elseif ($Script:DefaultCredential) { $Script:DefaultCredential }
                     else { $null }

    # A known target server: PowerShell remoting directly.
    if ($targetServer) {
        $uri = "http://$targetServer/PowerShell/"
        Write-Log "Exchange remoting session: $uri" -Level Info
        $params = @{
            ConfigurationName = 'Microsoft.Exchange'
            ConnectionUri     = $uri
            Authentication    = 'Kerberos'
            AllowRedirection  = $true
            ErrorAction       = 'Stop'
        }
        if ($effectiveCred) { $params['Credential'] = $effectiveCred }
        $session = New-PSSession @params
        $Script:ExchangeRemoteSession = $session
        $tmpModule = Import-PSSession $session -DisableNameChecking
        # Import-Module -Global fails with the Exchange remote PowerShell proxy (no -AllowClobber):
        # the module returned by Import-PSSession is passed directly to force the global scope.
        Invoke-ExchangeCmdletsGlobalExport -Module $tmpModule
        return
    }

    # No explicit server: try the local snap-in (2019, then 2013).
    foreach ($snap in 'Microsoft.Exchange.Management.PowerShell.SnapIn',
                       'Microsoft.Exchange.Management.PowerShell.E2010') {
        try {
            Add-PSSnapin $snap -ErrorAction Stop
            Write-Log "Snap-in $snap loaded locally" -Level Debug
            Invoke-ExchangeCmdletsGlobalExport
            return
        } catch { }
    }

    throw 'Cannot load the Exchange cmdlets (no server to connect to and no local snap-in). Use -ExchangeServer to give an Exchange 2019 server reachable with PowerShell remoting.'
}

function Invoke-ExchangeCmdletsGlobalExport {
    <#
    .SYNOPSIS
        Makes the Exchange proxy functions visible to the steps (global scope).
    .DESCRIPTION
        PowerShell 5.1: the Exchange proxy functions loaded by Import-PSSession or by the snap-in
        (implicit remoting) live in the private scope of their module. Import-Module -Global called
        from a module function does not make them visible to the script blocks of the steps
        (defined outside the module). They are therefore registered in the global scope through the
        Function: provider. The source module is the one that exports Get-DatabaseAvailabilityGroup.
        -Module receives the module returned by Import-PSSession directly, without calling
        Import-Module -Global (which fails with Exchange remote PowerShell in PowerShell 5.1).
    #>
    [CmdletBinding()]
    param(
        [System.Management.Automation.PSModuleInfo]$Module
    )

    $exchangeMod = $Module

    if (-not $exchangeMod) {
        $exchangeMod = Get-Module | Where-Object {
            $_.ExportedFunctions.ContainsKey('Get-DatabaseAvailabilityGroup')
        } | Select-Object -First 1
    }

    if (-not $exchangeMod) {
        # Get-Module -All also finds the modules loaded in a private (module) scope.
        $exchangeMod = Get-Module -All | Where-Object {
            $_.ExportedFunctions.ContainsKey('Get-DatabaseAvailabilityGroup')
        } | Select-Object -First 1
    }

    if (-not $exchangeMod) {
        # Fallback: the largest tmp_* module (Exchange remoting without DAG cmdlets).
        $exchangeMod = Get-Module -All | Where-Object {
            $_.ModuleType -eq 'Script' -and $_.Name -like 'tmp_*' -and $_.ExportedFunctions.Count -gt 5
        } | Sort-Object { $_.ExportedFunctions.Count } -Descending | Select-Object -First 1
    }

    if (-not $exchangeMod) { return }

    # No Get-Command check: Get-Command runs in the scope of this module, where the Exchange cmdlets
    # are already visible, so it would always find them and prevent the global export. -Force always.
    $forcedCount = 0
    foreach ($kv in $exchangeMod.ExportedFunctions.GetEnumerator()) {
        try {
            New-Item -Path "Function:Global:$($kv.Key)" -Value $kv.Value.ScriptBlock -Force | Out-Null
            $forcedCount++
        } catch { }
    }
    if ($forcedCount -gt 0) {
        Write-Log "$forcedCount Exchange cmdlets made global (module $($exchangeMod.Name))" -Level Debug
    }
}

function Get-Exchange2019Servers {
    <#
    .SYNOPSIS
        Exchange 2019 servers only (Mailbox role), optionally limited to -Names.
    .DESCRIPTION
        Edge Transport servers (perimeter network, outside the domain) have ServerRole 'Edge' and are
        excluded: IIS, DAG, database and log steps must never target them. This strict filter is
        what keeps the Exchange 2013/2016 servers out of every change.
    #>
    [CmdletBinding()]
    param([string[]]$Names)

    $all = Get-ExchangeServer -ErrorAction Stop |
           Where-Object { $_.AdminDisplayVersion -like 'Version 15.2*' -and
                          $_.ServerRole -notlike '*Edge*' }

    if ($Names) {
        $all = $all | Where-Object { $_.Name -in $Names }
    }
    return $all
}

function Get-Exchange2013Servers {
    [CmdletBinding()] param()
    Get-ExchangeServer -ErrorAction Stop |
        Where-Object { $_.AdminDisplayVersion -like 'Version 15.0*' }
}

function Get-Exchange2016Servers {
    [CmdletBinding()] param()
    Get-ExchangeServer -ErrorAction Stop |
        Where-Object { $_.AdminDisplayVersion -like 'Version 15.1*' }
}

function Get-LegacyExchangeServers {
    <#
    .SYNOPSIS
        Every Exchange server older than 2019 (2013 = 15.0.*, 2016 = 15.1.*).
    .DESCRIPTION
        Single source for the migration, certificate and receive connector steps in mixed
        2013/2016 organisations (replaces direct calls to Get-Exchange2013Servers).
    #>
    [CmdletBinding()] param()
    Get-ExchangeServer -ErrorAction Stop |
        Where-Object { $_.AdminDisplayVersion -like 'Version 15.0*' -or
                       $_.AdminDisplayVersion -like 'Version 15.1*' }
}

function Test-IsExchange2019Server {
    [CmdletBinding()]
    param([Parameter()] [string]$ServerName)
    if (-not $ServerName) { return $false }
    $srv = Get-ExchangeServer $ServerName -ErrorAction SilentlyContinue
    return ($srv -and $srv.AdminDisplayVersion -like 'Version 15.2*')
}

#endregion
#region 6. Configuration files and resumable state ---------------------------------------------

function Import-ConfigCsv {
    <#
    .SYNOPSIS
        Robust Import-Csv: trims the header names and the values, removes the optional first
        "VersionNumber" / "SchemaVersion" line of the historical CSV files.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Path,
        [string]$Delimiter = ','
    )

    if (-not (Test-Path $Path)) {
        throw "CSV file not found: $Path"
    }

    $rows = Import-Csv -Path $Path -Delimiter $Delimiter

    # Clean the header names (stray spaces in the historical CSV files).
    $cleaned = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        $obj = [ordered]@{}
        foreach ($p in $row.PSObject.Properties) {
            $key = $p.Name.Trim()
            $obj[$key] = if ($p.Value -is [string]) { $p.Value.Trim() } else { $p.Value }
        }
        $cleaned.Add([PSCustomObject]$obj) | Out-Null
    }

    # Remove the version header line when present (compatibility with the old CSV files).
    $cleaned = $cleaned | Where-Object {
        ($_.PSObject.Properties | Select-Object -First 1 -ExpandProperty Value) -ne 'VersionNumber' -and
        ($_.PSObject.Properties | Select-Object -First 1 -ExpandProperty Value) -ne 'SchemaVersion'
    }

    return ,$cleaned
}

function Get-StateFile {
    [CmdletBinding()] param()
    # The resume state stays in the base folder, shared by all the runs.
    $base = if ($Script:OutputBaseFolder) { $Script:OutputBaseFolder } else { $Script:OutputFolder }
    return (Join-Path $base 'DeploymentState.json')
}

function Get-DeploymentState {
    [CmdletBinding()] param()
    $path = Get-StateFile
    if (-not (Test-Path $path)) { return @{} }
    try {
        $json = Get-Content $path -Raw | ConvertFrom-Json
        # -AsHashtable does not exist in PowerShell 5.1: manual conversion.
        $ht = @{}
        foreach ($prop in $json.PSObject.Properties) { $ht[$prop.Name] = $prop.Value }
        return $ht
    }
    catch { return @{} }
}

function Set-StepState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$StepName,
        [Parameter(Mandatory)]
        [ValidateSet('Pending','InProgress','Completed','Failed')]
        [string]$Status,
        [string]$Detail = ''
    )

    $state = Get-DeploymentState
    if ($state -isnot [hashtable]) { $state = @{} }
    $state[$StepName] = @{
        Status     = $Status
        Detail     = $Detail
        LastUpdate = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Mode       = $Script:Mode
    }
    $state | ConvertTo-Json -Depth 5 | Set-Content -Path (Get-StateFile) -Encoding UTF8
}

#endregion

#region 7. HTML reports --------------------------------------------------------------------------
# -------------------------------------------------------------------------------------------------
# Every HTML report of the framework (run report, system mailbox plan, migration plan, migration
# status, HealthChecker issues, protocol-log analysis) uses the theme below: the same look as the
# Purview DLP Report files. Light or dark theme follows the system (or ?scoutTheme=light|dark).
# Reports are single self-contained files: no external font, script or image.
# -------------------------------------------------------------------------------------------------

$Script:HtmlStyle = @'
:root {
  color-scheme: light;
  --cp-bg: #f7f4ef; --cp-surface: #ffffff; --cp-surface-soft: #f5f5f5; --cp-border: #dedede;
  --cp-text: #242424; --cp-text-muted: #5c5c5c;
  --cp-accent: #b11f4b; --cp-accent-hover: #9a1a41; --cp-accent-soft: rgba(177, 31, 75, 0.08); --cp-accent-fg: #ffffff;
  --cp-success: #16a34a; --cp-danger: #dc2626; --cp-warning: #d97706; --cp-info: #0078d4; --cp-violet: #7c3aed; --cp-teal: #0d9488;
  --cp-link: #0078d4; --cp-shadow: 0 18px 48px rgba(0, 0, 0, 0.12); --cp-highlight: rgba(177, 31, 75, 0.12);
  --tone-soft: 0.10;
}
html[data-theme="dark"] {
  color-scheme: dark;
  --cp-bg: #3d3b3a; --cp-surface: #292929; --cp-surface-soft: #2e2e2e; --cp-border: #474747;
  --cp-text: #dedede; --cp-text-muted: #a3a3a3;
  --cp-accent: #fd8ea1; --cp-accent-hover: #fb7b91; --cp-accent-soft: rgba(253, 142, 161, 0.14); --cp-accent-fg: #1a1a1a;
  --cp-success: #4ade80; --cp-danger: #f87171; --cp-warning: #fbbf24; --cp-info: #4da6ff; --cp-violet: #a78bfa; --cp-teal: #2dd4bf;
  --cp-link: #4da6ff; --cp-shadow: 0 18px 48px rgba(0, 0, 0, 0.32); --cp-highlight: rgba(253, 142, 161, 0.12);
}
* { box-sizing: border-box; }
body { margin: 0; padding: 28px 32px; background: var(--cp-bg); color: var(--cp-text); font: 14px "Segoe UI", Aptos, Calibri, -apple-system, BlinkMacSystemFont, sans-serif; }
h1 { font-size: 30px; font-weight: 650; letter-spacing: -0.025em; margin: 8px 0 12px; }
h2 { font-size: 18px; font-weight: 600; margin: 0 0 12px; }
h3 { font-size: 15px; font-weight: 600; margin: 16px 0 8px; }
p { line-height: 1.5; margin: 8px 0; }
a { color: var(--cp-link); text-decoration: none; } a:hover { text-decoration: underline; }
code, .mono { font-family: Consolas, "Cascadia Mono", monospace; font-size: 12px; }
.muted, small { color: var(--cp-text-muted); }
[hidden] { display: none !important; }

/* Hero header ---------------------------------------------------------------------------------- */
header.hero, .hdr { position: relative; overflow: hidden; border: 1px solid var(--cp-border); border-top: 4px solid var(--cp-accent); border-radius: 16px; padding: 24px; margin-bottom: 20px; background: linear-gradient(125deg, var(--cp-surface) 35%, var(--cp-accent-soft)); color: var(--cp-text); box-shadow: none; position: relative; }
header.hero::before, .hdr::before { content: ""; position: absolute; width: 280px; height: 280px; right: -100px; top: -160px; border: 40px solid var(--cp-highlight); border-radius: 50%; pointer-events: none; }
.hero-top { position: relative; display: flex; justify-content: space-between; gap: 24px; align-items: flex-start; flex-wrap: wrap; }
.origin { color: var(--cp-accent); font-size: 11px; letter-spacing: 0.1em; text-transform: uppercase; font-weight: 700; }
.subtitle { color: var(--cp-text-muted); max-width: 720px; }
.period { position: relative; min-width: 280px; padding: 12px 16px; border: 1px solid var(--cp-border); border-radius: 0.625rem; background: var(--cp-surface); font-size: 13px; line-height: 1.6; }
.label { font-size: 11px; color: var(--cp-accent); font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; display: block; margin-bottom: 8px; }
.caption { font-size: 12px; color: var(--cp-text-muted); margin: 0 0 12px; }

/* Metric tiles --------------------------------------------------------------------------------- */
.metrics { position: relative; display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 16px; margin: 20px 0 16px; }
.tile { position: relative; overflow: hidden; padding: 18px 20px; border: 1px solid var(--cp-border); border-radius: 16px; background: var(--cp-surface); color: var(--cp-text); transition: transform 160ms ease, border-color 160ms ease; }
.tile:hover { transform: translateY(-3px); border-color: var(--cp-accent); }
.tile::after { content: ""; position: absolute; width: 80px; height: 80px; right: -28px; bottom: -32px; border: 14px solid var(--cp-highlight); border-radius: 50%; pointer-events: none; }
.tile-primary { background: linear-gradient(125deg, var(--cp-accent-hover), var(--cp-accent)); border-color: var(--cp-accent); color: var(--cp-accent-fg); }
.tile-tinted { background: linear-gradient(135deg, var(--cp-surface), var(--cp-accent-soft)); }
.tile-label { display: block; font-size: 12px; font-weight: 600; margin-bottom: 12px; }
.tile strong { display: block; font-size: 32px; font-weight: 650; letter-spacing: -0.025em; font-variant-numeric: tabular-nums; }
.tile small { display: block; margin-top: 8px; font-size: 11px; } .tile-primary small { color: var(--cp-accent-fg); }
.tile-icon { position: absolute; top: 16px; right: 16px; width: 20px; height: 20px; opacity: 0.7; }
.tile[style*="--tone"] { border-top: 4px solid var(--tone); }
.tile[style*="--tone"] strong, .tile[style*="--tone"] .tile-n { color: var(--tone); }
/* Compact tiles of the detailed reports (.tiles / .tile-n / .tile-l) */
.tiles { display: grid; grid-template-columns: repeat(auto-fill, minmax(170px, 1fr)); gap: 14px; padding: 16px 20px 20px; position: relative; }
.tiles .tile { text-align: left; min-width: 0; box-shadow: none; }
.tile-n { font-size: 28px; font-weight: 650; line-height: 1.1; letter-spacing: -0.02em; font-variant-numeric: tabular-nums; }
.tile-l { font-size: 12px; font-weight: 600; margin-top: 6px; color: var(--cp-text-muted); }
.tile-p { font-size: 11px; margin-top: 4px; color: var(--cp-text-muted); }
.tile.clickable { cursor: pointer; user-select: none; }
.tile.dimmed { opacity: 0.45; }
.tile.tile-active { border-color: var(--cp-accent); box-shadow: 0 0 0 2px var(--cp-accent); }
.tile-srv .tile-hdr { font-size: 12px; font-weight: 700; margin-bottom: 8px; }
.tile-srv .tile-row { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; font-variant-numeric: tabular-nums; }
.tile-srv .tile-row .lbl { font-size: 11px; font-weight: 600; color: var(--cp-text-muted); text-transform: uppercase; letter-spacing: 0.06em; }
.tile-srv .tile-row .val { font-size: 20px; font-weight: 650; }

/* Status distribution --------------------------------------------------------------------------- */
.distribution { position: relative; display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 16px; border-top: 1px solid var(--cp-border); padding-top: 12px; }
.band-head { display: flex; justify-content: space-between; font-size: 11px; gap: 8px; color: var(--cp-text-muted); }
.track { height: 5px; border-radius: 8px; background: var(--cp-border); margin-top: 8px; overflow: hidden; }
.fill { height: 100%; border-radius: 8px; background: var(--cp-accent); }

/* Header compatibility (.hdr-*) of the detailed reports ----------------------------------------- */
.hdr-r1 { position: relative; display: flex; align-items: flex-start; gap: 16px; flex-wrap: wrap; }
.hdr-title { font-size: 28px; font-weight: 650; letter-spacing: -0.025em; }
.hdr-meta { margin-left: auto; min-width: 260px; padding: 10px 14px; border: 1px solid var(--cp-border); border-radius: 0.625rem; background: var(--cp-surface); font-size: 12px; line-height: 1.6; color: var(--cp-text); text-align: left; opacity: 1; }
.hdr-r2 { position: relative; display: flex; gap: 8px; flex-wrap: wrap; margin-top: 16px; }
.hdr-pill { border-radius: 999px; padding: 3px 12px; font-size: 12px; font-weight: 600; color: #fff; border: 0; }

/* Panels --------------------------------------------------------------------------------------- */
.wrap { margin: 0; padding: 0; max-width: none; }
section.panel { background: var(--cp-surface); border: 1px solid var(--cp-border); border-radius: 16px; padding: 20px; margin-bottom: 20px; }
.card { background: var(--cp-surface); border: 1px solid var(--cp-border); border-radius: 16px; margin-bottom: 20px; overflow: hidden; box-shadow: none; }
.card-hdr { padding: 16px 20px 4px; font-size: 11px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: var(--cp-accent); background: transparent; border: 0; }
.card-body { padding: 12px 20px 20px; }
.note { color: var(--cp-text-muted); font-size: 12px; margin: 6px 0; }
.sub-title { margin: 16px 0 8px; font-weight: 600; font-size: 13px; padding: 4px 10px; border-left: 3px solid var(--cp-accent); background: var(--cp-accent-soft); border-radius: 0 8px 8px 0; }
.sub-title .sub-meta { color: var(--cp-text-muted); font-weight: 400; margin-left: 8px; }

/* Controls ------------------------------------------------------------------------------------- */
input, button, select { font: inherit; color: var(--cp-text); background: var(--cp-surface); border: 1px solid var(--cp-border); padding: 7px 10px; border-radius: 0.625rem; }
button { cursor: pointer; } button:hover { color: var(--cp-accent); border-color: var(--cp-accent); }
:focus-visible { outline: 2px solid var(--cp-accent); outline-offset: 2px; }
.primary-button { background: var(--cp-accent); color: var(--cp-accent-fg); border-color: var(--cp-accent); }
.toolbar, .controls { display: flex; gap: 8px 12px; flex-wrap: wrap; align-items: center; margin: 12px 0; padding: 0; border: 0; }
.btn { font-size: 13px; padding: 6px 12px; }
.filter-state { display: none; padding: 10px 20px; border-top: 1px solid var(--cp-border); font-size: 13px; font-weight: 600; }
.filter-state.on { display: flex; align-items: center; gap: 10px; }
.filter-state .clear { font-size: 12px; padding: 3px 10px; }

/* Step list ------------------------------------------------------------------------------------ */
.toc { display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 8px 16px; margin: 12px 0 0; padding: 0; list-style: none; }
.toc li, .toc-item { display: flex; align-items: center; gap: 10px; padding: 8px 12px; border: 1px solid var(--cp-border); border-radius: 0.625rem; background: var(--cp-surface); }
.toc li:hover, .toc-item:hover { border-color: var(--cp-accent); }
.toc-name { flex: 1; min-width: 0; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--cp-text); }
.toc-ok { color: var(--cp-success); font-weight: 700; } .toc-ko { color: var(--cp-danger); font-weight: 700; }
.card > .toolbar, .card > .toc { padding-left: 20px; padding-right: 20px; } .card > .toc { padding-bottom: 20px; }
.dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; background: var(--tone, var(--cp-text-muted)); }

/* Accordions ------------------------------------------------------------------------------------ */
details { border: 1px solid var(--cp-border); border-radius: 12px; background: var(--cp-surface); margin-bottom: 10px; overflow: hidden; box-shadow: none; }
summary { display: flex; align-items: center; gap: 10px; padding: 12px 16px; cursor: pointer; user-select: none; list-style: none; background: var(--cp-surface); border: 0; }
summary::-webkit-details-marker { display: none; }
summary:hover { background: var(--cp-accent-soft); }
details[open] > summary { border-bottom: 1px solid var(--cp-border); background: var(--cp-accent-soft); }
.acc-icon { display: inline-block; font-size: 10px; color: var(--cp-text-muted); transition: transform 0.2s; width: 12px; text-align: center; }
details[open] .acc-icon { transform: rotate(90deg); }
.acc-name, .acc-step { font-weight: 600; flex: 1; min-width: 0; color: var(--cp-text); }
.acc-summary, .acc-cnt { font-size: 12px; color: var(--cp-text-muted); margin-left: auto; white-space: nowrap; }
.acc-pills { display: flex; gap: 6px; flex-wrap: wrap; }
.srv-body { padding: 14px 16px; }
details.filter-hidden { display: none; }

/* Tables --------------------------------------------------------------------------------------- */
.tbl-wrap { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
thead th { background: var(--cp-surface-soft); color: var(--cp-text); font-size: 12px; font-weight: 600; letter-spacing: 0; text-align: left; padding: 10px; border-bottom: 1px solid var(--cp-border); white-space: nowrap; }
td { padding: 8px 10px; border-bottom: 1px solid var(--cp-border); vertical-align: top; }
tbody tr:hover td { background: var(--cp-accent-soft); }
tr:last-child td { border-bottom: none; }
tr.total td { font-weight: 700; background: var(--cp-surface-soft); }
th.num, td.num, .num { text-align: right; font-variant-numeric: tabular-nums; }
td.zero { color: var(--cp-text-muted); }
.val { word-break: break-word; white-space: pre-wrap; font-size: 12px; }
.em { color: var(--cp-text-muted); }
.er { color: var(--cp-danger); word-break: break-word; font-size: 12px; }
.td-ts { white-space: nowrap; color: var(--cp-text-muted); font-size: 12px; font-variant-numeric: tabular-nums; }
.td-tgt { font-weight: 600; }

/* Badges --------------------------------------------------------------------------------------- */
.bdg, .pill { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 11px; font-weight: 600; color: #fff; white-space: nowrap; background: var(--tone, var(--cp-text-muted)); }
.pill-soft { background: var(--cp-accent-soft); color: var(--cp-text); }
.bdg-2013 { background: var(--cp-danger); } .bdg-2016 { background: var(--cp-warning); } .bdg-2019 { background: var(--cp-success); } .bdg-other { background: var(--cp-text-muted); }

#totop { position: fixed; bottom: 22px; right: 22px; width: 40px; height: 40px; border-radius: 50%; background: var(--cp-accent); color: var(--cp-accent-fg); border: 0; font-size: 18px; box-shadow: var(--cp-shadow); opacity: 0.85; display: flex; align-items: center; justify-content: center; }
#totop:hover { opacity: 1; color: var(--cp-accent-fg); }
footer { color: var(--cp-text-muted); font-size: 12px; line-height: 1.6; margin-top: 8px; }
@media (max-width: 900px) { body { padding: 16px; } .metrics { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
@media print { body { background: #fff; } #totop, .toolbar, .controls { display: none; } details { break-inside: avoid; } }
'@

$Script:HtmlThemeScript = @'
<script>
  (function () {
    var param = new URLSearchParams(window.location.search).get("scoutTheme");
    var theme = (param === "light" || param === "dark") ? param : ((window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) ? "dark" : "light");
    document.documentElement.setAttribute("data-theme", theme);
  })();
</script>
'@

$Script:HtmlIcons = @{
    list   = '<path d="M9 6h11M9 12h11M9 18h11"/><path d="m3.5 6 1.5 1.5L8 4.5M3.5 12l1.5 1.5L8 10.5M3.5 18l1.5 1.5L8 16.5"/>'
    check  = '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>'
    change = '<path d="M20 11a8 8 0 0 0-14.3-4.9L4 8"/><path d="M4 4v4h4"/><path d="M4 13a8 8 0 0 0 14.3 4.9L20 16"/><path d="M20 20v-4h-4"/>'
    alert  = '<path d="M12 3 2 20h20z"/><path d="M12 10v4M12 17v.5"/>'
    server = '<rect x="4" y="3" width="16" height="7" rx="1.5"/><rect x="4" y="14" width="16" height="7" rx="1.5"/><path d="M8 6.5h.01M8 17.5h.01"/>'
    mail   = '<rect x="3" y="5" width="18" height="14" rx="3"/><path d="m3 6 9 7 9-7"/>'
    chart  = '<path d="M4 20V10m8 10V4m8 16V7"/>'
    clock  = '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>'
    user   = '<circle cx="12" cy="8" r="4"/><path d="M4 21v-2a8 8 0 0 1 16 0v2"/>'
}

function ConvertTo-SafeHtml {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return '' }
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'", '&#39;'
}

function Get-ExStatusColor {
    <# CSS colour (theme variable) of a report status, used by badges, tiles and bars. #>
    param([string]$Status)
    switch ($Status) {
        'Success'     { 'var(--cp-success)' }
        'AlreadyDone' { 'var(--cp-teal)' }
        'Inventoried' { 'var(--cp-info)' }
        'Simulated'   { 'var(--cp-violet)' }
        'Skipped'     { 'var(--cp-warning)' }
        'Failed'      { 'var(--cp-danger)' }
        default       { 'var(--cp-text-muted)' }
    }
}

function Get-ExHtmlHead {
    <#
    .SYNOPSIS
        Start of every HTML report: doctype, theme script, shared style, then <body>.
    .PARAMETER RefreshSeconds
        Adds <meta http-equiv="refresh"> (migration follow-up report).
    #>
    param([Parameter(Mandatory)][string]$Title, [int]$RefreshSeconds = 0, [string]$ExtraStyle = '')
    $refresh = if ($RefreshSeconds -gt 0) { "<meta http-equiv=""refresh"" content=""$RefreshSeconds"">" } else { '' }
    return @"
<!doctype html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">$refresh
<title>$(ConvertTo-SafeHtml $Title)</title>
$Script:HtmlThemeScript
<style>
$Script:HtmlStyle
$ExtraStyle
</style></head><body>
"@
}

function Get-ExHtmlTile {
    <#
    .SYNOPSIS
        One metric tile: label, value, note. -Kind Primary (accent), Tinted or Plain;
        -Tone gives a status colour to a plain tile (top border and value).
    #>
    param([Parameter(Mandatory)][string]$Label, [Parameter(Mandatory)][string]$Value, [string]$Note,
        [ValidateSet('Primary', 'Tinted', 'Plain')][string]$Kind = 'Plain', [string]$Tone, [string]$Icon)
    $class = @{ Primary = 'tile tile-primary'; Tinted = 'tile tile-tinted'; Plain = 'tile' }[$Kind]
    $style = if ($Tone) { " style=""--tone:$Tone""" } else { '' }
    $svg = if ($Icon -and $Script:HtmlIcons[$Icon]) { '<svg class="tile-icon" aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">' + $Script:HtmlIcons[$Icon] + '</svg>' } else { '' }
    $noteHtml = if ($Note) { '<small>' + (ConvertTo-SafeHtml $Note) + '</small>' } else { '' }
    return ('<article class="{0}"{1}><span class="tile-label">{2}</span>{3}<strong>{4}</strong>{5}</article>' -f $class, $style, (ConvertTo-SafeHtml $Label), $svg, (ConvertTo-SafeHtml $Value), $noteHtml)
}

function Get-ExHtmlHero {
    <#
    .SYNOPSIS
        Header of a report: origin label, title, subtitle, a framed box on the right (period,
        run...), metric tiles and optional extra HTML (distribution bars, pills).
    .PARAMETER BoxHtml
        HTML content of the right box (already encoded).
    #>
    param([string]$Origin = 'Exchange 2013/2016 &rarr; 2019 migration', [Parameter(Mandatory)][string]$Title, [string]$Subtitle,
        [string]$BoxLabel, [string]$BoxHtml, [string[]]$Tiles, [string]$ExtraHtml)
    $box = if ($BoxLabel -or $BoxHtml) { '<div class="period"><span class="label">' + (ConvertTo-SafeHtml $BoxLabel) + '</span>' + $BoxHtml + '</div>' } else { '' }
    $tilesHtml = if ($Tiles) { '<div class="metrics">' + ($Tiles -join '') + '</div>' } else { '' }
    $sub = if ($Subtitle) { '<p class="subtitle">' + $Subtitle + '</p>' } else { '' }
    return ('<header class="hero"><div class="hero-top"><div><span class="origin">{0}</span><h1>{1}</h1>{2}</div>{3}</div>{4}{5}</header>' -f $Origin, (ConvertTo-SafeHtml $Title), $sub, $box, $tilesHtml, $ExtraHtml)
}

function Get-ExHtmlFooter {
    <# Footer line of every report: what the file contains and which version produced it. #>
    param([string]$Text)
    $info = Get-ExToolInfo
    $generated = 'Generated {0} by Exchange 2013/2016 to 2019 Migration {1} (Deploy-Exchange2019.ps1).' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $info.Version
    return '<footer>' + $(if ($Text) { $Text + ' ' } else { '' }) + $generated + '</footer><button id="totop" type="button" title="Back to top" onclick="window.scrollTo({top:0,behavior:''smooth''})">&uarr;</button></body></html>'
}

function New-HtmlReport {
    <#
    .SYNOPSIS
        Run report Report_GLOBAL_<Mode>.html: every action of the run, step by step.
    .DESCRIPTION
        Rewritten after each step, so that it always reflects the run so far. Header with the run
        and the counts by status, a status distribution, the list of steps, then one panel per
        step with its actions (target, action, status, value before, value after, detail).
        A search box and a status filter narrow the tables.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Path
    )

    if ($Script:GlobalReportEntries.Count -eq 0) { return }

    $entries = $Script:GlobalReportEntries
    $mode    = $Script:Mode
    $total   = $entries.Count
    $order   = 'Success', 'AlreadyDone', 'Inventoried', 'Simulated', 'Skipped', 'Failed'
    $labels  = @{ Success = 'Changed'; AlreadyDone = 'Already done'; Inventoried = 'Inventoried'; Simulated = 'Simulated'; Skipped = 'Skipped'; Failed = 'Failed' }
    $counts  = @{}
    foreach ($g in ($entries | Group-Object Status)) { $counts[$g.Name] = $g.Count }
    $byStep  = @($entries | Group-Object Step | Sort-Object Name)
    $failed  = [int]$counts['Failed']
    $changed = [int]$counts['Success'] + [int]$counts['Simulated']
    $steady  = [int]$counts['AlreadyDone'] + [int]$counts['Inventoried']
    $targets = @($entries | Select-Object -ExpandProperty Target -Unique).Count
    $server  = if ($Script:DefaultExchangeServer) { $Script:DefaultExchangeServer } else { $env:COMPUTERNAME }
    # Step label: "Step 03 - title of the catalogue" when the report step starts with its number.
    $catalog = Get-DeploymentStepCatalog
    $stepLabel = {
        param([string]$Name)
        if ($Name -match '^(\d+)') {
            $entry = $catalog[[string][int]$Matches[1]]
            if ($entry) { return ('Step {0:D2} {1} {2}' -f [int]$Matches[1], [char]0x00B7, $entry.Title) }
        }
        return $Name
    }

    $sb = New-Object System.Text.StringBuilder 200000
    [void]$sb.Append((Get-ExHtmlHead -Title ("Exchange 2019 migration | Run {0}" -f $Script:TimestampRun)))

    # ---- Header ----------------------------------------------------------------------------------
    $tiles = @(
        (Get-ExHtmlTile -Label 'Actions' -Value (Format-ExNumber $total) -Note ('{0} step(s), {1} target(s)' -f $byStep.Count, $targets) -Kind Primary -Icon 'list')
        (Get-ExHtmlTile -Label $(if ($mode -eq 'Simulate') { 'Simulated changes' } else { 'Changes applied' }) -Value (Format-ExNumber $changed) -Note $(if ($mode -eq 'Inventory') { 'Inventory mode changes nothing' } else { 'Success or simulated' }) -Kind Tinted -Icon 'change')
        (Get-ExHtmlTile -Label 'Read or already in place' -Value (Format-ExNumber $steady) -Note 'Inventoried or already done' -Kind Tinted -Icon 'check')
        (Get-ExHtmlTile -Label 'Failed' -Value (Format-ExNumber $failed) -Note ('{0} skipped' -f [int]$counts['Skipped']) -Tone $(if ($failed) { 'var(--cp-danger)' } else { 'var(--cp-success)' }) -Icon 'alert')
    )
    $bands = foreach ($s in $order) {
        if (-not $counts[$s]) { continue }
        $pct = [Math]::Round(100.0 * $counts[$s] / $total)
        '<div><div class="band-head"><span>{0}</span><span>{1} / {2}%</span></div><div class="track"><div class="fill" style="width:{2}%;background:{3}"></div></div></div>' -f $labels[$s], (Format-ExNumber $counts[$s]), $pct, (Get-ExStatusColor $s)
    }
    $box = 'Started {0}<br>Run <code>{1}</code><br>Mode <strong>{2}</strong> &middot; server {3}' -f $Script:RunStart.ToString('yyyy-MM-dd HH:mm:ss'), $Script:TimestampRun, $mode, (ConvertTo-SafeHtml $server)
    $subtitle = 'Every action of the run, step by step: the target, the status, the value before and after. Run folder: <code>{0}</code>' -f (ConvertTo-SafeHtml $Script:OutputFolder)
    [void]$sb.Append((Get-ExHtmlHero -Title ('Run report ' + [char]0x00B7 + ' ' + $mode) -Subtitle $subtitle -BoxLabel 'Run' -BoxHtml $box -Tiles $tiles `
        -ExtraHtml ('<p class="caption">By status</p><div class="distribution">' + ($bands -join '') + '</div>')))

    # ---- Steps -------------------------------------------------------------------------------------
    [void]$sb.Append('<section class="panel"><h2>Steps</h2><p class="caption">Click a step to open its actions. Steps with a failure are open by default.</p><ul class="toc">')
    foreach ($g in $byStep) {
        $aid = 's-' + ($g.Name -replace '[^a-zA-Z0-9]', '-')
        $fc  = @($g.Group | Where-Object Status -eq 'Failed').Count
        $tone = if ($fc) { 'var(--cp-danger)' } elseif (@($g.Group | Where-Object Status -eq 'Skipped').Count) { 'var(--cp-warning)' } else { 'var(--cp-success)' }
        $info = if ($fc) { '<span class="toc-ko">{0} failed</span>' -f $fc } else { '<span class="muted">{0}</span>' -f $g.Count }
        [void]$sb.Append(('<li><span class="dot" style="--tone:{0}"></span><a class="toc-name" href="#{1}" title="{4}">{2}</a>{3}</li>' -f $tone, $aid, (ConvertTo-SafeHtml (& $stepLabel $g.Name)), $info, (ConvertTo-SafeHtml $g.Name)))
    }
    [void]$sb.Append('</ul></section>')

    # ---- Actions -----------------------------------------------------------------------------------
    [void]$sb.Append('<section class="panel"><h2>Actions</h2><div class="controls">')
    [void]$sb.Append('<input id="q" type="search" placeholder="Search target, action, value..." style="min-width:280px" autocomplete="off">')
    [void]$sb.Append('<select id="st"><option value="">All statuses</option>')
    foreach ($s in $order) { if ($counts[$s]) { [void]$sb.Append(('<option value="{0}">{1} ({2})</option>' -f $s, $labels[$s], $counts[$s])) } }
    [void]$sb.Append('</select><button type="button" onclick="toggleAll(true)">Expand all</button><button type="button" onclick="toggleAll(false)">Collapse all</button><span id="state" class="muted"></span></div>')

    foreach ($g in $byStep) {
        $aid  = 's-' + ($g.Name -replace '[^a-zA-Z0-9]', '-')
        $fc   = @($g.Group | Where-Object Status -eq 'Failed').Count
        $open = if ($fc) { ' open' } else { '' }
        $pills = foreach ($ss in ($g.Group | Group-Object Status | Sort-Object { [array]::IndexOf($order, $_.Name) })) {
            '<span class="pill" style="--tone:{0}">{1} {2}</span>' -f (Get-ExStatusColor $ss.Name), $labels[$ss.Name], $ss.Count
        }
        [void]$sb.Append(('<details{0} id="{1}"><summary><span class="acc-icon">&#9654;</span><span class="acc-step">{2} <span class="muted" style="font-weight:400">{5}</span></span><span class="acc-pills">{3}</span><span class="acc-cnt">{4} action(s)</span></summary>' -f $open, $aid, (ConvertTo-SafeHtml (& $stepLabel $g.Name)), ($pills -join ''), $g.Count, (ConvertTo-SafeHtml $g.Name)))
        [void]$sb.Append('<div class="tbl-wrap"><table><thead><tr><th>Time</th><th>Target</th><th>Action</th><th>Status</th><th>Before</th><th>After</th><th>Detail / error</th></tr></thead><tbody>')
        foreach ($e in $g.Group) {
            $bef = if ($e.BeforeValue) { '<span class="val">' + (ConvertTo-SafeHtml $e.BeforeValue) + '</span>' } else { '<span class="em">&mdash;</span>' }
            $aft = if ($e.AfterValue)  { '<span class="val">' + (ConvertTo-SafeHtml $e.AfterValue) + '</span>' }  else { '<span class="em">&mdash;</span>' }
            $det = if ($e.ErrorMessage) { '<span class="er">' + (ConvertTo-SafeHtml $e.ErrorMessage) + '</span>' }
                   elseif ($e.Detail)   { '<span class="val">' + (ConvertTo-SafeHtml $e.Detail) + '</span>' }
                   else                 { '<span class="em">&mdash;</span>' }
            [void]$sb.Append(('<tr data-status="{0}"><td class="td-ts">{1}</td><td class="td-tgt">{2}</td><td>{3}</td><td><span class="bdg" style="--tone:{4}">{5}</span></td><td>{6}</td><td>{7}</td><td>{8}</td></tr>' -f `
                $e.Status, $e.Timestamp, (ConvertTo-SafeHtml $e.Target), (ConvertTo-SafeHtml $e.Action), (Get-ExStatusColor $e.Status), $labels[$e.Status], $bef, $aft, $det))
        }
        [void]$sb.Append('</tbody></table></div></details>')
    }
    [void]$sb.Append('</section>')

    [void]$sb.Append(@'
<script>
(function () {
  var q = document.getElementById('q'), st = document.getElementById('st'), state = document.getElementById('state');
  function apply() {
    var text = q.value.trim().toLowerCase(), status = st.value, shown = 0, total = 0;
    document.querySelectorAll('details').forEach(function (d) {
      var visible = 0;
      d.querySelectorAll('tbody tr').forEach(function (tr) {
        total++;
        var ok = (!status || tr.getAttribute('data-status') === status) && (!text || tr.textContent.toLowerCase().indexOf(text) >= 0);
        tr.hidden = !ok; if (ok) { visible++; shown++; }
      });
      d.hidden = visible === 0;
      if ((text || status) && visible) { d.open = true; }
    });
    state.textContent = (text || status) ? shown + ' of ' + total + ' actions' : total + ' actions';
  }
  window.toggleAll = function (open) { document.querySelectorAll('details').forEach(function (d) { d.open = open; }); };
  q.addEventListener('input', apply); st.addEventListener('change', apply); apply();
  function goAnchor() { var h = location.hash; if (!h) { return; } var el = document.querySelector(h); if (el && el.tagName === 'DETAILS') { el.open = true; setTimeout(function () { el.scrollIntoView({ behavior: 'smooth', block: 'start' }); }, 80); } }
  goAnchor(); window.addEventListener('hashchange', goAnchor);
})();
</script>
'@)
    [void]$sb.Append((Get-ExHtmlFooter -Text 'This file contains server, database and mailbox names: store and share it accordingly.'))

    $sb.ToString() | Set-Content -Path $Path -Encoding UTF8
    Write-ExLog 'DEBUG' "Run report: $Path"
}

#endregion
