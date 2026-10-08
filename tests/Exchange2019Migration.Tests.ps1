#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Exchange 2013/2016 to 2019 Migration - automated tests (Pester 5 or later).
    Author  : Nicolas Fabert
    Version : 2.0.0

    Run:  Invoke-Pester -Path .\tests -Output Detailed

    No Exchange server and no Active Directory are needed: the tests check the files, the step
    catalogue, the configuration, the protocol-log filter, the console, the reports, the entry
    script (parameter checks only), the versions and the package. The framework itself runs in
    Windows PowerShell 5.1: the tests that matter for it start powershell.exe.
#>

BeforeAll {
    $script:Root = Split-Path $PSScriptRoot -Parent
    $script:PackageRoot = Join-Path $script:Root 'package'
    $script:Manifest = Join-Path $script:PackageRoot 'Modules\Exchange2019.Common.psd1'
    Import-Module $script:Manifest -Force -DisableNameChecking
    $script:Module = Get-Module Exchange2019.Common
    $script:Version = (Import-PowerShellDataFile $script:Manifest).ModuleVersion
    $script:Ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $script:Code = @(Get-ChildItem $script:Root -Recurse -File -Include '*.ps1', '*.psm1', '*.psd1' |
        Where-Object { $_.FullName -notmatch '\\(Reports|\.git)\\' -and $_.Name -ne 'HealthChecker.ps1' })
    $script:Config = Import-PowerShellDataFile (Join-Path $script:PackageRoot 'Configs\Deployment.config.psd1')

    function Get-FunctionFromFile([string]$File, [string]$Name) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($File, [ref]$null, [ref]$null)
        return $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    }
}

Describe 'Files' {
    It 'parse in Windows PowerShell 5.1, the shell of the framework' {
        $list = Join-Path $TestDrive 'files.txt'
        $script:Code.FullName | Set-Content -LiteralPath $list -Encoding UTF8
        $errors = & $script:Ps51 -NoProfile -Command "Get-Content -LiteralPath '$list' | ForEach-Object { `$t = `$null; `$e = `$null; [void][System.Management.Automation.Language.Parser]::ParseFile(`$_, [ref]`$t, [ref]`$e); foreach (`$x in `$e) { `$_ + ':' + `$x.Extent.StartLineNumber + ' ' + `$x.Message } }"
        $errors | Should -BeNullOrEmpty
    }
    It 'are UTF-8 with BOM (Windows PowerShell 5.1 reads a file without BOM as ANSI)' {
        $noBom = foreach ($f in $script:Code) { $b = [IO.File]::ReadAllBytes($f.FullName); if (-not ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)) { $f.Name } }
        $noBom | Should -BeNullOrEmpty
    }
    It 'are plain ASCII: console symbols are built from their code points' {
        $nonAscii = foreach ($f in $script:Code) { if ([IO.File]::ReadAllText($f.FullName) -match '[^\x00-\x7F]') { $f.Name } }
        $nonAscii | Should -BeNullOrEmpty
    }
    It 'contain no French left in the steps and the entry script' {
        $files = @(Get-ChildItem (Join-Path $script:PackageRoot 'Steps') -Filter *.ps1) + (Get-Item (Join-Path $script:PackageRoot 'Deploy-Exchange2019.ps1'))
        $hits = $files | Select-String -Pattern '\b(serveur|etape|aucun|fichier|deja|rapport|boite|sauvegarde|lancer|verifier)\b' -CaseSensitive:$false
        @($hits | ForEach-Object { "$($_.Filename):$($_.LineNumber)" }) | Should -BeNullOrEmpty
    }
}

Describe 'Step catalogue' {
    BeforeAll { $script:Catalog = Get-DeploymentStepCatalog }
    It 'has 26 steps numbered 1 to 26, with unique names' {
        @($script:Catalog.Keys) | Should -Be @(1..26 | ForEach-Object { [string]$_ })
        @($script:Catalog.Values | ForEach-Object { $_.Name } | Sort-Object -Unique).Count | Should -Be 26
    }
    It 'has one file per step, named StepNN-Name.ps1, that defines Invoke-Step' {
        foreach ($k in $script:Catalog.Keys) {
            $e = $script:Catalog[$k]
            $e.File | Should -Be ('Step{0:D2}-{1}.ps1' -f [int]$k, $e.Name)
            $path = Join-Path $script:PackageRoot "Steps\$($e.File)"
            Test-Path $path | Should -BeTrue
            Get-FunctionFromFile $path 'Invoke-Step' | Should -Not -BeNullOrEmpty
        }
        @(Get-ChildItem (Join-Path $script:PackageRoot 'Steps') -Filter 'Step*.ps1').Count | Should -Be 26
    }
    It 'marks steps 18 to 26 ManualOnly, and only them' {
        $manual = @($script:Catalog.Keys | Where-Object { $script:Catalog[$_].ManualOnly })
        $manual | Should -Be @(18..26 | ForEach-Object { [string]$_ })
    }
    It 'uses phases and icons that exist' {
        foreach ($e in $script:Catalog.Values) {
            $e.Phase | Should -BeIn 'Platform', 'HighAvailability', 'Runtime', 'Migration'
            foreach ($style in 'Emoji', 'Symbols', 'Ascii') { (Get-ExIconSet $style)[$e.Icon] | Should -Not -BeNullOrEmpty }
        }
    }
}

Describe 'Configuration' {
    It 'has every top-level key that the steps read' {
        $used = Get-ChildItem (Join-Path $script:PackageRoot 'Steps') -Filter *.ps1 | Select-String -Pattern '\$Config\.([A-Za-z0-9]+)' -AllMatches |
            ForEach-Object { $_.Matches } | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
        $missing = @($used | Where-Object { -not $script:Config.ContainsKey($_) })
        $missing | Should -BeNullOrEmpty
    }
    It 'keeps the strict real-user filters of the protocol-log analysis' {
        $la = $script:Config.LogAnalysis
        $la.IISRequireAuthenticatedUser | Should -BeTrue
        $la.IISRequireSuccessStatus | Should -BeTrue
        $la.SmtpRequireMailFromForValid | Should -BeTrue
        $la.SmtpRequireSenderAddress | Should -BeTrue
        $la.IISExcludeUserAgentPatterns | Should -Contain 'AMProbe'
        foreach ($p in @($la.IISExcludeUserPatterns) + @($la.SmtpExcludeSenderPatterns)) { { [regex]::new($p) } | Should -Not -Throw }
    }
}

Describe 'Protocol-log filter (Step 26)' {
    BeforeAll {
        $fn = Get-FunctionFromFile (Join-Path $script:PackageRoot 'Steps\Step26-AnalyzeProtocolLogs.ps1') 'Test-ProtocolLogExcludedIdentity'
        . ([scriptblock]::Create($fn.Extent.Text))
        $script:IisPatterns = @($script:Config.LogAnalysis.IISExcludeUserPatterns)
        $script:SmtpPatterns = @($script:Config.LogAnalysis.SmtpExcludeSenderPatterns)
    }
    It 'excludes the IIS identity <Identity>' -TestCases @(
        @{ Identity = 'CONTOSO\HealthMailbox0f3a1c2d' }, @{ Identity = 'HealthMailbox0f3a1c2d@contoso.com' }, @{ Identity = 'CONTOSO\SystemMailbox{1f05a927-0c1b}' }
        @{ Identity = 'AMProbe' }, @{ Identity = 'CONTOSO\extest_4a2b9c' }, @{ Identity = 'IIS APPPOOL\MSExchangeOWAAppPool' }, @{ Identity = 'IUSR' }
        @{ Identity = 'NT AUTHORITY\NETWORK SERVICE' }, @{ Identity = 'CONTOSO\DiscoverySearchMailbox{D919BA05}' }, @{ Identity = '-' }, @{ Identity = '' }
    ) { param($Identity) Test-ProtocolLogExcludedIdentity -Identity $Identity -Patterns $script:IisPatterns | Should -BeTrue }
    It 'excludes the SMTP sender <Identity>' -TestCases @(
        @{ Identity = 'HealthMailbox12ab34cd@contoso.com' }, @{ Identity = 'SystemMailbox{bb558c35}@contoso.com' }, @{ Identity = 'FederatedEmail.4c1f4d8b@contoso.com' }
        @{ Identity = 'OABGen@contoso.local' }, @{ Identity = 'postmaster@contoso.com' }, @{ Identity = 'MAILER-DAEMON' }, @{ Identity = '<>' }, @{ Identity = '' }
    ) { param($Identity) Test-ProtocolLogExcludedIdentity -Identity $Identity -Patterns $script:SmtpPatterns | Should -BeTrue }
    It 'keeps the real user <Identity>' -TestCases @(
        @{ Identity = 'CONTOSO\adele.vance' }, @{ Identity = 'adele.vance@contoso.com' }, @{ Identity = 'CONTOSO\health.team' }, @{ Identity = 'migration.project@contoso.com' }
    ) { param($Identity)
        Test-ProtocolLogExcludedIdentity -Identity $Identity -Patterns $script:IisPatterns | Should -BeFalse
        Test-ProtocolLogExcludedIdentity -Identity $Identity -Patterns $script:SmtpPatterns | Should -BeFalse
    }
}

Describe 'Console' {
    It 'uses only characters of the classic console fonts outside the emoji style' {
        # Repertoire of Consolas and Lucida Console: code page 437 and Latin-1. The classic console has
        # no font fallback: any other character is shown as an empty box.
        $cp437 = [Text.Encoding]::GetEncoding(437)
        $safe = New-Object 'System.Collections.Generic.HashSet[int]'
        foreach ($c in (0x20..0x7E) + (0xA0..0xFF)) { [void]$safe.Add($c) }
        foreach ($b in 0x80..0xFE) { $ch = $cp437.GetString([byte[]]@($b)); if ($ch.Length -eq 1) { [void]$safe.Add([int][char]$ch) } }
        # Glyphs of the control range 0x01-0x1F and 0x7F of code page 437 (decoders return control characters).
        foreach ($c in 0x263A, 0x263B, 0x2665, 0x2666, 0x2663, 0x2660, 0x2022, 0x25D8, 0x25CB, 0x25D9, 0x2642, 0x2640, 0x266A, 0x266B, 0x263C, 0x25BA, 0x25C4, 0x2195, 0x203C, 0x00B6, 0x00A7, 0x25AC, 0x21A8, 0x2191, 0x2193, 0x2192, 0x2190, 0x221F, 0x2194, 0x25B2, 0x25BC, 0x2302) { [void]$safe.Add($c) }
        $used = New-Object System.Collections.Generic.List[string]
        foreach ($set in (Get-ExIconSet 'Symbols'), (Get-ExIconSet 'Ascii'), (Get-ExFrameSet 'Symbols' 'Lucida Console'), (Get-ExFrameSet 'Ascii' $null)) {
            foreach ($key in $set.Keys) { $used.Add([string]$set[$key]) }
        }
        # Rounded corners only with the fonts that have them (Consolas).
        $rounded = Get-ExFrameSet 'Symbols' 'Consolas'
        $rounded.TopLeft | Should -Be ([char]0x256D)
        foreach ($key in 'Horizontal', 'Vertical') { $used.Add([string]$rounded[$key]) }
        $bad = @($used | Where-Object { @($_.ToCharArray() | Where-Object { -not $safe.Contains([int]$_) }).Count })
        $bad | Should -BeNullOrEmpty
        $used.Count | Should -BeGreaterThan 40
    }
    It 'has the same icon names in the three styles' {
        $names = (Get-ExIconSet 'Emoji').Keys | Sort-Object
        ((Get-ExIconSet 'Symbols').Keys | Sort-Object) | Should -Be $names
        ((Get-ExIconSet 'Ascii').Keys | Sort-Object) | Should -Be $names
    }
    It 'prints one line per result, with its status, target and action' {
        & $script:Module { $Script:Mode = 'Inventory'; Enable-ExConsoleCapture -Quiet }
        Add-Report -Step '03-VirtualDirectories' -Target 'EXCH201901' -Action 'OWA virtual directory' -Status Inventoried -BeforeValue 'InternalUrl=https://mail.contoso.com/owa' -Detail 'target'
        $text = (Get-ExConsoleCapture | ForEach-Object Text) -join ''
        $text | Should -Match 'INV'
        $text | Should -Match 'EXCH201901'
        $text | Should -Match 'OWA virtual directory'
        & $script:Module { $Script:Capture = $null; $Script:GlobalReportEntries.Clear(); $Script:ReportEntries.Clear() }
    }
}

Describe 'Reports' {
    It 'writes the step CSV and the run report with the shared theme' {
        & $script:Module { param($d) $Script:Mode = 'Apply'; $Script:OutputFolder = $d; $Script:ReportSaved = $false; Enable-ExConsoleCapture -Quiet } $TestDrive
        Add-Report -Step '14-EventLogSize' -Target 'EXCH201901' -Action 'Application log size' -Status Success -Phase 'After' -BeforeValue '20971520' -AfterValue '1073741824'
        Add-Report -Step '14-EventLogSize' -Target 'EXCH201902' -Action 'Application log size' -Status Failed -ErrorMessage 'Access denied'
        $csv = Save-Report -StepName 'Step14-EventLogSize'
        $rows = @(Import-Csv $csv -Delimiter ';')
        $rows.Count | Should -Be 2
        ($rows[0].PSObject.Properties.Name -join ',') | Should -Be 'Timestamp,Step,Phase,Target,Action,Status,BeforeValue,AfterValue,Detail,ErrorMessage,Mode'
        $html = [IO.File]::ReadAllText((Join-Path $TestDrive 'Report_GLOBAL_Apply.html'))
        $html | Should -Match '--cp-accent'
        $html | Should -Match 'Run report'
        $html | Should -Match 'event logs set to 1 GB'
        & $script:Module { $Script:Capture = $null; $Script:GlobalReportEntries.Clear(); Reset-Report }
    }
    It 'builds every HTML report of the steps with the shared theme' {
        $generators = Get-ChildItem (Join-Path $script:PackageRoot 'Steps') -Filter *.ps1 | Where-Object { (Get-Content $_.FullName -Raw) -match '</html>|Get-ExHtmlFooter' }
        @($generators).Count | Should -Be 5
        foreach ($g in $generators) {
            $text = Get-Content $g.FullName -Raw
            $text | Should -Match 'Get-ExHtmlHead'
            $text | Should -Match 'Get-ExHtmlFooter'
            $text | Should -Not -Match '#0f2240|<!DOCTYPE'
        }
    }
}

Describe 'Entry script' {
    BeforeAll {
        $script:Entry = Join-Path $script:PackageRoot 'Deploy-Exchange2019.ps1'
        function Invoke-Entry([string]$Arguments) {
            $out = & $script:Ps51 -NoProfile -Command "& '$script:Entry' $Arguments -OutputFolder '$TestDrive\Reports'; exit `$LASTEXITCODE" 2>&1 | Out-String
            [pscustomobject]@{ Output = $out; Code = $LASTEXITCODE }
        }
    }
    It 'shows the catalogue of the 26 steps' {
        $r = Invoke-Entry '-Step List'
        $r.Output | Should -Match 'Phase 4'
        $r.Output | Should -Match 'AnalyzeProtocolLogs'
    }
    It 'stops with exit code 11 on an unknown step' {
        (Invoke-Entry '-Step 99').Code | Should -Be 11
    }
    It 'stops with exit code 13 when -Follow is given without step 20' {
        (Invoke-Entry '-Step 1 -Follow').Code | Should -Be 13
    }
}

Describe 'Versions' {
    It 'is the same in the manifest, the module, the entry script, the steps, the tools, the changelog and the guide' {
        & $script:Module { $Script:ToolVersion } | Should -Be $script:Version
        $files = @(Get-ChildItem (Join-Path $script:PackageRoot 'Steps') -Filter *.ps1) + @(Get-ChildItem (Join-Path $script:Root 'tools') -Filter *.ps1) +
            @(Get-Item (Join-Path $script:PackageRoot 'Deploy-Exchange2019.ps1'), (Join-Path $script:PackageRoot 'Manage-IISLogs.ps1'), (Join-Path $script:PackageRoot 'Modules\Exchange2019.Common.psm1'), (Join-Path $script:PackageRoot 'Configs\Deployment.config.psd1'), $PSCommandPath)
        $wrong = foreach ($f in $files) {
            $m = [regex]::Match([IO.File]::ReadAllText($f.FullName), '(?m)^\s*#?\s*Version\s*:\s*(\S+)')
            if (-not $m.Success -or $m.Groups[1].Value -ne $script:Version) { "$($f.Name)=$($m.Groups[1].Value)" }
        }
        $wrong | Should -BeNullOrEmpty
        [regex]::Match([IO.File]::ReadAllText((Join-Path $script:Root 'CHANGELOG.md')), '## \[(\d+\.\d+(\.\d+)?)\]').Groups[1].Value | Should -Be $script:Version
        [regex]::Match([IO.File]::ReadAllText((Join-Path $script:PackageRoot 'Docs\Exchange2019Migration-Guide.md')), '(?m)^version:\s*(\S+)').Groups[1].Value | Should -Be $script:Version
    }
}

Describe 'Package' {
    It 'contains the runtime files only, without any value of the environment' {
        $destination = Join-Path $TestDrive 'package'
        & $script:Ps51 -NoProfile -ExecutionPolicy Bypass -File (Join-Path $script:Root 'tools\New-MigrationPackage.ps1') -Destination $destination | Out-Null
        $LASTEXITCODE | Should -Be 0
        Test-Path (Join-Path $destination 'Deploy-Exchange2019.ps1') | Should -BeTrue
        Test-Path (Join-Path $destination 'Docs\Exchange2019Migration-Guide.html') | Should -BeTrue
        foreach ($name in 'Reports', 'tools', 'tests') { Test-Path (Join-Path $destination $name) | Should -BeFalse }
        $config = Import-PowerShellDataFile (Join-Path $destination 'Configs\Deployment.config.psd1')
        $config.LicenseKey | Should -BeNullOrEmpty
        $config.UrlExterne | Should -BeNullOrEmpty
        @($config.PreferredGCs).Count | Should -Be 0
        @($config.KerberosASA.SPNs).Count | Should -Be 0
        # Same settings as the source otherwise.
        @($config.Keys).Count | Should -Be @($script:Config.Keys).Count
        $config.LogAnalysis.IISExcludeUserAgentPatterns | Should -Contain 'AMProbe'
    }
}
