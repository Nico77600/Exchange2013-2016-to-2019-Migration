#Requires -Version 5.1
<#
.SYNOPSIS
    Copies the files needed to run the framework into a separate folder, ready to be zipped and
    delivered, with every environment value removed.

.DESCRIPTION
    The package contains only what Deploy-Exchange2019.ps1 needs at run time, plus the HTML guide:
        Deploy-Exchange2019.ps1, Manage-IISLogs.ps1, Modules\, Steps\, Configs\,
        Docs\Exchange2019Migration-Guide.html, CHANGELOG.md, LICENSE
    It never copies Reports\ (run output: server, database and mailbox names), tools\ or tests\.

    Environment values:
      - Configs\Deployment.config.psd1 is copied with the environment values emptied: URLs, NetBIOS
        domain, product key, source servers, Kerberos ASA account, domain and SPNs, preferred
        domain controllers. The administrator fills them in (guide, chapter 6).
      - Configs\DiskLayout.csv and Configs\DAGInfo.csv are replaced by templates with example
        values (contoso.local, EXCH201901..04).
    The script then checks that none of the removed values - nor the server, domain and DAG names
    of the original CSV files - appears anywhere in the package.

.PARAMETER Destination
    Package folder. Default: package\Deploy-Exchange2019-<version>, next to the tool folder.

.PARAMETER Force
    Replaces the destination folder when it already contains a package. A folder that contains a
    Reports\ sub-folder (a package that has been run) is never replaced.

.EXAMPLE
    .\tools\New-MigrationPackage.ps1
    Creates ..\package\Deploy-Exchange2019-2.0.0.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    Part of : Exchange 2013/2016 to 2019 Migration (repository tool, not in the package)
#>
[CmdletBinding()]
param(
    [string]$Destination,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$version = (Import-PowerShellDataFile (Join-Path $root 'Modules\Exchange2019.Common.psd1')).ModuleVersion
if (-not $Destination) { $Destination = Join-Path (Split-Path $root -Parent) "package\Deploy-Exchange2019-$version" }
$Destination = [IO.Path]::GetFullPath($Destination).TrimEnd('\')

$rootPrefix = [IO.Path]::GetFullPath($root).TrimEnd('\') + '\'
if (($Destination + '\').StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or $rootPrefix.StartsWith($Destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "The destination must be outside the tool folder: $Destination"
}
if (Test-Path -LiteralPath $Destination) {
    if (-not $Force) { throw "The destination already exists: $Destination. Use -Force to replace it." }
    if (-not (Test-Path -LiteralPath (Join-Path $Destination 'Deploy-Exchange2019.ps1'))) { throw "The destination is not a package of the framework, it is not replaced: $Destination" }
    if (Test-Path -LiteralPath (Join-Path $Destination 'Reports')) { throw "The destination contains a Reports folder (run output), it is not replaced: $Destination" }
    Remove-Item -LiteralPath $Destination -Recurse -Force
}

# ---- Files needed at run time ---------------------------------------------------------------------
$files = New-Object System.Collections.Generic.List[string]
foreach ($f in 'Deploy-Exchange2019.ps1', 'Manage-IISLogs.ps1', 'CHANGELOG.md', 'LICENSE', 'Configs\HealthChecker.ps1', 'Docs\Exchange2019Migration-Guide.html') { $files.Add($f) }
foreach ($folder in 'Modules', 'Steps') {
    Get-ChildItem -LiteralPath (Join-Path $root $folder) -File | ForEach-Object { $files.Add($_.FullName.Substring($rootPrefix.Length)) }
}
foreach ($f in $files) {
    $source = Join-Path $root $f
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing file in the tool folder: $f" }
    $target = Join-Path $Destination $f
    [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
    Copy-Item -LiteralPath $source -Destination $target
}

# ---- Values of the environment, collected before they are removed ----------------------------------
$forbidden = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
function Add-Forbidden([string]$Value) {
    if (-not $Value) { return }
    $v = $Value.Trim()
    if ($v.Length -lt 4) { return }
    [void]$forbidden.Add($v)
    # Host and domain labels of a name (webmail.fabrikam.com -> webmail, fabrikam), except generic ones.
    foreach ($label in ($v -split '[./\\@]')) {
        if ($label.Length -ge 5 -and $label -notmatch '^(local|http|https|mail|autodiscover|com|net|org|live|exchange|configs?)$') { [void]$forbidden.Add($label) }
    }
}

$utf8Bom = New-Object Text.UTF8Encoding($true)
$configRelative = 'Configs\Deployment.config.psd1'
$config = [IO.File]::ReadAllText((Join-Path $root $configRelative))
$scalarKeys = 'UrlExterne', 'UrlInterne', 'DomainNetBios', 'LicenseKey', 'SourceServer2013ForConnectors', 'SourceServer2016ForConnectors', 'SourceServerForVDirs', 'AccountName', 'Domain', 'OUPath'
$emptied = 0
foreach ($key in $scalarKeys) {
    $pattern = "(?m)^(\s*$key\s*=\s*)'([^']*)'"
    $found = [regex]::Matches($config, $pattern)
    if ($found.Count -eq 0) { throw "The key $key is missing in $configRelative." }
    foreach ($m in $found) { if ($m.Groups[2].Value) { Add-Forbidden $m.Groups[2].Value; $emptied++ } }
    $config = [regex]::Replace($config, $pattern, '$1''''')
}
foreach ($key in 'SPNs', 'PreferredGCs') {
    $pattern = "(?ms)^(\s*$key\s*=\s*)@\((.*?)\)"
    $found = [regex]::Matches($config, $pattern)
    if ($found.Count -ne 1) { throw "The list $key must appear exactly once in $configRelative (found $($found.Count))." }
    foreach ($item in [regex]::Matches($found[0].Groups[2].Value, "'([^']*)'")) { Add-Forbidden $item.Groups[1].Value; $emptied++ }
    $config = [regex]::Replace($config, $pattern, '$1@()')
}
$configTarget = Join-Path $Destination $configRelative
[void][IO.Directory]::CreateDirectory((Split-Path $configTarget -Parent))
[IO.File]::WriteAllText($configTarget, $config, $utf8Bom)

# CSV files of the environment: their names are forbidden, the package gets templates.
$generic = 'local', 'contoso', 'Queue', 'Databases', 'Swap', 'True', 'False', 'BestAvailability', 'GoodAvailability', 'Lossless',
    'ServerName', 'Role', 'DriveLetter', 'DiskNumber', 'MinSizeGB', 'MaxSizeGB', 'DAGName', 'DAGIps', 'WitnessServer', 'WitnessDir',
    'AltWitnessServer', 'AltWitnessDir', 'Site1Servers', 'Site2Servers', 'MaximumActiveDatabasesSite1', 'MaximumPreferredActiveDatabasesSite1',
    'MaximumActiveDatabasesSite2', 'MaximumPreferredActiveDatabasesSite2', 'ManualDagnetworkConfiguration', 'ReplayLagManagerEnabled',
    'ReplicationPort', 'AutoDatabaseMountDial'
foreach ($csv in 'DiskLayout.csv', 'DAGInfo.csv') {
    $source = Join-Path $root "Configs\$csv"
    if (-not (Test-Path -LiteralPath $source)) { continue }
    foreach ($token in [regex]::Matches([IO.File]::ReadAllText($source), '[A-Za-z][A-Za-z0-9_-]{3,}')) {
        if ($token.Value -notin $generic -and $token.Value -notmatch '^(EXCH2019\d\d|DAG01|DC0\d|FSW)$') { [void]$forbidden.Add($token.Value) }
    }
}$templates = @{
    'DiskLayout.csv' = @(
        'ServerName;Role;DriveLetter;DiskNumber;MinSizeGB;MaxSizeGB'
        foreach ($s in 'EXCH201901', 'EXCH201902', 'EXCH201903', 'EXCH201904') { "$s;Queue;Q;3;90;200"; "$s;Databases;M;4;200;99999"; "$s;Swap;X;;20;90" }
    )
    'DAGInfo.csv' = @(
        'DAGName,DAGIps,WitnessServer,WitnessDir,AltWitnessServer,AltWitnessDir,Site1Servers,Site2Servers,GC,MaximumActiveDatabasesSite1,MaximumPreferredActiveDatabasesSite1,MaximumActiveDatabasesSite2,MaximumPreferredActiveDatabasesSite2,ManualDagnetworkConfiguration,ReplayLagManagerEnabled,ReplicationPort,AutoDatabaseMountDial'
        '"DAG01","255.255.255.255","DC01.contoso.local","C:\DAG01-FSW","","","EXCH201901,EXCH201902","EXCH201903,EXCH201904","DC02.contoso.local","2","","2","0","False","True","64327","BestAvailability"'
    )
}
foreach ($csv in $templates.Keys) { [IO.File]::WriteAllText((Join-Path $Destination "Configs\$csv"), (($templates[$csv] -join "`r`n") + "`r`n"), $utf8Bom) }
# Names of the examples are allowed.
foreach ($allowed in @($forbidden) | Where-Object { $_ -match '^(contoso|EXCH2019\d\d|DAG01|DC0\d)' }) { [void]$forbidden.Remove($allowed) }

# ---- Checks -----------------------------------------------------------------------------------------
$problems = New-Object System.Collections.Generic.List[string]
foreach ($name in 'Reports', 'tools', 'tests') {
    if (Test-Path -LiteralPath (Join-Path $Destination $name)) { $problems.Add("Folder $name\ must not be in the package.") }
}
# Where-Object, not -Include: with -LiteralPath, Windows PowerShell 5.1 ignores -Include.
Get-ChildItem -LiteralPath $Destination -Recurse -File | Where-Object { $_.Extension -in '.pfx', '.p12', '.log' -or $_.Name -eq 'DeploymentState.json' } | ForEach-Object { $problems.Add("Run or secret file in the package: $($_.Name)") }
$textFiles = Get-ChildItem -LiteralPath $Destination -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1', '.psd1', '.csv', '.html', '.md' -and $_.Name -ne 'HealthChecker.ps1' }
foreach ($file in $textFiles) {
    # Images embedded in the HTML guide are base64: random letters, not names.
    $content = [regex]::Replace([IO.File]::ReadAllText($file.FullName), 'data:image/[a-z]+;base64,[A-Za-z0-9+/=]+', '')
    foreach ($value in $forbidden) {
        if ($content.IndexOf($value, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $problems.Add("An environment value ('$value') appears in $($file.FullName.Substring($Destination.Length + 1)).")
        }
    }
}
if ($problems.Count) { throw ("Package not valid ($Destination):`n - " + ($problems -join "`n - ")) }

$all = Get-ChildItem -LiteralPath $Destination -Recurse -File
Write-Host ''
Write-Host "  Exchange 2013/2016 to 2019 Migration $version - package ready" -ForegroundColor Green
Write-Host "  Folder   : $Destination"
Write-Host ('  Content  : {0} files, {1:N1} MB' -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
Write-Host "  Config   : $emptied environment value(s) emptied - fill in Configs\Deployment.config.psd1 (guide, chapter 6)"
Write-Host "  CSV      : DiskLayout.csv and DAGInfo.csv replaced by templates (contoso examples)"
Write-Host "  Checked  : $($forbidden.Count) environment value(s) absent from the package"
Write-Host ''
$all | Sort-Object FullName | ForEach-Object { '    {0,12:N0}  {1}' -f $_.Length, $_.FullName.Substring($Destination.Length + 1) }
