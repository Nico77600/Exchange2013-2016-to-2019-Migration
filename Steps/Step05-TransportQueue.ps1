<#
.SYNOPSIS
    Step 05 - Move the Transport queue database to Q:\Queue.

.DESCRIPTION
    Moves the Transport (queue) database of each Exchange 2019 server to the configured queue root.

    Action: runs Move-TransportDatabase.ps1 (Microsoft script shipped with Exchange).
    Destination: Config.TransportQueueRoot (for example: Q:\Queue).

    The native script restarts the MSExchangeTransport and MSExchangeFrontEndTransport services.
    It must be run server by server (with a PSSession to the Exchange Management Shell of the target).

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

    Write-StepBanner -StepName '05' -Title 'Transport queue move' -Mode $Mode -Actions @(
        ('Target: {0}' -f $Config.TransportQueueRoot),
        'For each Exchange 2019 server:',
        '  - Inventory of the current QueueDatabasePath / QueueDatabaseLoggingPath',
        '  - If already on the target: AlreadyDone',
        '  - Otherwise: run Move-TransportDatabase.ps1 provided by Exchange',
        'The MSExchangeTransport and FrontendTransport services will be restarted by the Microsoft script'
    )

    Initialize-ExchangeShell -Credential $Credential

    $servers = @(Get-Exchange2019Servers)
    $queueRoot = $Config.TransportQueueRoot

    foreach ($srv in $servers) {
        $server = $srv.Name

        Invoke-Action -Step '05-TransportQueue' -Target $server -Action 'Move-TransportDatabase' `
            -Detail "Target $queueRoot" `
            -InventoryScript {
                # Read the remote EdgeTransport.exe.config (standard path)
                $cfg = Invoke-Command -ComputerName $server -ScriptBlock {
                    $exePath = Join-Path $env:ExchangeInstallPath 'Bin\EdgeTransport.exe.config'
                    if (Test-Path $exePath) {
                        [xml]$x = Get-Content $exePath
                        $kv = @{}
                        foreach ($k in 'QueueDatabasePath','QueueDatabaseLoggingPath','IPFilterDatabasePath','IPFilterDatabaseLoggingPath','TemporaryStoragePath') {
                            $node = $x.configuration.appSettings.add | Where-Object { $_.key -eq $k }
                            $kv[$k] = $node.value
                        }
                        return $kv
                    }
                } -ErrorAction Stop
                ($cfg.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
            } `
            -PreCheckScript {
                $cfg = Invoke-Command -ComputerName $server -ScriptBlock {
                    $exePath = Join-Path $env:ExchangeInstallPath 'Bin\EdgeTransport.exe.config'
                    [xml]$x = Get-Content $exePath
                    ($x.configuration.appSettings.add | Where-Object { $_.key -eq 'QueueDatabasePath' }).value
                } -ErrorAction SilentlyContinue
                return ([string]$cfg).StartsWith($queueRoot, [System.StringComparison]::OrdinalIgnoreCase)
            } `
            -ActionScript {
                Invoke-Command -ComputerName $server -ScriptBlock {
                    param($queueRoot)
                    $ErrorActionPreference = 'Stop'

                    # Move-TransportDatabase.ps1 uses Exchange .NET types
                    # (Microsoft.Exchange.Data.LocalLongFullPath, ...) which are NOT
                    # available in a standard PSRemoting session. The Exchange snap-in
                    # must be loaded in this session before invoking the script.
                    if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -EA SilentlyContinue)) {
                        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
                    }

                    if (-not (Test-Path $queueRoot)) {
                        New-Item -ItemType Directory -Path $queueRoot -Force | Out-Null
                    }
                    $scripts = Join-Path $env:ExchangeInstallPath 'Scripts'
                    Push-Location $scripts
                    try {
                        & .\Move-TransportDatabase.ps1 `
                            -QueueDatabasePath          (Join-Path $queueRoot 'QueueDB') `
                            -QueueDatabaseLoggingPath   (Join-Path $queueRoot 'QueueDB') `
                            -IPFilterDatabasePath       (Join-Path $queueRoot 'IPFilter') `
                            -IPFilterDatabaseLoggingPath(Join-Path $queueRoot 'IPFilter') `
                            -TemporaryStoragePath       (Join-Path $queueRoot 'Temp')
                    } finally {
                        Pop-Location
                    }
                } -ArgumentList $queueRoot -ErrorAction Stop
            }
    }

    return 0
}
