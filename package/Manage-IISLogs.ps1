<#
.SYNOPSIS
    Exchange IIS log management: weekly compression + purge after 6 months.

.DESCRIPTION
    This script replaces the former Purge_IIS_Logs.ps1 (hard deletion after
    30 days) with a two-stage approach:

      1) Weekly compression:
         - IIS logs older than -RetainUncompressedDays (7 days by default)
           are archived into .zip files, grouped by month.
           E.g.: D:\inetpub\logs\LogFiles\W3SVC1\ -> u_ex_2024-09.zip
         - The original .log files are deleted after successful compression.

      2) Purge:
         - .zip files older than -PurgeAfterDays (180 days by default) are deleted.

    The -Action Install mode registers a weekly scheduled task that
    invokes this same script with -Action Compress then Purge on each server
    provided through -Servers or -ServersCsv.

.PARAMETER Action
    Compress: compresses the .log files older than RetainUncompressedDays days.
    Purge: deletes the .zip archives older than PurgeAfterDays days.
    Install: deploys a weekly scheduled task on the target servers.

.PARAMETER LogPath
    Root of the IIS logs. Default: C:\inetpub\logs\LogFiles
    The script descends recursively (W3SVC1, W3SVC2, ...).

.PARAMETER RetainUncompressedDays
    Number of days during which the uncompressed .log files are kept. Default 7.

.PARAMETER PurgeAfterDays
    Age of the archives to purge. Default 180 (6 months).

.PARAMETER Servers
    List of servers on which to deploy the task (Install action).
    E.g.: 'srv01.contoso.local','srv02.contoso.local'

.PARAMETER ServersCsv
    Alternative to -Servers: a CSV file with a Fqdn or Name column.

.PARAMETER ScheduledTaskName
    Name of the scheduled task. Default: 'Manage IIS Logs - Exchange'

.PARAMETER ScriptDeployPath
    Local path on the target servers where the script is copied.
    Default: C:\Scripts\Manage-IISLogs.ps1

.PARAMETER RunDay
    Weekly execution day of the task. Default: Sunday.

.PARAMETER RunTime
    Execution time hh:mm. Default: 02:00.

.PARAMETER ServiceAccount
    Account under which the task runs (UPN or DOMAIN\user).
    If omitted: SYSTEM.

.EXAMPLE
    # Local test: compression of everything older than 7 days
    .\Manage-IISLogs.ps1 -Action Compress

.EXAMPLE
    # Purge of the archives older than 6 months
    .\Manage-IISLogs.ps1 -Action Purge

.EXAMPLE
    # Deployment of the scheduled task on 4 servers (from an admin workstation)
    .\Manage-IISLogs.ps1 -Action Install `
        -Servers 'EXCH201901.contoso.local','EXCH201902.contoso.local',
                 'EXCH201903.contoso.local','EXCH201904.contoso.local' `
        -ServiceAccount 'CONTOSO\svc-iislogs'

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    Part of : Exchange 2013/2016 to 2019 Migration - Deploy-Exchange2019.ps1

    No external dependency: Compress-Archive (PS5+), Get-ChildItem, Remove-Item.
    The script writes its own log to $LogPath\ManageIISLogs_<date>.txt
#>

[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Compress','Purge','Install')]
    [string]$Action,

    [string]$LogPath = 'C:\inetpub\logs\LogFiles',

    [int]$RetainUncompressedDays = 7,

    [int]$PurgeAfterDays = 180,

    # ---- Install mode only ----
    [Parameter(ParameterSetName = 'Install')]
    [string[]]$Servers,

    [Parameter(ParameterSetName = 'Install')]
    [string]$ServersCsv,

    [Parameter(ParameterSetName = 'Install')]
    [string]$ScheduledTaskName = 'Manage IIS Logs - Exchange',

    [Parameter(ParameterSetName = 'Install')]
    [string]$ScriptDeployPath = 'C:\Scripts\Manage-IISLogs.ps1',

    [Parameter(ParameterSetName = 'Install')]
    [ValidateSet('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday')]
    [string]$RunDay = 'Sunday',

    [Parameter(ParameterSetName = 'Install')]
    [string]$RunTime = '02:00',

    [Parameter(ParameterSetName = 'Install')]
    [string]$ServiceAccount,

    [Parameter(ParameterSetName = 'Install')]
    [System.Management.Automation.PSCredential]$Credential
)

$ErrorActionPreference = 'Stop'

#region -- Internal logging ------------------------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error')] [string]$Level = 'Info'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) { 'Success' { 'Green' } 'Warning' { 'Yellow' } 'Error' { 'Red' } default { 'Gray' } }
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $color

    # Persist to disk (locally only, not in Install)
    if ($Action -ne 'Install' -and (Test-Path $LogPath -ErrorAction SilentlyContinue)) {
        $logFile = Join-Path $LogPath ("ManageIISLogs_{0}.txt" -f (Get-Date -Format 'yyyyMM'))
        Add-Content -Path $logFile -Value "[$ts] [$Level] $Message" -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}
#endregion

#region -- Compression -----------------------------------------------------------
function Invoke-CompressLogs {
    param(
        [string]$Root,
        [int]$KeepDays
    )
    if (-not (Test-Path $Root)) {
        Write-Log "Path not found: $Root" -Level Warning
        return
    }

    $cutoff = (Get-Date).Date.AddDays(-$KeepDays)
    Write-Log "Compression: .log < $cutoff, root=$Root"

    # For each W3SVCx subfolder
    $sites = Get-ChildItem -Path $Root -Directory -ErrorAction SilentlyContinue
    if (-not $sites) {
        Write-Log "No IIS site subfolder under $Root" -Level Warning
        return
    }

    $totalCompressed = 0
    $totalDeleted    = 0

    foreach ($site in $sites) {
        # -Recurse: IIS typically creates the logs in W3SVC<n>\ subfolders
        # (one per IIS site). The .log files are searched at any depth under each
        # first-level folder of $Root.
        $logs = Get-ChildItem -Path $site.FullName -Filter '*.log' -File -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $cutoff }
        if (-not $logs) { continue }

        Write-Log "Site $($site.Name): $($logs.Count) log files to compress"

        # Grouping by (parent subfolder + week of the month).
        # One compression per week -> name "weekN_MMYYYY" is more telling than
        # the whole month (e.g. week1_052026 = May 1-7, 2026, week2_052026 = May 8-14...).
        # The subfolder is kept in the name to avoid collisions
        # between W3SVC1 and W3SVC2 (same .log names are possible).
        $groups = $logs | Group-Object -Property {
            $d = $_.LastWriteTime
            $weekOfMonth = [math]::Ceiling($d.Day / 7.0)
            $weekKey = 'week{0}_{1:D2}{2}' -f $weekOfMonth, $d.Month, $d.Year
            if ($_.Directory.FullName -eq $site.FullName) {
                # File directly at the site root (simple case) -> no prefix
                $weekKey
            } else {
                # File in a subfolder (W3SVC1, W3SVC2, etc.)
                "$($_.Directory.Name)_$weekKey"
            }
        }

        foreach ($g in $groups) {
            $zipName = "$($site.Name)_$($g.Name).zip"
            $zipPath = Join-Path $site.FullName $zipName

            try {
                if (Test-Path $zipPath) {
                    # Append to it (Update)
                    Compress-Archive -Path $g.Group.FullName -DestinationPath $zipPath -Update -ErrorAction Stop
                } else {
                    Compress-Archive -Path $g.Group.FullName -DestinationPath $zipPath -ErrorAction Stop
                }

                # Verification: if the archive is valid, the .log files can be deleted
                $zipSize = (Get-Item $zipPath).Length
                if ($zipSize -gt 0) {
                    foreach ($f in $g.Group) {
                        Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue
                        $totalDeleted++
                    }
                    $totalCompressed++
                    Write-Log "  -> Archive OK: $zipName ($([math]::Round($zipSize/1MB,2)) MB)" -Level Success
                } else {
                    Write-Log "  -> Empty archive: $zipName, .log files kept" -Level Warning
                }
            } catch {
                Write-Log "  -> Compression error on $zipName - $($_.Exception.Message)" -Level Error
            }
        }
    }

    Write-Log "Compression completed: $totalCompressed archives produced, $totalDeleted .log deleted" -Level Success
}
#endregion

#region -- Purge -----------------------------------------------------------------
function Invoke-PurgeArchives {
    param(
        [string]$Root,
        [int]$AgeDays
    )
    if (-not (Test-Path $Root)) {
        Write-Log "Path not found: $Root" -Level Warning
        return
    }

    $cutoff = (Get-Date).Date.AddDays(-$AgeDays)
    Write-Log "Purge: .zip < $cutoff, root=$Root"

    $archives = Get-ChildItem -Path $Root -Filter '*.zip' -File -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $cutoff }

    if (-not $archives) {
        Write-Log "No archive to purge." -Level Info
        return
    }

    $deletedCount = 0
    $deletedBytes = 0
    foreach ($a in $archives) {
        try {
            $size = $a.Length
            Remove-Item -Path $a.FullName -Force -ErrorAction Stop
            $deletedCount++
            $deletedBytes += $size
            Write-Log "  -> Deleted: $($a.FullName) ($([math]::Round($size/1MB,2)) MB)"
        } catch {
            Write-Log "  -> Delete error on $($a.Name): $($_.Exception.Message)" -Level Error
        }
    }

    Write-Log "Purge completed: $deletedCount archives deleted ($([math]::Round($deletedBytes/1GB,2)) GB freed)" -Level Success
}
#endregion

#region -- Install (scheduled task deployment) ----------------------------------
function Install-ScheduledTaskOnServers {
    param(
        [string[]]$TargetServers,
        [string]$TaskName,
        [string]$DeployScriptPath,
        [string]$Day,
        [string]$Time,
        [string]$Account,
        [System.Management.Automation.PSCredential]$Cred,
        [string]$LocalLogPath,
        [int]$KeepDays,
        [int]$AgeDays
    )

    $scriptSource = $PSCommandPath
    if (-not $scriptSource) {
        $scriptSource = $MyInvocation.MyCommand.Definition
    }
    if (-not (Test-Path $scriptSource)) {
        throw "Source script not found: $scriptSource"
    }

    foreach ($srv in $TargetServers) {
        Write-Log "===== $srv =====" -Level Info

        try {
            # 1. Copy the script
            $sessionParams = @{ ComputerName = $srv ; ErrorAction = 'Stop' }
            if ($Cred) { $sessionParams['Credential'] = $Cred }

            $session = New-PSSession @sessionParams
            try {
                $deployDir = Split-Path $DeployScriptPath -Parent
                Invoke-Command -Session $session -ScriptBlock {
                    param($dir)
                    if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
                } -ArgumentList $deployDir | Out-Null

                Copy-Item -Path $scriptSource -Destination $DeployScriptPath -ToSession $session -Force
                Write-Log "Script copied to $srv at $DeployScriptPath" -Level Success

                # 2. Creation of the scheduled task
                Invoke-Command -Session $session -ScriptBlock {
                    param($taskName, $scriptPath, $day, $time, $account, $logPath, $keep, $age)

                    $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" " +
                           "-Action Compress -LogPath `"$logPath`" -RetainUncompressedDays $keep -PurgeAfterDays $age; " +
                           "& `"$scriptPath`" -Action Purge -LogPath `"$logPath`" -RetainUncompressedDays $keep -PurgeAfterDays $age"

                    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
                    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $day -At $time
                    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 4)

                    if ($account) {
                        $principal = New-ScheduledTaskPrincipal -UserId $account -LogonType Password -RunLevel Highest
                    } else {
                        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
                    }

                    # If the task already exists, it is updated
                    $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
                    if ($existing) {
                        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                    }

                    Register-ScheduledTask -TaskName $taskName `
                        -Action $action -Trigger $trigger -Settings $settings -Principal $principal `
                        -Description "Weekly compression + 6-month purge of the Exchange IIS logs." | Out-Null
                } -ArgumentList $TaskName, $DeployScriptPath, $Day, $Time, $Account, $LocalLogPath, $KeepDays, $AgeDays | Out-Null

                Write-Log "Scheduled task '$TaskName' created on $srv ($Day $Time)" -Level Success

            } finally {
                Remove-PSSession -Session $session -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Log "Error on $srv - $($_.Exception.Message)" -Level Error
        }
    }
}
#endregion

# ===============================================================================
# Main dispatch
# ===============================================================================
switch ($Action) {

    'Compress' {
        Invoke-CompressLogs -Root $LogPath -KeepDays $RetainUncompressedDays
    }

    'Purge' {
        Invoke-PurgeArchives -Root $LogPath -AgeDays $PurgeAfterDays
    }

    'Install' {
        $list = @()
        if ($Servers) { $list += $Servers }
        if ($ServersCsv) {
            if (-not (Test-Path $ServersCsv)) { throw "CSV not found: $ServersCsv" }
            $rows = Import-Csv -Path $ServersCsv
            foreach ($r in $rows) {
                if ($r.Fqdn)      { $list += $r.Fqdn }
                elseif ($r.Name)  { $list += $r.Name }
                elseif ($r.Server){ $list += $r.Server }
            }
        }
        $list = $list | Sort-Object -Unique
        if (-not $list) { throw 'No server provided (-Servers or -ServersCsv).' }

        Write-Log "Deploying the scheduled task to $($list.Count) server(s): $($list -join ', ')" -Level Info
        Install-ScheduledTaskOnServers `
            -TargetServers     $list `
            -TaskName          $ScheduledTaskName `
            -DeployScriptPath  $ScriptDeployPath `
            -Day               $RunDay `
            -Time              $RunTime `
            -Account           $ServiceAccount `
            -Cred              $Credential `
            -LocalLogPath      $LogPath `
            -KeepDays          $RetainUncompressedDays `
            -AgeDays           $PurgeAfterDays
    }
}
