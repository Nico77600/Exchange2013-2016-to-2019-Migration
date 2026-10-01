<#
.SYNOPSIS
    Step 17 - Deployment of the Exchange IIS log management.

.DESCRIPTION
    Deploys on each Exchange 2019 server:
      1. The Manage-IISLogs.ps1 script (weekly compression + purge after 6 months)
      2. A Windows scheduled task that calls this script every Sunday at 02:00
         (day and time configurable in IISLogsManagement of the configuration file)

    Modes:
      - Inventory: checks on each server whether the task and the script exist
      - Simulate: shows what would be deployed, without copying or creating anything
      - Apply: interactively offers an execution account (SYSTEM by default),
               copies the script and registers the scheduled task

    Execution account (Apply only):
      - [Enter/N] SYSTEM: built-in account, no credential required
      - [Y] > [1] Service: DOMAIN\account + password
      - [Y] > [2] gMSA: DOMAIN\account$ + no password (managed by AD)

    Idempotent: task already present AND script already deployed => AlreadyDone.
    The existing task is replaced when the step is run again (update possible).

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

    $iisConf     = $Config.IISLogsManagement
    $taskName    = $iisConf.ScheduledTaskName
    $deployPath  = $iisConf.ScriptDeployPath
    $runDay      = $iisConf.RunDay
    $runTime     = $iisConf.RunTime
    $keepDays    = [int]$iisConf.RetainUncompressedDays
    $purgeDays   = [int]$iisConf.PurgeAfterDays
    $iisLogRoot  = $iisConf.LogPath

    # Source path: Manage-IISLogs.ps1 is at the project root (one level above Steps\)
    $scriptSource = Join-Path (Split-Path $PSScriptRoot -Parent) 'Manage-IISLogs.ps1'

    Write-StepBanner -StepName '17' -Title 'Deployment of the Exchange IIS log management' -Mode $Mode -Actions @(
        "Scheduled task: '$taskName' ($runDay at $runTime)",
        "Deployed script: $deployPath",
        "IIS logs (root): runtime detection via WebAdministration (config fallback: $iisLogRoot)",
        "Retention: $keepDays days (uncompressed) / $purgeDays days (.zip archives)",
        'Account: SYSTEM by default - custom account offered in Apply mode'
    )

    # =========================================================================
    # Helper: runtime detection of the common root of the IIS logs on a server
    # Step06 moves the IIS logs to a custom path (D:\Logs\IIS\<Site>).
    # IIS is queried through WebAdministration rather than trusting the
    # config value, which may be out of sync.
    # =========================================================================
    function Get-RemoteIISLogRoot {
        param(
            [Parameter(Mandatory)] [string]$ServerFqdn,
            [Parameter(Mandatory)] [string]$Fallback
        )
        try {
            $paths = Invoke-Command -ComputerName $ServerFqdn -ScriptBlock {
                Import-Module WebAdministration -ErrorAction Stop
                Get-WebSite | Where-Object { $_.Name -in 'Default Web Site','Exchange Back End' } |
                    ForEach-Object { [System.Environment]::ExpandEnvironmentVariables($_.LogFile.directory) }
            } -ErrorAction Stop
            $paths = @($paths | Where-Object { $_ } | Select-Object -Unique)

            if ($paths.Count -eq 0) { return @{ Path = $Fallback; Source = 'fallback (no IIS site detected)' } }
            if ($paths.Count -eq 1) { return @{ Path = $paths[0];  Source = 'runtime detection (1 IIS site)' } }

            # Computation of the common parent (preserves the drive root).
            # Note: an array of arrays is built. ForEach-Object flattens it, so
            # foreach + ',($...)' is used to preserve the nested structure.
            $partsAll = @()
            foreach ($p in $paths) {
                $partsAll += ,(($p.TrimEnd('\','/')) -split '[/\\]')
            }
            $minDepth = ($partsAll | ForEach-Object { $_.Count } | Measure-Object -Minimum).Minimum
            $common = @()
            for ($i = 0; $i -lt $minDepth; $i++) {
                $atI = $partsAll | ForEach-Object { $_[$i] }
                $uniq = @($atI | Select-Object -Unique)
                if ($uniq.Count -eq 1) { $common += $uniq[0] } else { break }
            }
            # At least drive + 1 segment to avoid returning 'C:\' (too broad)
            if ($common.Count -ge 2) {
                return @{ Path = ($common -join '\'); Source = "common parent ($($paths.Count) IIS sites)" }
            }
            return @{ Path = $Fallback; Source = "fallback (diverging IIS paths: $($paths -join ' | '))" }
        } catch {
            return @{ Path = $Fallback; Source = "fallback (detection error: $($_.Exception.Message))" }
        }
    }

    # Helper: extract -LogPath "..." (or '...') from the argument string of a task
    # Historical format (-File): -LogPath "D:\Logs\IIS"
    # Current format (-Command): -LogPath 'D:\Logs\IIS'
    function Get-LogPathFromTaskArgument {
        param([string]$Argument)
        if (-not $Argument) { return $null }
        $m = [regex]::Match($Argument, "-LogPath\s+['""]([^'""]+)['""]")
        if ($m.Success) { return $m.Groups[1].Value }
        return $null
    }

    Initialize-ExchangeShell -Credential $Credential

    $servers = @(Get-Exchange2019Servers)
    if (-not $servers) {
        Write-Log 'No Exchange 2019 server detected.' -Level Warning
        Save-Report -StepName 'Step17-ManageIISLogs' | Out-Null
        return 0
    }

    if (-not (Test-Path $scriptSource)) {
        Write-Log "Source script not found: $scriptSource" -Level Error
        Save-Report -StepName 'Step17-ManageIISLogs' | Out-Null
        return 1
    }

    # =========================================================================
    # Interactive choice of the execution account (Apply only)
    # =========================================================================
    $taskAccount  = $null   # $null = SYSTEM
    $taskPlainPwd = $null   # $null = SYSTEM or gMSA

    if ($Mode -eq 'Apply') {
        Write-Host ''
        Write-Log '=== Scheduled task execution account ===' -Level Step
        Write-Log '    Default: SYSTEM (no credential required)' -Level Info
        Write-Host ''
        $useCustom = Read-Host '    Specify a custom account? [Y/N] (Enter = N)'

        if ($useCustom -ieq 'Y') {
            Write-Host ''
            Write-Log '    [1] Service account  (password required)' -Level Info
            Write-Log '    [2] gMSA account     (no password, managed by AD)' -Level Info
            $acctType = Read-Host '    Choice [1/2]'

            $login = Read-Host '    Account name (e.g. DOMAIN\svc-iislogs  or  DOMAIN\gMSA-iislogs$)'

            if ($acctType -eq '2') {
                $taskAccount  = $login
                $taskPlainPwd = $null
                Write-Log "    gMSA account selected: $login" -Level Info
            } else {
                $secPwd = Read-Host '    Password' -AsSecureString
                $bstr   = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secPwd)
                $taskPlainPwd = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
                $taskAccount = $login
                Write-Log "    Service account selected: $login" -Level Info
            }
        } else {
            Write-Log '    SYSTEM selected.' -Level Info
        }
        Write-Host ''
    }

    # =========================================================================
    # Loop over the Exchange 2019 servers
    # =========================================================================
    foreach ($srv in $servers) {
        $srvName = $srv.Name
        $srvFqdn = if ($srv.Fqdn) { $srv.Fqdn } else { $srvName }

        $accountLabel = if ($taskAccount) { $taskAccount } else { 'SYSTEM' }

        # Runtime detection of the IIS log root (per server, may differ if the
        # IIS config was customized on a particular server).
        $resolved      = Get-RemoteIISLogRoot -ServerFqdn $srvFqdn -Fallback $iisLogRoot
        $resolvedLogRoot = $resolved.Path
        $resolvedSource  = $resolved.Source
        Write-Log "  $srvName - detected IIS LogPath = '$resolvedLogRoot' ($resolvedSource)" -Level Sub

        Invoke-Action -Step '17-IISLogs' -Target $srvName -Action 'Install-IISLogsTask' `
            -Detail "Task='$taskName' | Script=$deployPath | Account=$accountLabel | LogPath=$resolvedLogRoot ($resolvedSource)" `
            -InventoryScript {
                try {
                    $res = Invoke-Command -ComputerName $srvFqdn -ScriptBlock {
                        param($tName, $sPath)
                        $task = Get-ScheduledTask -TaskName $tName -ErrorAction SilentlyContinue
                        $hasScr = Test-Path $sPath
                        $taskArg = if ($task) {
                            ($task.Actions | Where-Object { $_.Execute -match 'powershell' } |
                                Select-Object -First 1).Arguments
                        } else { '' }
                        [PSCustomObject]@{
                            TaskState   = if ($task) { [string]$task.State } else { 'NotFound' }
                            ScriptFound = $hasScr
                            TaskArg     = $taskArg
                        }
                    } -ArgumentList $taskName, $deployPath -ErrorAction Stop
                    $taskLogPath = Get-LogPathFromTaskArgument -Argument $res.TaskArg
                    "Task=$($res.TaskState) | Script=$($res.ScriptFound) | LogPathInTask=$taskLogPath | LogPathExpected=$resolvedLogRoot"
                } catch {
                    "Unreachable: $($_.Exception.Message)"
                }
            } `
            -PreCheckScript {
                try {
                    $res = Invoke-Command -ComputerName $srvFqdn -ScriptBlock {
                        param($tName, $sPath)
                        $task = Get-ScheduledTask -TaskName $tName -ErrorAction SilentlyContinue
                        $hasScr = Test-Path $sPath
                        $taskArg = if ($task) {
                            ($task.Actions | Where-Object { $_.Execute -match 'powershell' } |
                                Select-Object -First 1).Arguments
                        } else { '' }
                        [PSCustomObject]@{
                            TaskFound   = [bool]$task
                            ScriptFound = $hasScr
                            TaskArg     = $taskArg
                        }
                    } -ArgumentList $taskName, $deployPath -ErrorAction Stop
                    if (-not $res.TaskFound -or -not $res.ScriptFound) { return $false }
                    # AlreadyDone only if the LogPath in the task matches the
                    # runtime-detected LogPath. If different (e.g. Step06 moved the
                    # logs after Step17), the task must be registered again.
                    $taskLogPath = Get-LogPathFromTaskArgument -Argument $res.TaskArg
                    if (-not $taskLogPath) { return $false }
                    return ($taskLogPath -ieq $resolvedLogRoot)
                } catch { return $false }
            } `
            -ActionScript {
                if ($Mode -eq 'Simulate') {
                    Write-Log "  [WhatIf] Would copy $scriptSource -> $srvFqdn at $deployPath" -Level Sub
                    Write-Log "  [WhatIf] Would create task '$taskName' ($runDay $runTime) under $accountLabel with LogPath=$resolvedLogRoot" -Level Sub
                    return
                }

                $session = New-PSSession -ComputerName $srvFqdn -ErrorAction Stop
                try {
                    # 1. Create the target directory if absent
                    Invoke-Command -Session $session -ScriptBlock {
                        param($dir)
                        if (-not (Test-Path $dir)) {
                            New-Item -Path $dir -ItemType Directory -Force | Out-Null
                        }
                    } -ArgumentList (Split-Path $deployPath -Parent) | Out-Null

                    # 2. Copy the script
                    Copy-Item -Path $scriptSource -Destination $deployPath -ToSession $session -Force
                    Write-Log "  Script copied to $srvName at $deployPath" -Level Sub

                    # 3. Creation of the scheduled task
                    Invoke-Command -Session $session -ScriptBlock {
                        param($tName, $sPath, $day, $time, $logRoot, $keep, $age, $acct, $pwd)

                        # -Command is used (not -File) to be able to chain Compress then Purge
                        # with ';'. With -File, '; & script -Action Purge' is interpreted as
                        # additional positional parameters of the 1st script, which produces
                        # a binding failure and an immediate exit 1 (return code 2147942401).
                        # -Command reinterprets the string as a PowerShell shell command.
                        $cmd = "& '$sPath' -Action Compress -LogPath '$logRoot' -RetainUncompressedDays $keep -PurgeAfterDays $age; " +
                               "& '$sPath' -Action Purge -LogPath '$logRoot' -RetainUncompressedDays $keep -PurgeAfterDays $age"
                        $arg = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command `"$cmd`""

                        $action   = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
                        $trigger  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $day -At $time
                        $settings = New-ScheduledTaskSettingsSet `
                            -StartWhenAvailable `
                            -AllowStartIfOnBatteries `
                            -DontStopIfGoingOnBatteries `
                            -ExecutionTimeLimit (New-TimeSpan -Hours 4)

                        # Removal if the task already exists (idempotent / update)
                        $existing = Get-ScheduledTask -TaskName $tName -ErrorAction SilentlyContinue
                        if ($existing) { Unregister-ScheduledTask -TaskName $tName -Confirm:$false }

                        if ($acct -and $pwd) {
                            # Service account with password
                            $principal = New-ScheduledTaskPrincipal -UserId $acct -LogonType Password -RunLevel Highest
                            Register-ScheduledTask -TaskName $tName `
                                -Action $action -Trigger $trigger -Settings $settings `
                                -Principal $principal -Password $pwd -ErrorAction Stop | Out-Null
                        } elseif ($acct) {
                            # gMSA: no password (managed by AD)
                            $principal = New-ScheduledTaskPrincipal -UserId $acct -LogonType Password -RunLevel Highest
                            Register-ScheduledTask -TaskName $tName `
                                -Action $action -Trigger $trigger -Settings $settings `
                                -Principal $principal -ErrorAction Stop | Out-Null
                        } else {
                            # SYSTEM (default)
                            $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
                            Register-ScheduledTask -TaskName $tName `
                                -Action $action -Trigger $trigger -Settings $settings `
                                -Principal $principal -ErrorAction Stop | Out-Null
                        }
                    } -ArgumentList $taskName, $deployPath, $runDay, $runTime, $resolvedLogRoot, $keepDays, $purgeDays, $taskAccount, $taskPlainPwd | Out-Null

                    Write-Log "  Task '$taskName' created on $srvName ($runDay $runTime) LogPath=$resolvedLogRoot" -Level Sub
                } finally {
                    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
                }
            }
    }

    Save-Report -StepName 'Step17-ManageIISLogs' | Out-Null
    return 0
}
