<#
.SYNOPSIS
    Step 14 - Set the Application and System event log size to 1 GB.

.DESCRIPTION
    Pre-check: the presence of a GPO applying
    HKLM\Software\Policies\Microsoft\Windows\EventLog\<Log>\MaxSize is detected.
    If so, no change is attempted (the GPO would overwrite the change at each
    gpupdate) and a warning is raised:
        "Warning: the log sizes are managed by GPO.
         The GPO must be modified to change the size of the
         Application and System logs"

    Otherwise: Limit-EventLog -LogName <Log> -MaximumSize <bytes> on each
    Exchange 2019 server. -WhatIf is propagated in Simulate mode so that NOTHING
    is modified (a simulation run must not change the target).

    The current size is read through Get-EventLog -List -> MaximumKilobytes
    property (in kilobytes). Get-LogProperties is not available on these
    servers.

    Idempotent: if the current size = target exactly, AlreadyDone.
    A larger OR smaller size is considered non-compliant and is set back to
    the target.

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

    $targetBytes = [int64]$Config.EventLogs.TargetSizeBytes
    $targetKb    = [int64]($targetBytes / 1KB)
    $logs        = $Config.EventLogs.Logs   # @('Application','System')
    $isWhatIf    = ([bool]($Mode -eq 'Simulate'))

    Write-StepBanner -StepName '14' -Title "Event log size set to $($targetBytes/1GB) GB" -Mode $Mode -Actions @(
        'On each Exchange 2019 server:',
        '  - Detect whether HKLM\Software\Policies\Microsoft\Windows\EventLog\<Log>\MaxSize is defined by GPO',
        '  - If GPO: warning and no change',
        "  - Otherwise: Limit-EventLog -MaximumSize $targetBytes (with -WhatIf in Simulate)",
        "  - Target logs: $($logs -join ', ')"
    )

    Initialize-ExchangeShell -Credential $Credential
    Reset-Report
    $servers = Get-Exchange2019Servers
    if (-not $servers) {
        Write-Log "No Exchange 2019 server detected." -Level Warning
        Save-Report -StepName 'Step14-EventLogSize'
        return 0
    }

    foreach ($srv in $servers) {
        $icArgs = @{ ComputerName = $srv.Fqdn }
        if ($Credential) { $icArgs['Credential'] = $Credential }

        foreach ($log in $logs) {
            Invoke-Action -Step '14-EvtLog' -Target "$($srv.Name)\$log" -Action "Set MaxSize=$targetKb KB" `
                -InventoryScript {
                    $sb = {
                        param($log)
                        $current = 0
                        $entry = Get-EventLog -List | Where-Object { $_.Log -eq $log }
                        if ($entry) { $current = [int64]$entry.MaximumKilobytes }
                        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\EventLog\$log"
                        $byGpo   = $false
                        if (Test-Path $regPath) {
                            $v = Get-ItemProperty -Path $regPath -Name 'MaxSize' -ErrorAction SilentlyContinue
                            if ($v -and $v.MaxSize) { $byGpo = $true }
                        }
                        "Current={0} KB | GPO={1}" -f $current, $byGpo
                    }
                    Invoke-Command @icArgs -ScriptBlock $sb -ArgumentList $log
                } `
                -PreCheckScript {
                    $sb = {
                        param($log, $targetKb)
                        $entry = Get-EventLog -List | Where-Object { $_.Log -eq $log }
                        if (-not $entry) { return $false }
                        # Strict equality: any different value triggers a Set
                        return ([int64]$entry.MaximumKilobytes -eq $targetKb)
                    }
                    Invoke-Command @icArgs -ScriptBlock $sb -ArgumentList $log, $targetKb
                } `
                -ActionScript {
                    $result = Invoke-Command @icArgs -ScriptBlock {
                        param($log, $targetBytes, $whatIf)
                        # GPO check: if MaxSize is managed by GPO, do not modify
                        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\EventLog\$log"
                        if (Test-Path $regPath) {
                            $v = Get-ItemProperty -Path $regPath -Name 'MaxSize' -ErrorAction SilentlyContinue
                            if ($v -and $v.MaxSize) {
                                return @{ ManagedByGPO = $true ; CurrentMaxSize = $v.MaxSize }
                            }
                        }
                        # Limit-EventLog: -WhatIf in Simulate => no real change
                        Limit-EventLog -LogName $log -MaximumSize $targetBytes -WhatIf:$whatIf -ErrorAction Stop
                        return @{ ManagedByGPO = $false ; CurrentMaxSize = $targetBytes }
                    } -ArgumentList $log, $targetBytes, $isWhatIf -ErrorAction Stop

                    if ($result.ManagedByGPO) {
                        $msg = "Warning: the log sizes are managed by GPO. The GPO must be modified to change the size of the Application and System logs"
                        Write-Log "[$($srv.Name)\$log] $msg (MaxSize GPO=$($result.CurrentMaxSize))" -Level Warning
                        throw $msg   # captured by Invoke-Action -> Failed (warning visible in the report)
                    }
                }
        }
    }

    Save-Report -StepName 'Step14-EventLogSize' | Out-Null
    return 0
}
