<#
.SYNOPSIS
    Step 21 - Manual completion of the MigrationBatch objects.

.DESCRIPTION
    This step can be re-run safely and has two uses:

    1. Reporting: without any additional parameter, it reports the state of
       the MigrationBatch objects (Status, TotalCount, SyncedCount,
       FinalizedCount, FailedCount).

    2. Completion: with the environment variable BATCH=<number> (or the
       -BatchNumber parameter in the orchestrator script), it triggers
       Complete-MigrationBatch -Identity BatchNN on the batch concerned
       (Exchange completes all the MRs of the batch).

    To complete all the batches: BATCH=ALL.
    To complete only some of them: BATCH=01 (then 02, 03, ...).

    Refactor v1.2: uses Complete-MigrationBatch (native Exchange object)
    instead of an individual Resume-MoveRequest per MR. Benefits:
      - Consistency with Step20, which creates MigrationBatch objects via
        New-MigrationBatch
      - 1 command per batch instead of N commands per MR
      - Visibility in EAC > Migration > Batches

    v1.3: -ScheduledCompletionTime delegates the timing to Exchange. Actual
    behaviour on Exchange 2019 on-premises (tested 2026-05-16 on Batch04):

      - `Complete-MigrationBatch` IGNORES the `-CompleteAfter` set on the MRs:
        it performs an immediate implicit Resume, overwrites the value, and
        MRS finalizes without waiting. Scheduling through
        Complete-MigrationBatch is NOT possible on-premises.

      - The only pattern that honors the timing is:
          1. `Set-MoveRequest    -CompleteAfter $when` on each MR of the batch
             (pipe Get-MoveRequest -BatchName "MigrationService:BatchNN")
          2. `Resume-MoveRequest` directly on each MR
        MRS resumes each MR but honors its `CompleteAfter` (= it stays in
        Synced/AutoSuspended until the given time, then finalizes).

      - Trade-off: `Complete-MigrationBatch` is NOT called in scheduled mode,
        so the MigrationBatch stays at `Status=Synced` after the MRs are
        finalized. The EAC will not reflect the real end. Track progress via
        `Get-MoveRequest` / `Get-MoveRequestStatistics`, not via
        `Get-MigrationBatch`.

    The step ends after setting the instruction, the console can be closed.

    v1.4: -FollowAfter (orchestrator side) - after setting the
    Complete-MigrationBatch instruction, automatically chains into Step20
    -Follow in the same session. Combined with -BatchNumber NN, Step20 enters
    single-batch focus mode: only this batch is queried (Get-MoveRequest
    -BatchName, Get-MigrationBatch -Identity), the generated HTML is
    MigrationStatus-BatchNN.html and its title mentions "(focus)". This makes
    it possible to follow a completion without flooding the session logs with
    the 11 other batches.

    Idempotency: Complete-MigrationBatch on a batch already Completed returns
    an error that is captured cleanly. A batch with Status='Syncing' cannot be
    completed (Exchange requires Status='Synced').

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

    # The batch number can be passed:
    #   1) via the BATCH environment variable (e.g. $env:BATCH = '03')
    #   2) via $global:BatchNumber set by the orchestrator
    $batchNumber   = $env:BATCH
    if (-not $batchNumber -and (Test-Path Variable:Global:BatchNumber)) {
        $batchNumber = $Global:BatchNumber
    }

    $scheduledTime = $null
    $schedRaw = if ($env:SCHEDULE_TIME) { $env:SCHEDULE_TIME }
                elseif ((Test-Path Variable:Global:ScheduledCompletionTime)) { $Global:ScheduledCompletionTime }
    if ($schedRaw) {
        try { $scheduledTime = [DateTime]::Parse($schedRaw) } catch {
            Write-Log "Invalid ScheduledCompletionTime format: '$schedRaw' - ignored" -Level Warning
        }
    }

    $schedInfo = if ($scheduledTime) {
        "Scheduled time: $($scheduledTime.ToString('yyyy-MM-dd HH:mm')) (Set-MoveRequest -CompleteAfter + Resume-MoveRequest per MR, without Complete-MigrationBatch; batch will stay Synced in EAC)"
    } else { 'Immediate execution (no scheduled time)' }

    Write-StepBanner -StepName '21' -Title 'Manual completion of MigrationBatch' -Mode $Mode -Actions @(
        'List the MigrationBatch objects (BatchNN) with their Status + stats',
        'Show progress per batch (TotalCount / SyncedCount / FinalizedCount / FailedCount)',
        'If BATCH=<NN>: Complete-MigrationBatch -Identity BatchNN (Status=Synced required)',
        'If BATCH=ALL: Complete-MigrationBatch on all Synced batches',
        'If -ScheduledCompletionTime: Set-MoveRequest -CompleteAfter + Resume-MoveRequest on each MR (without Complete-MigrationBatch - see header)',
        $schedInfo
    )

    Reset-Report
    Initialize-ExchangeShell -Credential $Credential

    # --- 1) Global snapshot ------------------------------------------------
    # Source of truth = Get-MigrationBatch (name filter Batch\d{2} - excludes manual/legacy).
    $allBatches = @(Get-MigrationBatch -ErrorAction SilentlyContinue |
                    Where-Object { [string]$_.Identity -match '^Batch\d{2}$' })

    if ($allBatches.Count -eq 0) {
        Write-Log "No MigrationBatch 'Batch\d{2}' found. Run Step19+Step20 first." -Level Warning
        Save-Report -StepName 'Step21-CompleteMigration' | Out-Null
        return 0
    }

    Write-Log '===== STATUS PER MIGRATIONBATCH =====' -Level Step
    foreach ($b in $allBatches | Sort-Object Identity) {
        $bid    = [string]$b.Identity
        $bStat  = [string]$b.Status
        $total  = [int]$b.TotalCount
        $synced = [int]$b.SyncedCount
        $final  = [int]$b.FinalizedCount
        $failed = [int]$b.FailedCount

        Write-Log ("  {0,-10} | Status={1,-12} | Total={2,4} | Synced={3,4} | Finalized={4,4} | Failed={5,4}" `
                   -f $bid, $bStat, $total, $synced, $final, $failed) -Level Info

        Add-Report -Step '21-Complete' -Target $bid -Action 'Snapshot MigrationBatch' -Status Inventoried `
            -AfterValue ("Status={0} | Total={1} | Synced={2} | Finalized={3} | Failed={4}" -f `
                        $bStat, $total, $synced, $final, $failed)
    }
    Write-Log '====================================' -Level Step

    # --- 2) Inventory mode: stop here --------------------------------------
    if ($Mode -eq 'Inventory') {
        Save-Report -StepName 'Step21-CompleteMigration' | Out-Null
        return 0
    }

    # --- 3) If no batch requested, reporting only --------------------------
    if (-not $batchNumber) {
        Write-Log "No batch requested (BATCH=<NN> or BATCH=ALL not set). Reporting only." -Level Warning
        Save-Report -StepName 'Step21-CompleteMigration' | Out-Null
        return 0
    }

    # --- 4) Selection of the batches to complete ---------------------------
    # Criterion: a batch can be completed when its Status is 'Synced' (all the
    # underlying MRs are AutoSuspended and ready for completion).
    # Note: [string] cast on Status because in a fresh remoting session, the
    # MigrationBatchStatus enum is serialized as a typed PSObject and
    # `$_.Status -eq 'Synced'` does not match (empty filter). The [string]
    # forces the enum's .ToString().
    # $batchNumber accepts: 'ALL', a scalar string ('02'), or an array ('06','07','08').
    $batchItems = @($batchNumber)
    if ($batchItems.Count -eq 1 -and $batchItems[0] -eq 'ALL') {
        $toComplete = @($allBatches | Where-Object { [string]$_.Status -eq 'Synced' })
        Write-Log "Completing ALL Synced batches: $($toComplete.Count) batch(es)" -Level Step
    } else {
        $paddedSet  = @($batchItems | ForEach-Object { 'Batch{0:D2}' -f [int]$_ })
        $toComplete = @($allBatches | Where-Object { ([string]$_.Identity -in $paddedSet) -and [string]$_.Status -eq 'Synced' })
        Write-Log ("Completing batches: {0} -> {1} batch(es) (Status=Synced)" -f ($paddedSet -join ', '), $toComplete.Count) -Level Step
    }

    if ($toComplete.Count -eq 0) {
        Write-Log "No MigrationBatch with Status='Synced' for the requested scope." -Level Warning
        $eligibleBatches = $allBatches | Where-Object { $_.Status -ne 'Completed' }
        if ($eligibleBatches) {
            Write-Log "Current statuses:" -Level Sub
            foreach ($b in $eligibleBatches) {
                Write-Log ("    {0} : Status={1}" -f $b.Identity, $b.Status) -Level Sub
            }
        }
        Save-Report -StepName 'Step21-CompleteMigration' | Out-Null
        return 0
    }

    # --- 5) Determine whether to schedule on the server side ----------------
    # If a scheduled time is provided AND in the future: -CompleteAfter is set on each
    # MR, then Resume-MoveRequest is called directly. Complete-MigrationBatch
    # ignores CompleteAfter on-premises (immediate implicit Resume, value overwritten).
    # Otherwise: immediate completion via Complete-MigrationBatch (standard behaviour).
    $useServerScheduling = $false
    if ($scheduledTime -and (Get-Date) -lt $scheduledTime) {
        $useServerScheduling = $true
        Write-Log ("Scheduled completion: Set-MoveRequest -CompleteAfter={0} + Resume-MoveRequest per MR (without Complete-MigrationBatch)" -f $scheduledTime.ToString('yyyy-MM-dd HH:mm')) -Level Step
        Write-Log "  MigrationBatch will stay at Status=Synced after the MRs are finalized (track via Get-MoveRequest)." -Level Sub
        Write-Log "  The step sets the instruction and exits. The console can be closed." -Level Sub
    }

    # --- 6) Completion (immediate or scheduled) ----------------------------
    foreach ($b in $toComplete) {
        $bid = [string]$b.Identity

        $detailSuffix = if ($useServerScheduling) {
            "CompleteAfter=$($scheduledTime.ToString('yyyy-MM-dd HH:mm')) via Resume-MoveRequest"
        } else { 'immediate completion via Complete-MigrationBatch' }

        Invoke-Action -Step '21-Complete' -Target $bid -Action 'Complete-MigrationBatch' `
            -Detail "Status=$($b.Status); Total=$($b.TotalCount); Synced=$($b.SyncedCount); $detailSuffix" `
            -InventoryScript {
                $cur = Get-MigrationBatch -Identity $bid -ErrorAction SilentlyContinue
                if ($cur) { "Status=$($cur.Status); Finalized=$($cur.FinalizedCount)/$($cur.TotalCount)" } else { '<absent>' }
            } `
            -PreCheckScript {
                if ($useServerScheduling) {
                    # In scheduled mode the batch stays Synced -> check the MR state
                    # for idempotency. NB: read CompleteAfter via Get-MoveRequestStatistics
                    # because Get-MoveRequest serializes it badly via RPS (always empty on the
                    # client side, see memory feedback-getmoverequest-completeafter-serialization).
                    $mrs = @(Get-MoveRequest -BatchName ("MigrationService:" + $bid) -ResultSize Unlimited -ErrorAction SilentlyContinue |
                             Get-MoveRequestStatistics -ErrorAction SilentlyContinue)
                    $pending = @($mrs | Where-Object { [string]$_.Status -notin @('Completed','CompletionInProgress') })
                    if ($pending.Count -eq 0) { return $true }
                    return -not (@($pending | Where-Object { -not $_.CompleteAfter }).Count -gt 0)
                } else {
                    $cur = Get-MigrationBatch -Identity $bid -ErrorAction SilentlyContinue
                    if (-not $cur) { return $true }
                    return ([string]$cur.Status -in @('Completed','CompletedWithErrors'))
                }
            } `
            -ActionScript {
                if ($useServerScheduling) {
                    # BatchName on the MR side has a "MigrationService:" prefix
                    $mrCount = 0
                    Get-MoveRequest -BatchName ("MigrationService:" + $bid) -ResultSize Unlimited |
                        Where-Object { [string]$_.Status -notin @('Completed','CompletionInProgress') } |
                        ForEach-Object {
                            $mrId = $_.Identity
                            Set-MoveRequest    -Identity $mrId -CompleteAfter $scheduledTime -Confirm:$false -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                            Resume-MoveRequest -Identity $mrId -Confirm:$false -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                            $mrCount++
                        }
                    Write-Log ("    CompleteAfter={0} + Resume set on {1} MR(s) of batch {2} (MigrationBatch will stay Synced)" -f $scheduledTime.ToString('yyyy-MM-dd HH:mm'), $mrCount, $bid) -Level Sub
                } else {
                    Complete-MigrationBatch -Identity $bid -Confirm:$false -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                }
            }
    }

    # --- 7) Final snapshot (post-completion) ------------------------------
    if ($Mode -eq 'Apply') {
        Start-Sleep -Seconds 5
        $finalBatches = @(Get-MigrationBatch -ErrorAction SilentlyContinue |
                          Where-Object { [string]$_.Identity -match '^Batch\d{2}$' })
        $finalRows = foreach ($b in $finalBatches) {
            [PSCustomObject]@{
                Identity       = [string]$b.Identity
                Status         = [string]$b.Status
                TotalCount     = [int]$b.TotalCount
                SyncedCount    = [int]$b.SyncedCount
                FinalizedCount = [int]$b.FinalizedCount
                FailedCount    = [int]$b.FailedCount
                CompleteAfter  = $b.CompleteAfter
            }
        }
        $finalPath = Join-Path $OutputFolder 'MigrationBatchStatus.csv'
        $finalRows | Export-Csv -Path $finalPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        Add-Report -Step '21-Complete' -Target 'GLOBAL' -Action 'Final snapshot MigrationBatch' -Status Success `
                   -AfterValue $finalPath
    }

    Save-Report -StepName 'Step21-CompleteMigration' | Out-Null
    return 0
}
