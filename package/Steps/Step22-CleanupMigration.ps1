<#
.SYNOPSIS
    Step 22 - Post-migration cleanup (MoveRequest, MigrationBatch, MigrationUser).

.DESCRIPTION
    Purges the MoveRequest, MigrationBatch and MigrationUser objects left
    over after a completed migration.

    Default purge scope (terminal statuses):
      - MigrationBatch  : Synced, Completed, CompletedWithErrors
      - MoveRequest     : Completed
      - MigrationUser   : orphans (BatchId pointing to a missing MigrationBatch)

    Extended scope (environment variable CLEANUP_INCLUDE_FAILED='1'):
      - MigrationBatch  : adds Failed, Stopped, Corrupted
      - MoveRequest     : adds Failed

    Statuses that must NEVER be purged (in progress): Syncing, Completing,
    active AutoSuspended, InProgress. The step filters them out upstream.

    Main use case: after the Step21 pattern "Set-MoveRequest -CompleteAfter
    + Resume-MoveRequest" (without Complete-MigrationBatch), the
    MigrationBatch objects stay at Status=Synced even after the MRs are
    really finalized. This step detects them (Synced + 0 pending MR) and
    removes them cleanly.

    Order of operations:
      1. Inventory snapshot (affected MR + MB + MU)
      2. Apply: Remove-MoveRequest on each eligible MR
      3. Apply: Remove-MigrationBatch -Force (removes the associated MU in cascade)
      4. Apply: Remove-MigrationUser on the remaining orphans

    CSV report generated: one line per target (MB, MR, MU) with the status
    before, the action performed and the result (Purged / Failed / Skipped).

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

    $includeFailed = ($env:CLEANUP_INCLUDE_FAILED -eq '1')

    $mbEligibleStatuses = @('Synced','Completed','CompletedWithErrors')
    $mrEligibleStatuses = @('Completed')
    if ($includeFailed) {
        $mbEligibleStatuses += @('Failed','Stopped','Corrupted')
        $mrEligibleStatuses += @('Failed')
    }

    Write-StepBanner -StepName '22' -Title 'Post-migration cleanup (MR, MigrationBatch, MigrationUser)' -Mode $Mode -Actions @(
        ("Eligible MigrationBatch  : " + ($mbEligibleStatuses -join ', ')),
        ("Eligible MoveRequest     : " + ($mrEligibleStatuses -join ', ')),
        'Orphaned MigrationUser   : detection (BatchId pointing to a missing MB)',
        "Environment variable CLEANUP_INCLUDE_FAILED='1' to include the Failed statuses",
        "Inventory: lists the targets. Simulate: -WhatIf. Apply: actual deletion."
    )

    Reset-Report
    Initialize-ExchangeShell -Credential $Credential

    # --- 1) Snapshot: eligible MigrationBatch -------------------------------
    Write-Log '===== SCAN MigrationBatch =====' -Level Step
    $allBatches = @(Get-MigrationBatch -ErrorAction SilentlyContinue |
                    Where-Object { [string]$_.Identity -match '^Batch\d{2}$' })

    $mbToPurge = @($allBatches | Where-Object { [string]$_.Status -in $mbEligibleStatuses })
    Write-Log ("  MigrationBatch total: {0} | Eligible for purge: {1}" -f $allBatches.Count, $mbToPurge.Count) -Level Info

    foreach ($b in $allBatches | Sort-Object Identity) {
        $bid    = [string]$b.Identity
        $bStat  = [string]$b.Status
        $total  = [int]$b.TotalCount
        $final  = [int]$b.FinalizedCount
        $eligible = ($bStat -in $mbEligibleStatuses)
        $mark = if ($eligible) { 'PURGE' } else { 'KEEP ' }
        Write-Log ("  [{0}] {1,-10} | Status={2,-20} | Total={3,4} | Finalized={4,4}" -f $mark, $bid, $bStat, $total, $final) -Level Info

        Add-Report -Step '22-Cleanup' -Target $bid -Action 'Scan MigrationBatch' -Status Inventoried `
            -AfterValue ("Status={0} | Total={1} | Finalized={2} | Eligible={3}" -f $bStat, $total, $final, $eligible)
    }

    # --- 2) Snapshot: eligible MoveRequest ----------------------------------
    Write-Log '===== SCAN MoveRequest =====' -Level Step
    $allMrs = @(Get-MoveRequest -ResultSize Unlimited -ErrorAction SilentlyContinue)
    $mrToPurge = @($allMrs | Where-Object { [string]$_.Status -in $mrEligibleStatuses })
    Write-Log ("  MoveRequest total: {0} | Eligible for purge: {1}" -f $allMrs.Count, $mrToPurge.Count) -Level Info

    # Log one CSV line per MR to purge (summary), not for the 400+ non-eligible ones
    foreach ($mr in $mrToPurge | Sort-Object DisplayName) {
        Add-Report -Step '22-Cleanup' -Target ([string]$mr.DisplayName) -Action 'Scan MoveRequest' -Status Inventoried `
            -AfterValue ("Status={0} | BatchName={1} | Alias={2}" -f $mr.Status, $mr.BatchName, $mr.Alias)
    }

    # --- 3) Snapshot: orphaned MigrationUser --------------------------------
    # Note: Get-MigrationUser.BatchId is a STRING (batch name, e.g. "Batch01"),
    # not an object with .Guid. It is therefore compared directly with the
    # Identity of the current MigrationBatch objects (also strings over RPS remoting).
    Write-Log '===== SCAN orphaned MigrationUser =====' -Level Step
    $allMus    = @(Get-MigrationUser -ResultSize Unlimited -ErrorAction SilentlyContinue)
    $validBatchNames = @{}
    foreach ($b in (Get-MigrationBatch -ErrorAction SilentlyContinue)) {
        $validBatchNames[[string]$b.Identity] = $true
    }
    $muOrphans = @($allMus | Where-Object {
        $bn = [string]$_.BatchId
        # Orphan if BatchId is missing or not present in the current batch list
        (-not $bn) -or (-not $validBatchNames.ContainsKey($bn))
    })
    Write-Log ("  MigrationUser total: {0} | Orphans detected: {1}" -f $allMus.Count, $muOrphans.Count) -Level Info

    foreach ($mu in $muOrphans | Sort-Object Identity) {
        Add-Report -Step '22-Cleanup' -Target ([string]$mu.Identity) -Action 'Scan orphaned MigrationUser' -Status Inventoried `
            -AfterValue ("Status={0} | BatchId={1}" -f $mu.Status, $mu.BatchId)
    }

    Write-Log ('=' * 60) -Level Step

    # --- 4) Inventory mode: stop here ---------------------------------------
    if ($Mode -eq 'Inventory') {
        Write-Log "Inventory mode: nothing deleted. Re-run with -Mode Apply to purge." -Level Sub
        Save-Report -StepName 'Step22-CleanupMigration' | Out-Null
        return 0
    }

    if ($mbToPurge.Count -eq 0 -and $mrToPurge.Count -eq 0 -and $muOrphans.Count -eq 0) {
        Write-Log "Nothing to purge. Exiting." -Level Success
        Save-Report -StepName 'Step22-CleanupMigration' | Out-Null
        return 0
    }

    # --- 5) Apply / Simulate: Remove-MoveRequest first ----------------------
    # The eligible MRs are removed BEFORE the MigrationBatch objects:
    # Remove-MigrationBatch can fail if still-active MRs are attached to it.
    # Completed MRs are independent of the batch and can be removed at any time.
    Write-Log "===== PURGE MoveRequest =====" -Level Step
    foreach ($mr in $mrToPurge) {
        $mrId   = [string]$mr.Identity
        $mrName = [string]$mr.DisplayName
        $mrStat = [string]$mr.Status

        Invoke-Action -Step '22-Cleanup' -Target $mrName -Action 'Remove-MoveRequest' `
            -Detail "Status=$mrStat; Identity=$mrId; BatchName=$($mr.BatchName)" `
            -InventoryScript {
                $cur = Get-MoveRequest -Identity $mrId -ErrorAction SilentlyContinue
                if ($cur) { "Status=$($cur.Status)" } else { '<absent>' }
            } `
            -PreCheckScript {
                $cur = Get-MoveRequest -Identity $mrId -ErrorAction SilentlyContinue
                return (-not $cur)
            } `
            -ActionScript {
                Remove-MoveRequest -Identity $mrId -Confirm:$false -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
            }
    }

    # --- 6) Apply / Simulate: Remove-MigrationBatch ------------------------
    # -Force: also removes the associated MigrationUser objects in cascade (recommended).
    Write-Log "===== PURGE MigrationBatch =====" -Level Step
    foreach ($b in $mbToPurge) {
        $bid   = [string]$b.Identity
        $bStat = [string]$b.Status

        Invoke-Action -Step '22-Cleanup' -Target $bid -Action 'Remove-MigrationBatch -Force' `
            -Detail "Status=$bStat; Total=$($b.TotalCount); Finalized=$($b.FinalizedCount)" `
            -InventoryScript {
                $cur = Get-MigrationBatch -Identity $bid -ErrorAction SilentlyContinue
                if ($cur) { "Status=$($cur.Status); Total=$($cur.TotalCount)" } else { '<absent>' }
            } `
            -PreCheckScript {
                $cur = Get-MigrationBatch -Identity $bid -ErrorAction SilentlyContinue
                return (-not $cur)
            } `
            -ActionScript {
                Remove-MigrationBatch -Identity $bid -Force -Confirm:$false -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
            }
    }

    # --- 7) Apply / Simulate: remove orphaned MigrationUser objects ---------
    # After Remove-MigrationBatch -Force, the MU objects are re-scanned to catch the
    # remaining orphans (SyncFailed, etc.) that were not removed in cascade.
    #
    # IMPORTANT: Get-MigrationUser over RPS can return MU objects that were already
    # deleted on the backend (stale cache ~10-15s after Remove-MigrationBatch -Force).
    # Symptom: Remove-MigrationUser raises "Could not load the batch information
    # for migration user '...'. Associated migration subscription cannot be removed."
    # -> Exchange warning, not an error; the MU IS already deleted on the backend.
    # The script sleeps 15s to let the cache refresh before the re-scan.
    if ($Mode -eq 'Apply') {
        Start-Sleep -Seconds 15
        $remainingValidBatchNames = @{}
        foreach ($b in (Get-MigrationBatch -ErrorAction SilentlyContinue)) {
            $remainingValidBatchNames[[string]$b.Identity] = $true
        }
        $muOrphansAfter = @(Get-MigrationUser -ResultSize Unlimited -ErrorAction SilentlyContinue |
                            Where-Object {
                                $bn = [string]$_.BatchId
                                (-not $bn) -or (-not $remainingValidBatchNames.ContainsKey($bn))
                            })
    } else {
        $muOrphansAfter = $muOrphans
    }

    if ($muOrphansAfter.Count -gt 0) {
        Write-Log "===== PURGE residual orphaned MigrationUser =====" -Level Step
        foreach ($mu in $muOrphansAfter) {
            $muId   = [string]$mu.Identity
            $muStat = [string]$mu.Status

            Invoke-Action -Step '22-Cleanup' -Target $muId -Action 'Remove-MigrationUser' `
                -Detail "Status=$muStat; BatchId=$($mu.BatchId.Guid)" `
                -InventoryScript {
                    $cur = Get-MigrationUser -Identity $muId -ErrorAction SilentlyContinue
                    if ($cur) { "Status=$($cur.Status)" } else { '<absent>' }
                } `
                -PreCheckScript {
                    $cur = Get-MigrationUser -Identity $muId -ErrorAction SilentlyContinue
                    return (-not $cur)
                } `
                -ActionScript {
                    Remove-MigrationUser -Identity $muId -Confirm:$false -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
                }
        }
    } else {
        Write-Log "No orphaned MigrationUser to purge." -Level Sub
    }

    # --- 8) Final snapshot (post-purge) -------------------------------------
    if ($Mode -eq 'Apply') {
        Start-Sleep -Seconds 2
        $finalMb = @(Get-MigrationBatch -ErrorAction SilentlyContinue | Where-Object { [string]$_.Identity -match '^Batch\d{2}$' }).Count
        $finalMr = @(Get-MoveRequest -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
        $finalMu = @(Get-MigrationUser -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
        Write-Log ("===== FINAL STATE =====") -Level Step
        Write-Log ("  MigrationBatch remaining: {0}" -f $finalMb) -Level Info
        Write-Log ("  MoveRequest remaining:    {0}" -f $finalMr) -Level Info
        Write-Log ("  MigrationUser remaining:  {0}" -f $finalMu) -Level Info

        Add-Report -Step '22-Cleanup' -Target 'GLOBAL' -Action 'Final snapshot' -Status Success `
            -AfterValue ("MigrationBatch={0} | MoveRequest={1} | MigrationUser={2}" -f $finalMb, $finalMr, $finalMu)
    }

    Save-Report -StepName 'Step22-CleanupMigration' | Out-Null
    return 0
}
