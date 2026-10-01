<#
.SYNOPSIS
    Step 19 - Migration batch preparation.

.DESCRIPTION
    Builds the distribution of the mailboxes to migrate into balanced batches.

    Implementation choices (v1.3+):
        - Step20 uses New-MigrationBatch -Local to create real MigrationBatch
          objects (visible in EAC > Migration > Batches). Step21 uses
          Complete-MigrationBatch.
        - Primary + Archive of the same mailbox are consolidated in the same
          batch (otherwise MigrationBatch conflict: a mailbox can only be in 1
          active batch).
        - PublicFolderMailbox: stays on New-MoveRequest -PublicFolder
          (MigrationBatch -Local limitation). The BatchName stays consistent
          (Batch01...) but the MigrationBatch object does not contain these MRs.
        - System mailboxes (Arbitration / AuditLog / AuxAuditLog / Discovery)
          are handled by Step18 and excluded here to avoid duplicates.
        - Monitoring mailboxes are fully excluded: they are not migrated and
          will be removed when the legacy servers are uninstalled.
        - Balanced distribution by size AND by count using the LPT algorithm
          + count cap (see Export-MailboxMigrationBatches.ps1).
        - Produces a global CSV "MigrationPlan.csv" in Reports\, consumed by
          the following steps (Step20-RunMigration, Step21-CompleteMigration).
        - In Apply/Simulate mode: HTML report "MigrationPlan.html" generated
          to visualize the distribution per batch (relative bars, details per
          mailbox).
        - No migration action is performed here: only the distribution is
          prepared.

    Filters:
        - HealthMailbox* excluded.
        - System mailboxes (Arbitration/AuditLog/AuxAuditLog/Discovery) excluded -> Step18.
        - Only the mailboxes whose primary server is an Exchange 2013 or 2016
          are selected (otherwise there is nothing to migrate).

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

    $batchCount = [int]$Config.Migration.BatchCount

    # =========================================================================
    # Helper: Exchange size -> MB conversion
    # =========================================================================
    function ConvertTo-SizeMB {
        param([object]$Size)
        if ($null -eq $Size) { return 0.00 }
        $txt = [string]$Size
        if ([string]::IsNullOrWhiteSpace($txt)) { return 0.00 }
        if ($txt -match '([\d,\.\s]+)\s*bytes') {
            $raw = ($Matches[1] -replace '[,\s]', '')
            [double]$b = 0
            if ([double]::TryParse($raw,
                [System.Globalization.NumberStyles]::Any,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [ref]$b)) {
                if ($b -gt 0) { return [math]::Round($b / 1MB, 2) }
            }
        }
        return 0.00
    }

    # =========================================================================
    # Helper: migration plan HTML report
    # =========================================================================
    function New-MigrationPlanHtml {
        param(
            [System.Collections.Generic.List[object]]$Data,
            [hashtable]$Batches,
            [string]$OutputPath,
            [string]$ReportMode,
            [int]$BatchCount
        )

        function Esc {
            param([string]$s)
            if ([string]::IsNullOrEmpty($s)) { return '' }
            $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
        }

        $now   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

        $primary  = @($Data | Where-Object Type -eq 'Primary')
        $archives = @($Data | Where-Object Type -eq 'Archive')
        $totalSizeMB    = ($Data | Measure-Object SizeMB -Sum).Sum
        $totalSizeGB    = [math]::Round($totalSizeMB / 1024, 2)
        $maxBatchSizeMB = ($Batches.Values | Measure-Object SizeMB -Maximum).Maximum
        if (-not $maxBatchSizeMB -or $maxBatchSizeMB -eq 0) { $maxBatchSizeMB = 1 }

        $typeClr = @{
            'UserMailbox'         = '#1a9e45'  # green
            'SharedMailbox'       = '#0d8c93'  # turquoise
            'RoomMailbox'         = '#d97706'  # amber
            'EquipmentMailbox'    = '#7c3aed'  # purple
            'Archive'             = '#2563eb'  # royal blue
            'ArbitrationMailbox'  = '#6b7280'  # gray
            'AuditLogMailbox'     = '#6b7280'
            'AuxAuditLogMailbox'  = '#6b7280'
            'DiscoveryMailbox'    = '#6b7280'
            'PublicFolderMailbox' = '#be185d'  # pink
            'default'             = '#5f6b7a'
        }

        $sb = [System.Text.StringBuilder]::new(200000)

        # ---- HEAD and HEADER: shared report theme of the module ------------
        # Classes of the batch overview and of the mailbox tables, on top of the shared theme.
        $extraStyle = '.td-dn{font-weight:600}.td-mb{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums;color:var(--cp-text-muted)}' +
                      '.bn{font-weight:700}.nr{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}.ov-tbl th.r{text-align:right}' +
                      '.bar-wrap{background:var(--cp-border);border-radius:8px;height:8px;overflow:hidden;min-width:120px}.bar-fill{height:8px;border-radius:8px;background:var(--cp-accent)}' +
                      'tr.arc td{background:var(--cp-accent-soft)}'
        [void]$sb.Append((Get-ExHtmlHead -Title "Exchange 2019 migration | Migration plan $stamp" -ExtraStyle $extraStyle))
        $maxGb = [math]::Round($maxBatchSizeMB / 1024, 1)
        $tiles = @(
            (Get-ExHtmlTile -Label 'Mailboxes to migrate' -Value (Format-ExNumber $primary.Count) -Note ('+ {0} in-place archive(s)' -f $archives.Count) -Kind Primary -Icon 'mail')
            (Get-ExHtmlTile -Label 'Migration batches' -Value (Format-ExNumber $BatchCount) -Note 'Balanced by size, primary and archive together' -Kind Tinted -Icon 'list')
            (Get-ExHtmlTile -Label 'Total volume (GB)' -Value ([string]::Format([Globalization.CultureInfo]::GetCultureInfo('en-US'), '{0:N1}', $totalSizeGB)) -Note 'Primary mailboxes and archives' -Kind Tinted -Icon 'chart')
            (Get-ExHtmlTile -Label 'Largest batch (GB)' -Value ([string]::Format([Globalization.CultureInfo]::GetCultureInfo('en-US'), '{0:N1}', $maxGb)) -Note 'The bars below are relative to it' -Icon 'server')
        )
        $box = "Mode <strong>$ReportMode</strong><br>$now<br>$BatchCount batch(es)"
        [void]$sb.Append((Get-ExHtmlHero -Title 'Migration plan' -Subtitle 'Mailboxes of the legacy servers, distributed into balanced migration batches (Step 19). Step 20 creates one Exchange migration batch per batch of this plan, Step 21 completes them.' `
            -BoxLabel 'Plan' -BoxHtml $box -Tiles $tiles))

        # ---- BATCH OVERVIEW -------------------------------------------------
        [void]$sb.AppendLine("<div class='card'><div class='card-hdr'>Batch overview &mdash; relative distribution</div>")
        [void]$sb.AppendLine("<div class='tbl-wrap'><table class='ov-tbl'>")
        [void]$sb.AppendLine("  <thead><tr>")
        [void]$sb.AppendLine("    <th style='min-width:80px'>Batch</th>")
        [void]$sb.AppendLine("    <th class='r' style='min-width:70px'>Mailboxes</th>")
        [void]$sb.AppendLine("    <th class='r' style='min-width:70px'>Archives</th>")
        [void]$sb.AppendLine("    <th class='r' style='min-width:110px'>Volume (MB)</th>")
        [void]$sb.AppendLine("    <th class='r' style='min-width:80px'>Volume (GB)</th>")
        [void]$sb.AppendLine("    <th style='min-width:220px'>Relative distribution</th>")
        [void]$sb.AppendLine("  </tr></thead><tbody>")

        foreach ($b in ($Batches.Values | Sort-Object Number)) {
            $bName   = "Batch{0:D2}" -f $b.Number
            $aid     = "batch-$($b.Number)"
            $nbArc   = ($Data | Where-Object { $_.Batch -eq $b.Number -and $_.Type -eq 'Archive' }).Count
            $pct     = [math]::Round($b.SizeMB * 100 / $maxBatchSizeMB)
            $szMbFmt = "{0:N0}" -f $b.SizeMB
            $szGbFmt = "{0:N2}" -f ($b.SizeMB / 1024)
            [void]$sb.AppendLine("  <tr>")
            [void]$sb.AppendLine("    <td class='bn'><a href='#$aid'>$bName</a></td>")
            [void]$sb.AppendLine("    <td class='nr'>$($b.Count)</td>")
            [void]$sb.AppendLine("    <td class='nr'>$nbArc</td>")
            [void]$sb.AppendLine("    <td class='nr'>$szMbFmt</td>")
            [void]$sb.AppendLine("    <td class='nr'>$szGbFmt</td>")
            [void]$sb.AppendLine("    <td><div class='bar-wrap'><div class='bar-fill' style='width:$pct%'></div></div></td>")
            [void]$sb.AppendLine("  </tr>")
        }

        [void]$sb.AppendLine("</tbody></table></div></div>")

        # ---- TOOLBAR + ACCORDIONS PER BATCH ---------------------------------
        [void]$sb.AppendLine("<div class='toolbar'><span class='label' style='margin:0'>Detail per batch</span>")
        [void]$sb.AppendLine("    <button class='btn' onclick=""document.querySelectorAll('details').forEach(function(d){d.open=true})"">Expand all</button>")
        [void]$sb.AppendLine("    <button class='btn' onclick=""document.querySelectorAll('details').forEach(function(d){d.open=false})"">Collapse all</button></div>")

        foreach ($b in ($Batches.Values | Sort-Object Number)) {
            $batchItems = @($Data | Where-Object Batch -eq $b.Number | Sort-Object @{E='Type';Descending=$false}, DisplayName)
            $bName   = "Batch{0:D2}" -f $b.Number
            $aid     = "batch-$($b.Number)"
            $szMbFmt = "{0:N0}" -f $b.SizeMB
            $szGbFmt = "{0:N2}" -f ($b.SizeMB / 1024)

            # Pills per type for the accordion summary
            $typePills = ''
            foreach ($tg in ($batchItems | Group-Object RecipientTypeDetails | Sort-Object Count -Descending)) {
                $tc = if ($typeClr[$tg.Name]) { $typeClr[$tg.Name] } else { $typeClr['default'] }
                $typePills += "<span class='pill' style='background:$tc'>$(Esc $tg.Name)&nbsp;$($tg.Count)</span>"
            }

            [void]$sb.AppendLine("<details id='$aid'>")
            [void]$sb.AppendLine("  <summary>")
            [void]$sb.AppendLine("    <span class='acc-icon'>&#9654;</span>")
            [void]$sb.AppendLine("    <span class='acc-step'>$bName</span>")
            [void]$sb.AppendLine("    <span class='acc-cnt'>$($b.Count)&thinsp;mailbox(es) &mdash; $szMbFmt&thinsp;MB&thinsp;($szGbFmt&thinsp;GB)</span>")
            [void]$sb.AppendLine("    <span class='acc-pills'>$typePills</span>")
            [void]$sb.AppendLine("  </summary>")
            [void]$sb.AppendLine("  <div class='tbl-wrap'><table>")
            [void]$sb.AppendLine("    <thead><tr>")
            [void]$sb.AppendLine("      <th style='min-width:200px'>Display name</th>")
            [void]$sb.AppendLine("      <th style='min-width:120px'>Alias</th>")
            [void]$sb.AppendLine("      <th style='min-width:130px'>Type</th>")
            [void]$sb.AppendLine("      <th style='min-width:160px'>Source database</th>")
            [void]$sb.AppendLine("      <th style='min-width:90px;text-align:right'>Size&thinsp;(MB)</th>")
            [void]$sb.AppendLine("      <th style='min-width:65px;text-align:center'>Archive</th>")
            [void]$sb.AppendLine("    </tr></thead><tbody>")

            foreach ($mbx in $batchItems) {
                $tc    = if ($typeClr[$mbx.RecipientTypeDetails]) { $typeClr[$mbx.RecipientTypeDetails] } else { $typeClr['default'] }
                $dn    = Esc $mbx.DisplayName
                $al    = Esc $mbx.Alias
                $rtype = Esc $mbx.RecipientTypeDetails
                $db    = Esc ([string]$mbx.SourceDatabase)
                $szFmt = "{0:N0}" -f $mbx.SizeMB
                $arcIco = if ($mbx.Type -eq 'Primary' -and $mbx.HasArchive) {
                    "<span style='color:var(--cp-success);font-weight:700;font-size:13px'>&#10003;</span>"
                } else {
                    "<span class='em'>&mdash;</span>"
                }
                $rowBg = if ($mbx.Type -eq 'Archive') { " class='arc'" } else { '' }

                [void]$sb.AppendLine("    <tr$rowBg>")
                [void]$sb.AppendLine("      <td class='td-dn'>$dn</td>")
                [void]$sb.AppendLine("      <td><span class='val'>$al</span></td>")
                [void]$sb.AppendLine("      <td><span class='bdg' style='background:$tc'>$rtype</span></td>")
                [void]$sb.AppendLine("      <td><span class='val'>$db</span></td>")
                [void]$sb.AppendLine("      <td class='td-mb'>$szFmt</td>")
                [void]$sb.AppendLine("      <td style='text-align:center'>$arcIco</td>")
                [void]$sb.AppendLine("    </tr>")
            }

            [void]$sb.AppendLine("    </tbody></table></div>")
            [void]$sb.AppendLine("</details>")
        }


        # ---- JS -------------------------------------------------------------
        [void]$sb.AppendLine('<script>')
        [void]$sb.AppendLine('(function(){')
        [void]$sb.AppendLine('  function goAnchor(){')
        [void]$sb.AppendLine('    var h=location.hash; if(!h)return;')
        [void]$sb.AppendLine('    var el=document.querySelector(h);')
        [void]$sb.AppendLine("    if(el&&el.tagName==='DETAILS'){")
        [void]$sb.AppendLine('      el.open=true;')
        [void]$sb.AppendLine("      setTimeout(function(){el.scrollIntoView({behavior:'smooth',block:'start'})},80);")
        [void]$sb.AppendLine('    }')
        [void]$sb.AppendLine('  }')
        [void]$sb.AppendLine('  goAnchor();')
        [void]$sb.AppendLine("  window.addEventListener('hashchange',goAnchor);")
        [void]$sb.AppendLine('})();')
        [void]$sb.AppendLine('</script>')
        [void]$sb.Append((Get-ExHtmlFooter -Text 'This file contains mailbox names, aliases and sizes: store and share it accordingly.'))

        $sb.ToString() | Set-Content -Path $OutputPath -Encoding UTF8
    }

    # =========================================================================
    Write-StepBanner -StepName '19' -Title 'Migration batch preparation' -Mode $Mode -Actions @(
        'Collect user mailboxes (User, Shared, Room, Equipment)',
        'Include PublicFolder mailboxes',
        'Include in-place archive mailboxes',
        'Exclude HealthMailbox*',
        'Exclude system mailboxes (Arbitration / AuditLog / AuxAuditLog / Discovery) - handled in Step18',
        'Exclude Monitoring mailboxes (not migrated, removed when the legacy servers are uninstalled)',
        'Exclude mailboxes already on Exchange 2019',
        "Distribute into $batchCount balanced batches (size + count)",
        'Generate MigrationPlan.csv + MigrationPlan.html (Apply/Simulate)'
    )

    Reset-Report
    Initialize-ExchangeShell -Credential $Credential

    $ex2013Names = (Get-LegacyExchangeServers).Name
    $ex2019Names = (Get-Exchange2019Servers).Name

    if (-not $ex2013Names) {
        Write-Log "No Exchange 2013/2016 server detected. Nothing left to migrate." -Level Warning
        Save-Report -StepName 'Step19-PrepareMigration' | Out-Null
        return 0
    }

    # --- Mailbox collection --------------------------------------------------
    Write-Log "Collecting the mailboxes to migrate..." -Level Sub
    $all = New-Object System.Collections.Generic.List[object]

    # System mailboxes (Arbitration / AuditLog / AuxAuditLog / Discovery): handled by Step18
    # They are excluded here through RecipientTypeDetails to avoid duplicates.
    $systemTypes = @('ArbitrationMailbox','AuditLogMailbox','AuxAuditLogMailbox','DiscoveryMailbox')

    $regular = Get-Mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue |
               Where-Object {
                   $_.ServerName -in $ex2013Names -and
                   $_.Name -notlike 'HealthMailbox*' -and
                   $_.RecipientTypeDetails -notin $systemTypes
               }
    foreach ($m in $regular) { $all.Add($m) }

    $pf = Get-Mailbox -PublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue |
          Where-Object { $_.ServerName -in $ex2013Names -and $_.Name -notlike 'HealthMailbox*' }
    foreach ($m in $pf) { $all.Add($m) }

    $unique = $all | Sort-Object -Property @{ Expression = { [string]$_.Guid } } -Unique
    Write-Log "  -> $($unique.Count) primary mailboxes" -Level Sub

    # --- Dataset construction (size via Get-MailboxStatistics) ---------------
    $data = New-Object System.Collections.Generic.List[object]

    foreach ($mbx in $unique) {
        $stats = $null
        try { $stats = Get-MailboxStatistics -Identity ([string]$mbx.Guid) -ErrorAction SilentlyContinue } catch { }

        $sizeMB = if ($stats) { ConvertTo-SizeMB -Size $stats.TotalItemSize } else { 0.00 }

        $data.Add([PSCustomObject]@{
            DisplayName          = [string]$mbx.DisplayName
            PrimarySmtpAddress   = [string]$mbx.PrimarySmtpAddress
            Alias                = [string]$mbx.Alias
            RecipientTypeDetails = [string]$mbx.RecipientTypeDetails
            SourceServer         = [string]$mbx.ServerName
            SourceDatabase       = [string]$mbx.Database
            HasArchive           = [bool]($mbx.ArchiveStatus -eq 'Active' -or $mbx.ArchiveDatabase)
            SizeMB               = $sizeMB
            Type                 = 'Primary'
            Batch                = 0
        })

        if ($mbx.ArchiveDatabase) {
            $archStats = $null
            try { $archStats = Get-MailboxStatistics -Identity ([string]$mbx.Guid) -Archive -ErrorAction SilentlyContinue } catch { }
            $archMB = if ($archStats) { ConvertTo-SizeMB -Size $archStats.TotalItemSize } else { 0.00 }

            $data.Add([PSCustomObject]@{
                DisplayName          = [string]$mbx.DisplayName + ' (Archive)'
                PrimarySmtpAddress   = [string]$mbx.PrimarySmtpAddress
                Alias                = [string]$mbx.Alias
                RecipientTypeDetails = 'Archive'
                SourceServer         = [string]$mbx.ServerName
                SourceDatabase       = [string]$mbx.ArchiveDatabase
                HasArchive           = $true
                SizeMB               = $archMB
                Type                 = 'Archive'
                Batch                = 0
            })
        }
    }

    $total = $data.Count
    Write-Log "Total entries to distribute (primary + archive): $total" -Level Sub

    if ($total -eq 0) {
        Write-Log "No mailbox to migrate." -Level Warning
        Save-Report -StepName 'Step19-PrepareMigration' | Out-Null
        return 0
    }

    # --- LPT distribution + cap ----------------------------------------------
    $countCap = [math]::Ceiling($total / $batchCount)
    $batches  = @{}
    for ($i = 1; $i -le $batchCount; $i++) {
        $batches[$i] = [PSCustomObject]@{ Number = $i; Count = 0; SizeMB = 0.0 }
    }

    $sorted = $data | Sort-Object -Property SizeMB -Descending
    foreach ($mbx in $sorted) {
        $candidates = $batches.Values | Where-Object { $_.Count -lt $countCap }
        if (-not $candidates) { $candidates = $batches.Values }
        $target = $candidates | Sort-Object SizeMB, Count | Select-Object -First 1
        $mbx.Batch      = $target.Number
        $target.Count  += 1
        $target.SizeMB += $mbx.SizeMB
    }

    # --- Per-batch summary (all modes) ---------------------------------------
    foreach ($b in ($batches.Values | Sort-Object Number)) {
        Add-Report -Step '19-Plan' -Target ("Batch{0:D2}" -f $b.Number) -Action 'Distribution' `
                   -Status Inventoried `
                   -BeforeValue '' `
                   -AfterValue ("Mbx={0} | Size={1:N0} MB" -f $b.Count, $b.SizeMB) `
                   -Detail ''
    }

    # --- Inventory mode: display only, no export -----------------------------
    if ($Mode -eq 'Inventory') {
        Add-Report -Step '19-Plan' -Target 'GLOBAL' -Action 'Distribution simulation' -Status Inventoried `
                   -AfterValue "$total entries across $batchCount batches (plan not generated in Inventory mode)" `
                   -Detail ''
        Save-Report -StepName 'Step19-PrepareMigration' | Out-Null
        return 0
    }

    # --- CSV export ----------------------------------------------------------
    $planPath = Join-Path $OutputFolder 'MigrationPlan.csv'
    $data | Sort-Object Batch, DisplayName |
        Export-Csv -Path $planPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Log "Migration plan CSV: $planPath" -Level Success

    Add-Report -Step '19-Plan' -Target 'GLOBAL' -Action 'Export-Csv' -Status Success `
               -AfterValue $planPath -Detail "$total entries distributed across $batchCount batches"

    # --- HTML export ---------------------------------------------------------
    $htmlPath = Join-Path $OutputFolder 'MigrationPlan.html'
    New-MigrationPlanHtml -Data $data -Batches $batches -OutputPath $htmlPath -ReportMode $Mode -BatchCount $batchCount
    Write-Log "Migration HTML report: $htmlPath" -Level Success

    Add-Report -Step '19-Plan' -Target 'GLOBAL' -Action 'Export-Html' -Status Success `
               -AfterValue $htmlPath -Detail 'Visual report of the distribution per batch'

    Save-Report -StepName 'Step19-PrepareMigration' | Out-Null
    return 0
}
