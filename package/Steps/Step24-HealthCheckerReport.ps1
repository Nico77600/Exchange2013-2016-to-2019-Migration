<#
.SYNOPSIS
    Step 24 - Run Microsoft HealthChecker and build a custom Warnings/Errors report.

.DESCRIPTION
    Runs the Microsoft HealthChecker on all deployed Exchange 2019 servers, then
    builds a CUSTOM HTML report focused on Warnings and Errors only.

    The official report (HealthChecker -BuildHtmlServersReport) remains
    available to review the OK items.

    Modes:
      - Inventory: checks that the HealthChecker.ps1 script is present and lists
                   the target Exchange 2019 servers. Also lists the HealthChecker
                   XML files present in the output folder.
      - Simulate : displays the exact command that would be run.
      - Apply    : runs HealthChecker.ps1, parses the XML files produced, keeps
                   only the lines with WriteType=Red / Yellow, then builds a
                   custom HTML report:
                      1. Summary per server (number of errors / warnings / categories)
                      2. Detail per server, grouped by category

    ManualOnly step: never run automatically in "All".
    Idempotent in the diagnostic sense: running it again = regenerating the report.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    Part of : Exchange 2013/2016 to 2019 Migration - Deploy-Exchange2019.ps1
#>

# ---------------------------------------------------------------------------
# Step-local helpers
# ---------------------------------------------------------------------------

function _HtmlEscape {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).
        Replace('&','&amp;').
        Replace('<','&lt;').
        Replace('>','&gt;').
        Replace('"','&quot;')
}

function Get-HealthCheckerIssueLines {
    <#
        Parses the object imported by Import-Clixml from a HealthChecker XML
        (output of HealthChecker.ps1, hashtable containing DisplayResults and
        HealthCheckerExchangeServer).

        Returns only the lines considered as Warning or Error:
            WriteType = 'Red'    -> Severity 'Error'
            WriteType = 'Yellow' -> Severity 'Warning'
            WriteType = 'OutColumns' -> the cells are inspected; if at least
            one cell has DisplayColor Red/Yellow, the whole table is returned.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $AnalyzedResult
    )

    $issues = New-Object System.Collections.Generic.List[object]
    $displayResults = $AnalyzedResult.DisplayResults
    if (-not $displayResults) { return ,@() }

    foreach ($keyEntry in $displayResults.GetEnumerator()) {
        $catKey   = $keyEntry.Key
        $catName  = if ($catKey.Name) { [string]$catKey.Name } else { 'Other' }
        $catOrder = if ($null -ne $catKey.DisplayOrder) { [int]$catKey.DisplayOrder } else { 9999 }

        foreach ($line in $keyEntry.Value) {
            $writeType = $line.WriteType
            if ($writeType -eq 'OutColumns') {
                $hasRed = $false; $hasYellow = $false
                $tv = $line.TestingValue
                if ($tv) {
                    foreach ($row in $tv) {
                        foreach ($p in $row.PSObject.Properties) {
                            $cell = $p.Value
                            if ($cell -and $cell.DisplayColor) {
                                if ($cell.DisplayColor -eq 'Red')    { $hasRed = $true }
                                elseif ($cell.DisplayColor -eq 'Yellow') { $hasYellow = $true }
                            }
                        }
                    }
                }
                if (-not ($hasRed -or $hasYellow)) { continue }

                # Rebuild a tabular text for the report
                $textRows = @()
                foreach ($row in $tv) {
                    $cells = @()
                    foreach ($p in $row.PSObject.Properties) {
                        $val = $null
                        if ($p.Value -and ($p.Value.PSObject.Properties.Name -contains 'Value')) {
                            $val = $p.Value.Value
                        } else {
                            $val = $p.Value
                        }
                        $cells += ('{0}={1}' -f $p.Name, $val)
                    }
                    $textRows += ($cells -join '; ')
                }
                $severity = if ($hasRed) { 'Error' } else { 'Warning' }
                $issues.Add([PSCustomObject]@{
                    CategoryName  = $catName
                    CategoryOrder = $catOrder
                    Severity      = $severity
                    Name          = '(table)'
                    DisplayValue  = ($textRows -join "`n")
                }) | Out-Null
            }
            elseif ($writeType -eq 'Red' -or $writeType -eq 'Yellow') {
                $severity = if ($writeType -eq 'Red') { 'Error' } else { 'Warning' }
                $name = if ([string]::IsNullOrEmpty($line.Name)) { '' } else { [string]$line.Name }
                $val  = if ($null -eq $line.DisplayValue) { '' } else { [string]$line.DisplayValue }
                $issues.Add([PSCustomObject]@{
                    CategoryName  = $catName
                    CategoryOrder = $catOrder
                    Severity      = $severity
                    Name          = $name
                    DisplayValue  = $val
                }) | Out-Null
            }
        }
    }
    # Return as object[] (PS 5.1 + List<object> + @() = "Argument types do not match")
    return ,$issues.ToArray()
}

function New-HealthCheckerCustomHtmlReport {
    <#
        Builds a custom HTML report focused on Warnings/Errors.
        Uses the shared report theme of the module (Get-ExHtmlHead / Get-ExHtmlHero):
        header with clickable tiles, filter bar, server list and one accordion per server.

        ServerIssues : hashtable ServerName -> array of issues (PSCustomObject)
        Path         : HTML output path
        OfficialHtmlPath : (optional) path to the official HealthChecker report
                          to display as a link in the header.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable]$ServerIssues,
        [Parameter(Mandatory)] [string]$Path,
        [string]$OfficialHtmlPath = '',
        [string]$XmlDirectory     = ''
    )

    $now = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    # Pre-compute totals
    $totalErr = 0; $totalWarn = 0; $serversOk = 0; $serversKo = 0
    $perSrv = @{}
    foreach ($srv in $ServerIssues.Keys) {
        $items = $ServerIssues[$srv]
        $e = @($items | Where-Object Severity -eq 'Error').Count
        $w = @($items | Where-Object Severity -eq 'Warning').Count
        $totalErr += $e; $totalWarn += $w
        if ($e -eq 0 -and $w -eq 0) { $serversOk++ } else { $serversKo++ }
        $perSrv[$srv] = @{ Err = $e; Warn = $w; Items = $items }
    }
    $totalServers = $ServerIssues.Keys.Count

    $sb = [System.Text.StringBuilder]::new(120000)

    # ---- HEAD and HEADER: shared report theme of the module ----------------
    # Filters and tables of this report, on top of the shared theme.
    $extraStyle = @"
.clk{cursor:pointer}.tile.active,.hdr-pill.active{box-shadow:0 0 0 2px var(--cp-accent)}
.fbar{position:relative;display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:16px}
.fbar-lbl{font-size:11px;color:var(--cp-accent);font-weight:700;letter-spacing:.08em;text-transform:uppercase}
.fbtn{border-radius:999px;padding:5px 14px;font-size:12px;font-weight:600}
.fbtn.active{background:var(--cp-accent);border-color:var(--cp-accent);color:var(--cp-accent-fg)}
.fbtn.fb-err.active{background:var(--cp-danger);border-color:var(--cp-danger)}.fbtn.fb-warn.active{background:var(--cp-warning);border-color:var(--cp-warning)}
.fbar-help{font-size:12px;color:var(--cp-text-muted);margin-left:auto}
body.f-err tbody tr.warn{display:none}body.f-warn tbody tr.err{display:none}
body.f-err .cat-sec:not(:has(tr.err)){display:none}body.f-warn .cat-sec:not(:has(tr.warn)){display:none}
body.f-err details.srv[data-err="0"]{display:none}body.f-warn details.srv[data-warn="0"]{display:none}
body.f-err .toc-item[data-err="0"]{display:none}body.f-warn .toc-item[data-warn="0"]{display:none}
.filter-empty{display:none;padding:16px 20px;color:var(--cp-text-muted);text-align:center}
body.f-err .filter-empty.for-err,body.f-warn .filter-empty.for-warn{display:block}
.notice{position:relative;margin-top:16px;padding:10px 14px;border:1px solid var(--cp-border);border-left:4px solid var(--cp-warning);border-radius:.625rem;background:var(--cp-surface);font-size:12px;line-height:1.7}
.toc-cnt{font-size:12px;font-weight:700;white-space:nowrap}.toc-cnt.ko{color:var(--cp-danger)}.toc-cnt.wn{color:var(--cp-warning)}.toc-cnt.ok{color:var(--cp-success)}
.pill.err,.bdg.err{background:var(--cp-danger)}.pill.warn,.bdg.warn{background:var(--cp-warning)}.pill.ok{background:var(--cp-success)}
.cat-sec{border-top:1px solid var(--cp-border);padding:12px 16px 4px}.cat-sec:first-of-type{border-top:none}
.cat-hdr{display:flex;align-items:center;gap:8px;margin-bottom:8px}.cat-name{font-weight:600}.cat-meta{font-size:12px;color:var(--cp-text-muted);margin-left:auto}
.cat-sec .tbl-wrap{border:1px solid var(--cp-border);border-radius:.625rem;margin-bottom:8px}
th.w-st{min-width:80px}th.w-nm{min-width:180px;width:28%}
tbody tr.err td:first-child{box-shadow:inset 3px 0 0 var(--cp-danger)}tbody tr.warn td:first-child{box-shadow:inset 3px 0 0 var(--cp-warning)}
.nm{font-weight:600;word-break:break-word}.nm.empty{color:var(--cp-text-muted);font-style:italic;font-weight:400}
.cat-sec .val{font-family:Consolas,"Cascadia Mono",monospace;font-size:12px;line-height:1.45}
.empty-srv{padding:18px 16px;text-align:center;color:var(--cp-success);font-weight:600}.empty-srv .ic{font-size:24px;display:block;margin-bottom:4px}
"@
    [void]$sb.Append((Get-ExHtmlHead -Title 'Exchange 2019 migration | HealthChecker warnings and errors' -ExtraStyle $extraStyle))

    # Clickable tiles: same markup as Get-ExHtmlTile, with the filter they apply.
    $tile = {
        param($Class, $Filter, $Label, $Value, $Note, $Tone)
        $data  = if ($Filter) { " data-filter='$Filter'" } else { '' }
        $style = if ($Tone) { " style='--tone:$Tone'" } else { '' }
        "<article class='$Class'$data$style><span class='tile-label'>$Label</span><strong>$Value</strong><small>$Note</small></article>"
    }
    $tiles = @(
        (& $tile 'tile tile-primary clk' 'err' 'Errors' (Format-ExNumber $totalErr) 'Click to show the errors only' '')
        (& $tile 'tile clk' 'warn' 'Warnings' (Format-ExNumber $totalWarn) 'Click to show the warnings only' 'var(--cp-warning)')
        (& $tile 'tile tile-tinted clk' 'all' 'Servers analysed' (Format-ExNumber $totalServers) 'Click to show everything' '')
        (& $tile 'tile' '' 'Servers without issue' ("$serversOk / $totalServers") ("$serversKo to review") $(if ($serversKo) { 'var(--cp-warning)' } else { 'var(--cp-success)' }))
    )
    $extra = "<div class='fbar'><span class='fbar-lbl'>Show</span>" +
             "<button class='fbtn fb-all active' data-filter='all'>All</button>" +
             "<button class='fbtn fb-err' data-filter='err'>Errors only</button>" +
             "<button class='fbtn fb-warn' data-filter='warn'>Warnings only</button>" +
             "<span class='fbar-help'>Click a tile or a button; click again to show everything.</span></div>"
    if ($XmlDirectory -or $OfficialHtmlPath) {
        $extra += "<div class='notice'>"
        if ($XmlDirectory) { $extra += "<b>HealthChecker XML:</b> <code>$(_HtmlEscape $XmlDirectory)</code><br>" }
        if ($OfficialHtmlPath) {
            $relOff = Split-Path $OfficialHtmlPath -Leaf
            $extra += "<b>Official HealthChecker report (all statuses):</b> <a href='./$(_HtmlEscape $relOff)'>$(_HtmlEscape $relOff)</a> &mdash; open it to see the OK items"
        }
        $extra += '</div>'
    }
    $box = "Analysed $now<br>$totalServers Exchange 2019 server(s)"
    [void]$sb.Append((Get-ExHtmlHero -Title 'HealthChecker: warnings and errors' -Subtitle 'Results of the Microsoft Exchange HealthChecker on every Exchange 2019 server (Step 24), limited to the warnings and errors. The OK items are in the official HealthChecker report.' `
        -BoxLabel 'Analysis' -BoxHtml $box -Tiles $tiles -ExtraHtml $extra))

    # ---- TABLE OF CONTENTS + TOOLBAR ---------------------------------------
    [void]$sb.AppendLine("<div class='card'>")
    [void]$sb.AppendLine("  <div class='card-hdr'>Servers &mdash; $totalServers</div>")
    [void]$sb.AppendLine("  <div class='toolbar'>")
    [void]$sb.AppendLine("    <button class='btn' onclick=""document.querySelectorAll('details.srv').forEach(function(d){d.open=true})"">&#9660;&thinsp;Expand all</button>")
    [void]$sb.AppendLine("    <button class='btn' onclick=""document.querySelectorAll('details.srv').forEach(function(d){d.open=false})"">&#9650;&thinsp;Collapse all</button>")
    [void]$sb.AppendLine("  </div>")
    [void]$sb.AppendLine("  <div class='toc'>")
    foreach ($srv in ($ServerIssues.Keys | Sort-Object)) {
        $aid  = 'srv-' + ($srv -replace '[^a-zA-Z0-9]','-')
        $e    = $perSrv[$srv].Err
        $w    = $perSrv[$srv].Warn
        $lbl  = _HtmlEscape $srv
        $ind  = if ($e -gt 0) {
                    "<span class='toc-cnt ko'>&#9888;&thinsp;$e err</span>"
                } elseif ($w -gt 0) {
                    "<span class='toc-cnt wn'>&#9888;&thinsp;$w warn</span>"
                } else {
                    "<span class='toc-cnt ok'>&#10003; No issues</span>"
                }
        [void]$sb.AppendLine("    <div class='toc-item' data-err='$e' data-warn='$w'><a class='toc-name' href='#$aid'>$lbl</a>$ind</div>")
    }
    [void]$sb.AppendLine("  </div>")
    [void]$sb.AppendLine("</div>")

    # ---- ACCORDIONS PER SERVER ---------------------------------------------
    foreach ($srv in ($ServerIssues.Keys | Sort-Object)) {
        $aid   = 'srv-' + ($srv -replace '[^a-zA-Z0-9]','-')
        $items = $perSrv[$srv].Items
        $e     = $perSrv[$srv].Err
        $w     = $perSrv[$srv].Warn
        $openA = if ($e -gt 0 -or $w -gt 0) { ' open' } else { '' }
        $lbl   = _HtmlEscape $srv

        $pills = ''
        if ($e -gt 0) { $pills += "<span class='pill err'>$e&nbsp;err</span>" }
        if ($w -gt 0) { $pills += "<span class='pill warn'>$w&nbsp;warn</span>" }
        if ($e -eq 0 -and $w -eq 0) { $pills = "<span class='pill ok'>No issues</span>" }

        $catCount = if ($items.Count -gt 0) {
            ($items | Select-Object -ExpandProperty CategoryName -Unique).Count
        } else { 0 }
        $cntLabel = if ($catCount -gt 0) { "$catCount category(ies)" } else { '0 categories' }

        [void]$sb.AppendLine("<details class='srv'$openA id='$aid' data-err='$e' data-warn='$w'>")
        [void]$sb.AppendLine("  <summary>")
        [void]$sb.AppendLine("    <span class='acc-icon'>&#9654;</span>")
        [void]$sb.AppendLine("    <span class='acc-name'>$lbl</span>")
        [void]$sb.AppendLine("    <span class='acc-cnt'>$cntLabel</span>")
        [void]$sb.AppendLine("    <span class='acc-pills'>$pills</span>")
        [void]$sb.AppendLine("  </summary>")

        if (-not $items -or $items.Count -eq 0) {
            [void]$sb.AppendLine("  <div class='empty-srv'><span class='ic'>&#10003;</span>No warnings or errors reported by HealthChecker.</div>")
            [void]$sb.AppendLine("</details>")
            continue
        }

        $catGroups = $items | Group-Object CategoryName | Sort-Object {
            ($_.Group | Select-Object -First 1).CategoryOrder
        }
        foreach ($g in $catGroups) {
            $cErr  = @($g.Group | Where-Object Severity -eq 'Error').Count
            $cWarn = @($g.Group | Where-Object Severity -eq 'Warning').Count
            $catPills = ''
            if ($cErr -gt 0)  { $catPills += "<span class='pill err'>$cErr</span>" }
            if ($cWarn -gt 0) { $catPills += "<span class='pill warn'>$cWarn</span>" }

            [void]$sb.AppendLine("  <div class='cat-sec'>")
            [void]$sb.AppendLine("    <div class='cat-hdr'>")
            [void]$sb.AppendLine("      <span class='cat-name'>$(_HtmlEscape $g.Name)</span>")
            [void]$sb.AppendLine("      <span class='cat-meta'>$($g.Group.Count) entry(ies)</span>")
            [void]$sb.AppendLine("      <span class='acc-pills'>$catPills</span>")
            [void]$sb.AppendLine("    </div>")
            [void]$sb.AppendLine("    <div class='tbl-wrap'><table>")
            [void]$sb.AppendLine("      <thead><tr><th class='w-st'>Status</th><th class='w-nm'>Name</th><th>Value</th></tr></thead><tbody>")
            $sorted = $g.Group | Sort-Object Severity, Name
            foreach ($it in $sorted) {
                $cls   = if ($it.Severity -eq 'Error') { 'err' } else { 'warn' }
                $name  = if ([string]::IsNullOrEmpty($it.Name)) {
                            "<span class='nm empty'>(no name)</span>"
                         } else {
                            "<span class='nm'>$(_HtmlEscape $it.Name)</span>"
                         }
                $val   = _HtmlEscape $it.DisplayValue
                [void]$sb.AppendLine("      <tr class='$cls'><td><span class='bdg $cls'>$($it.Severity)</span></td><td>$name</td><td><div class='val'>$val</div></td></tr>")
            }
            [void]$sb.AppendLine("    </tbody></table></div>")
            [void]$sb.AppendLine("  </div>")
        }

        [void]$sb.AppendLine("</details>")
    }

    # "Nothing to show" messages specific to the active filter
    [void]$sb.AppendLine("<div class='card filter-empty for-err'>No <b>Error</b> to display. Switch to <i>Warnings only</i> or <i>All</i> to see the rest.</div>")
    [void]$sb.AppendLine("<div class='card filter-empty for-warn'>No <b>Warning</b> to display. Switch to <i>Errors only</i> or <i>All</i> to see the rest.</div>")

    # ---- BACK TO TOP + JS --------------------------------------------------
    [void]$sb.AppendLine(@'
<script>
(function(){
  // ---- Errors / Warnings / All filtering --------------------------------
  function applyFilter(mode){
    var b=document.body;
    b.classList.remove('f-err','f-warn');
    if(mode==='err')  b.classList.add('f-err');
    if(mode==='warn') b.classList.add('f-warn');
    // Sync the active state of the controls (toolbar buttons, tiles, header pills)
    document.querySelectorAll('.fbtn').forEach(function(x){
      x.classList.toggle('active', x.getAttribute('data-filter')===mode);
    });
    document.querySelectorAll('.tile.clk').forEach(function(x){
      x.classList.toggle('active', x.getAttribute('data-filter')===mode);
    });
    document.querySelectorAll('.hdr-pill.clk').forEach(function(x){
      x.classList.toggle('active', x.getAttribute('data-filter')===mode);
    });
  }
  document.addEventListener('click',function(ev){
    var t=ev.target.closest('[data-filter]');
    if(!t) return;
    var mode=t.getAttribute('data-filter');
    if(!mode) return;
    // Re-click on the active filter => back to "All" mode
    var b=document.body;
    var current=b.classList.contains('f-err')?'err':b.classList.contains('f-warn')?'warn':'all';
    if(mode===current && mode!=='all') mode='all';
    applyFilter(mode);
  });

  // ---- Anchor (ToC) ------------------------------------------------------
  function goAnchor(){
    var h=location.hash; if(!h)return;
    var el=document.querySelector(h);
    if(el&&el.tagName==='DETAILS'){
      el.open=true;
      setTimeout(function(){el.scrollIntoView({behavior:'smooth',block:'start'})},80);
    }
  }
  goAnchor();
  window.addEventListener('hashchange',goAnchor);
})();
</script>
'@)
    [void]$sb.Append((Get-ExHtmlFooter -Text 'This file contains server names and configuration details: store and share it accordingly.'))

    # UTF-8 BOM (consistent with the other reports)
    $bytes = [System.Text.UTF8Encoding]::new($true).GetBytes($sb.ToString())
    [System.IO.File]::WriteAllBytes($Path, $bytes)
}

# ---------------------------------------------------------------------------
# The step itself
# ---------------------------------------------------------------------------

function Invoke-Step {
    [CmdletBinding()]
    param(
        [string]$Mode,
        [string]$OutputFolder,
        [string]$CsvFolder,
        [hashtable]$Config,
        [System.Management.Automation.PSCredential]$Credential
    )

    # --- Script path resolution --------------------------------------------
    $hcCfg = $Config.HealthCheckerReport
    if (-not $hcCfg) {
        Write-Log "HealthCheckerReport section is missing from the configuration file." -Level Error
        Save-Report -StepName 'Step24-HealthCheckerReport' | Out-Null
        return 1
    }

    $rawPath = [string]$hcCfg.ScriptPath
    if ([System.IO.Path]::IsPathRooted($rawPath)) {
        $scriptPath = $rawPath
    } else {
        # Relative to the project root (parent of the Steps\ folder)
        $projectRoot = Split-Path $PSScriptRoot -Parent
        $scriptPath  = Join-Path $projectRoot $rawPath
    }

    $skipVer  = [bool]$hcCfg.SkipVersionCheck
    $doHtml   = [bool]$hcCfg.BuildOfficialHtml

    # Dedicated subfolder: 1 subfolder PER RUN
    # Structure: <OutputFolder>\Step24-HealthChecker\<yyyyMMdd_HHmmss>\
    # Each run is isolated (XML, txt, official HTML, custom HTML).
    $hcParent = Join-Path $OutputFolder 'Step24-HealthChecker'
    $runId    = Get-Date -Format 'yyyyMMdd_HHmmss'
    $runDir   = Join-Path $hcParent $runId

    Write-StepBanner -StepName '24' -Title 'HealthChecker run + Warnings/Errors report' -Mode $Mode -Actions @(
        "HealthChecker.ps1: $scriptPath",
        "Target:            all Exchange 2019 servers",
        "Output:            $runDir",
        "Custom report:     $runDir\HealthChecker-Issues.html (Warnings/Errors only)",
        ("Official report:   {0}" -f ($(if ($doHtml) { "$runDir\ExchangeAllServersReport-*.html (all statuses)" } else { 'not generated (BuildOfficialHtml = $false)' }))),
        "SkipVersionCheck:  $skipVer"
    )

    Initialize-ExchangeShell -Credential $Credential

    $servers = @(Get-Exchange2019Servers)
    if (-not $servers) {
        Write-Log 'No Exchange 2019 server detected.' -Level Warning
        Save-Report -StepName 'Step24-HealthCheckerReport' | Out-Null
        return 0
    }

    $serverNames = $servers | ForEach-Object { $_.Name }

    if (-not (Test-Path $scriptPath)) {
        Write-Log "HealthChecker.ps1 not found: $scriptPath" -Level Error
        foreach ($srv in $servers) {
            Add-Report -Step '24-HC' -Target $srv.Name -Action 'HealthChecker' -Status Failed `
                       -ErrorMessage "Script not found: $scriptPath"
        }
        Save-Report -StepName 'Step24-HealthCheckerReport' | Out-Null
        return 1
    }

    # ------------------------------------------------------------------
    # Inventory mode: no execution, only list the targets
    # ------------------------------------------------------------------
    if ($Mode -eq 'Inventory') {
        foreach ($srv in $servers) {
            Add-Report -Step '24-HC' -Target $srv.Name -Action 'HealthChecker (target)' `
                       -Status Inventoried -Phase 'Before' `
                       -BeforeValue "Script=$scriptPath - will be run in Apply mode"
        }
        Save-Report -StepName 'Step24-HealthCheckerReport' | Out-Null
        return 0
    }

    # ------------------------------------------------------------------
    # Simulate mode: display the exact command but do not run it
    # ------------------------------------------------------------------
    if ($Mode -eq 'Simulate') {
        $cmd = "& '$scriptPath' -Server $($serverNames -join ',') -OutputFilePath <dir>"
        if ($skipVer) { $cmd += ' -SkipVersionCheck' }
        Write-Log "Simulated command: $cmd" -Level Info
        foreach ($srv in $servers) {
            Add-Report -Step '24-HC' -Target $srv.Name -Action 'HealthChecker (simulate)' `
                       -Status Simulated -Detail $cmd
        }
        Save-Report -StepName 'Step24-HealthCheckerReport' | Out-Null
        return 0
    }

    # ------------------------------------------------------------------
    # Apply mode: actual execution
    # ------------------------------------------------------------------
    if (-not (Test-Path $runDir)) {
        New-Item -Path $runDir -ItemType Directory -Force | Out-Null
    }

    Write-Log "Running HealthChecker on $($serverNames.Count) Exchange 2019 server(s)..." -Level Step
    Write-Log "  Output: $runDir" -Level Info

    $hcArgs = @{
        Server         = $serverNames
        OutputFilePath = $runDir
    }
    if ($skipVer) { $hcArgs['SkipVersionCheck'] = $true }

    $hcSuccess = $false
    try {
        & $scriptPath @hcArgs
        $hcSuccess = $true
    } catch {
        Write-Log "Error while running HealthChecker.ps1: $($_.Exception.Message)" -Level Error
    }

    # Optional generation of the official HealthChecker report (HTML, all statuses)
    $officialHtml = ''
    if ($hcSuccess -and $doHtml) {
        try {
            Write-Log "Generating the official HealthChecker report (HTML, all statuses)..." -Level Sub
            $htmlName = 'ExchangeAllServersReport.html'
            $htmlArgs = @{
                BuildHtmlServersReport = $true
                XMLDirectoryPath       = $runDir
                HtmlReportFile         = (Join-Path $runDir $htmlName)
            }
            if ($skipVer) { $htmlArgs['SkipVersionCheck'] = $true }
            & $scriptPath @htmlArgs
            $officialHtml = Join-Path $runDir $htmlName
        } catch {
            Write-Log "Official report generation failed: $($_.Exception.Message)" -Level Warning
        }
    }

    # ------------------------------------------------------------------
    # Parsing the XML files produced and building the custom report
    # ------------------------------------------------------------------
    # The run has its own subfolder => read all the XML files found in it.
    $xmls = @(Get-ChildItem -Path $runDir -Filter 'HealthChecker-*.xml' -File -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -notlike 'HealthChecker-ScriptDebugObject*' -and `
                             $_.Name -notlike 'HealthChecker-VulnerabilityReport*' })

    if (-not $xmls -or $xmls.Count -eq 0) {
        Write-Log "No HealthChecker XML found in $runDir." -Level Error
        foreach ($srv in $servers) {
            Add-Report -Step '24-HC' -Target $srv.Name -Action 'HealthChecker' -Status Failed `
                       -ErrorMessage 'No XML produced'
        }
        Save-Report -StepName 'Step24-HealthCheckerReport' | Out-Null
        return 1
    }

    $serverIssues = @{}
    foreach ($xml in $xmls) {
        try {
            $obj = Import-Clixml -Path $xml.FullName -ErrorAction Stop
            $srvName = $null
            if ($obj.HealthCheckerExchangeServer -and $obj.HealthCheckerExchangeServer.ServerName) {
                $srvName = [string]$obj.HealthCheckerExchangeServer.ServerName
            } else {
                # Fallback: extract from the file name "HealthChecker-<SERVER>-<stamp>.xml"
                $base = [System.IO.Path]::GetFileNameWithoutExtension($xml.Name)
                $parts = $base -split '-'
                if ($parts.Count -ge 3) { $srvName = $parts[1] }
            }
            if (-not $srvName) { $srvName = $xml.BaseName }

            # Normalization: FQDN -> short name (unique key aligned with Get-Exchange2019Servers)
            $srvName = ($srvName -split '\.')[0]

            $issues = Get-HealthCheckerIssueLines -AnalyzedResult $obj
            $serverIssues[$srvName] = $issues

            $err  = @($issues | Where-Object Severity -eq 'Error').Count
            $warn = @($issues | Where-Object Severity -eq 'Warning').Count
            $cats = ($issues | Select-Object -ExpandProperty CategoryName -Unique | Sort-Object) -join ', '

            # The analysis completed => Success. The Err/Warn counts are in AfterValue (the custom
            # HTML report is the main tool to investigate them).
            $afterVal = "Err=$err | Warn=$warn"
            if ($cats) { $afterVal += " | Cats=$cats" }

            Add-Report -Step '24-HC' -Target $srvName -Action 'HealthChecker analysis' `
                       -Status Success -Phase 'After' -AfterValue $afterVal -Detail $xml.Name
        } catch {
            Write-Log "Failed to parse $($xml.Name): $($_.Exception.Message)" -Level Error
            Add-Report -Step '24-HC' -Target $xml.BaseName -Action 'HealthChecker parsing' `
                       -Status Failed -ErrorMessage $_.Exception.Message
        }
    }

    # Also include (as OK) the expected servers that have no XML
    foreach ($srv in $serverNames) {
        if (-not $serverIssues.ContainsKey($srv)) {
            $serverIssues[$srv] = @()
            Add-Report -Step '24-HC' -Target $srv -Action 'HealthChecker' -Status Failed `
                       -ErrorMessage 'XML missing for this server (HealthChecker collection probably failed)'
        }
    }

    $customHtml = Join-Path $runDir 'HealthChecker-Issues.html'
    New-HealthCheckerCustomHtmlReport -ServerIssues $serverIssues -Path $customHtml `
                                       -OfficialHtmlPath $officialHtml -XmlDirectory $runDir

    Write-Log "Custom report (Warnings/Errors): $customHtml" -Level Success
    if ($officialHtml) {
        Write-Log "Official report (all statuses): $officialHtml" -Level Success
    }

    Save-Report -StepName 'Step24-HealthCheckerReport' | Out-Null
    return 0
}
