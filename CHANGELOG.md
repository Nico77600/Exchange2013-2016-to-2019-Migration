# Changelog — Exchange 2013/2016 to 2019 Migration

All notable changes are listed here. Versions follow MAJOR.MINOR.PATCH (see the guide, Annex D).
Author: Nicolas Fabert.

## [2.0.0] — 2026-10-01

Same deployment logic as 1.4, already used in production: this version changes the presentation, the documentation and the delivery, aligned with the sister project Purview DLP Report.

### Added
- **Console** in the Purview DLP Report style: framed title card (version, author, mode, steps, server, configuration, run folder, log), one numbered pill per step with its icon, planned actions and scope, one line per result (`OK`, `DONE`, `INV`, `SIM`, `SKIP`, `FAILED`) with the current value on a dimmed line, a result line per step and a framed summary card (green, yellow or red). Emoji in Windows Terminal, symbols of the classic console fonts elsewhere, ASCII on demand (`EXM_ICONS`); `NO_COLOR`. Colours are console colours, so the transcripts stay clean. Source files are plain ASCII: the icons are built from their code points.
- **Run log** `Reports\<run>\Deploy-Exchange2019.log` (timestamps and levels, no colours), next to the transcripts.
- **Exit codes**: `0` all steps completed, `1` a step failed or is missing, `2` every step completed but some actions failed; `10` to `14` unchanged.
- **Shared HTML theme** for every report (run report, system mailbox plan, migration plan, migration status, HealthChecker issues, protocol-log analysis): header with key figures, tiles, status distribution, light and dark theme, self-contained files. The run report gets a search box and a status filter, and shows the catalogue title of each step.
- Module manifest `Modules\Exchange2019.Common.psd1` (version, exported functions); the step catalogue moves into the module (`Get-DeploymentStepCatalog`) with a phase and an icon per step.
- **Administrator guide** `Docs\Exchange2019Migration-Guide.md` (+ `.html`), in the format of the Purview DLP Report guide: parts, chapters, cards, steps, diagrams, callouts, the 26 steps one by one, the migration runbook, troubleshooting, protocol-log filtering rules, versioning. Built by `tools\Build-Documentation.ps1`; screenshots from synthetic data by `tools\New-DocumentationImages.ps1`.
- `tools\New-MigrationPackage.ps1`: package with the runtime files only, environment values emptied, CSV templates, and a check that no value of the environment remains.
- Pester tests `tests\Exchange2019Migration.Tests.ps1` (offline).

### Changed
- **English** everywhere: comments, comment-based help, console messages, report texts. The translation of the 26 steps was checked token by token: the code is unchanged, only comments and display texts differ. Confirmations are now `YES` / `NO` (`Y` / `N` in Step 17); report phases `Before` / `After`.
- Dates of the reports in ISO format (`yyyy-MM-dd HH:mm:ss`).
- **Step 26 - stricter real-user filtering.** IIS: an authenticated `cs-username` is required (anonymous requests, including anonymous `AMProbe` probes, are no longer counted), more system identities (`IIS APPPOOL\`, `IUSR`, `OABGen`, `Microsoft.Exchange*`, `MSExchange*`, local service accounts) and more probe user agents. SMTP: a session counts only when its `MAIL FROM` is a user address; HealthMailbox, SystemMailbox, FederatedEmail, Migration, OABGen, monitoring, `postmaster` and `mailer-daemon` senders are excluded. The report shows the excluded system sessions separately.

### Removed
- `Docs\Documentation.md`, `Docs\Documentation.html` and `Docs\README.md` (French), replaced by the new guide.

## [1.4] — 2026-05-15
- Step 21: `-ScheduledCompletionTime` hands the timing to Exchange (`CompleteAfter`) instead of a PowerShell wait loop; the step ends as soon as the order is set and the console can be closed.
- Orchestrator: new switch `-FollowAfter` (Step 21 only): after the completion order, continues with Step 20 `-Follow` (Inventory) in the same session to watch Synced → Completing → Completed.
- Step 20 focus mode: with `-FollowAfter -BatchNumber NN`, the Exchange calls are limited to the batch (`Get-MigrationBatch -Identity`, `Get-MoveRequest -BatchName`), about ten times faster; file `MigrationStatus-BatchNN.html`, so that a global and a focused follow-up can run side by side.
- Step 20 report: `StatusDetail` of the move request statistics shown next to the status; MRS-quarantined users (`StalledDueToMRS_*`, `StalledDueToTarget_*`) detected even when their status is Synced (sub-badge, tile, batch pill, action banner). Synced moved from orange to purple. Average sync and finalized percentage shown separately per batch.
- Step 10: `Remove-Mailbox -Monitoring` instead of `Disable-Mailbox` for the HealthMailbox objects (also removes the AD object, recreated by the Health Manager service).

## [1.3] — 2026-05-13
- Step 20: `New-MigrationBatch -Local` + `Start-MigrationBatch` instead of `New-MoveRequest -BatchName`: native batches visible in the admin center; primary and archive of a mailbox in the same batch; public folder mailboxes stay on `New-MoveRequest -PublicFolder`.
- Step 21: `Complete-MigrationBatch` (one command per batch); eligibility = batch status Synced; final snapshot `MigrationBatchStatus.csv`.
- Step 20 report: source of truth = `Get-MigrationBatch`; orphan move requests (Step 18, manual) excluded.

## [1.2] — 2026-05-13
- Step 20: follow-up mode `-Follow` / `-FollowInterval` (also `$env:FOLLOW`); `MigrationStatus.html` refreshes itself; errors (FailureType + message) of failed move requests; only the move requests of Steps 19/20 are shown.
- Step 17: IIS log root detected at run time (WebAdministration), recursion on the `W3SVCx` folders, weekly archive names.

## [1.1] — 2026-05-13
- Step 02: MANUAL mode (PFX placed beforehand) next to the AUTO mode; PFX check (private key, thumbprint).
- New parameter `-PfxPassword` (SecureString) for Step 02: no prompt (unattended runs, Posh-ACME).

## [1.0] — 2026-05-13
- First version: 26 steps (Step01-LicenseKey to Step26-AnalyzeProtocolLogs), modes Inventory / Simulate / Apply, idempotence and resume (`DeploymentState.json`, `-Resume`), strict Exchange 2019 filter (Edge, 2013 and 2016 excluded), disk preparation and page file (Step 04), CSV per step + global CSV + global HTML, target servers discovered in Active Directory.
