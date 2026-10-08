---
title: Exchange 2013/2016 to 2019 Migration
subtitle: User guide
version: 2.0.0
author: Nicolas Fabert
updated: 2026-10-08
requires: Windows PowerShell 5.1
scope: Exchange 2019 only, legacy protected
---

# Exchange 2013/2016 to 2019 Migration — User guide

> What you need before the first command, then the runbook you follow from one end to the other: **what is the state of my servers?**, **build the Exchange 2019 platform**, **move the system mailboxes**, **plan and create the batches**, **follow the moves**, **complete a batch tonight**, **clean up and set the production quotas**, **prove that nobody uses the legacy servers any more**. Each step gives the command to copy and what to check afterwards. The background, every configuration key, the 26 steps in detail, the internals and the filtering rules are in the [developer guide](Exchange2019Migration-Guide.md).

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows and fail to run. Before using this project, unblock every file in the downloaded folder:
>
> ```powershell
> Get-ChildItem "C:\Chemin\Du\Dossier" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you downloaded or extracted this project.

```cards
checklist | Prerequisites | Chapter 1: the shell, the account, the rights, where to run the framework.
download | One-time setup | Chapter 2: copy the folder, fill in the configuration, read the current state.
terminal | Everyday use | Chapters 3 to 5: one command per step, the platform, then the migration runbook.
chart | Results | Chapter 6: the reports, the statuses and the exit codes.
```

# Part I · Start here

<!-- icon: checklist -->
## 1. Prerequisites

| Item | Requirement |
|---|---|
| Operating system | Windows Server 2019 or 2022 |
| PowerShell | **Windows PowerShell 5.1** — the version of the Exchange Management Shell. Not PowerShell 7. |
| Exchange | The **Exchange 2019 Management Shell**: local snap-in, or remoting to at least one Exchange 2019 server (`http://<server>/PowerShell/`, Kerberos) |
| Active Directory | RSAT module `ActiveDirectory` (Step 15) |
| Account | Member of **Organization Management**, and **local administrator** on the four Exchange 2019 servers |
| Network | **WinRM 5985/tcp** open to the four Exchange 2019 servers |
| Console | **Windows Terminal** recommended (emoji icons). The classic console works with symbol icons. |
| Browser | Edge, Chrome or Firefox, to open the HTML reports |

### 1.1 Where to run it

```cards
gear | On an Exchange 2019 server | The simplest case: the snap-in is loaded automatically.
user | On an administration workstation | With the Exchange 2019 management tools installed.
key | From any PowerShell 5.1 | With WinRM access to an Exchange 2019 server: `-ExchangeServer exch201901.contoso.local -Credential (Get-Credential)`.
```

> [!WARNING]
> **Remoting goes further than the Exchange session.** Several steps open their own connection to **each** Exchange 2019 server: disks (Step 04), restart of `MSExchangeIS` (Steps 01, 10, 11), transport queue (05), IIS and log paths (06, 12, 17), event logs (14), log analysis (26). Check that the account has the rights on **every** server, not only on the one used for the Exchange session.

### 1.2 Rights needed by some steps

| Step | Needs |
|---|---|
| 02 · Certificate | Export rights on the legacy certificate (exportable private key), write access to `CertificateExportPath` |
| 15 · Kerberos ASA | RSAT `ActiveDirectory`; **Create Computer Objects** on the target OU, or an ASA account pre-created by the AD team; `RollAlternateServiceAccountPassword.ps1` from `$env:ExchangeInstallPath\Scripts` — run Step 15 where Exchange 2019 or its management tools are installed |
| 24 · HealthChecker | Local administrator on each Exchange 2019 server |
| 26 · Protocol logs | Read access to the IIS and SMTP log folders of every Exchange server, legacy included |

<!-- icon: download -->
## 2. One-time setup

```steps
Copy the folder | Copy the package to the server or the workstation, for example `C:\Scripts\Deploy-Exchange2019`. No installer.
Unblock the files | `Get-ChildItem C:\Scripts\Deploy-Exchange2019 -Recurse -File -Force | Unblock-File` (files copied from the Internet or a share).
Edit the configuration | `Configs\Deployment.config.psd1`, `DiskLayout.csv` and `DAGInfo.csv`: URLs, names, disks, DAG, databases — the environment values are empty in the package ([developer guide, chapter 6](Exchange2019Migration-Guide.md#6-configuration)).
Check | `.\Deploy-Exchange2019.ps1 -Step List`, then `-Step All -Mode Inventory`: the current state, without any change.
```

```powershell
git clone https://github.com/Nico77600/Exchange2013-2016-to-2019-Migration.git Deploy-Exchange2019
cd Deploy-Exchange2019\package
notepad .\Configs\Deployment.config.psd1                     # URLs, domain, product key, databases...
notepad .\Configs\DiskLayout.csv ; notepad .\Configs\DAGInfo.csv

.\Deploy-Exchange2019.ps1 -Step List                         # the 26 steps
.\Deploy-Exchange2019.ps1 -Step All -Mode Inventory          # current state, changes nothing
```

**You should see:** the catalogue of the 26 steps, then an Inventory run that writes one CSV per step and a run report under `Reports\<run>\`. A key left empty or with its placeholder (`XXXXX-XXXXX-...`) stops the step that needs it with the return code `2`.

![Console — step catalogue](images/console-list.png)

# Part II · Everyday use

<!-- icon: terminal -->
## 3. One command, three modes

Open **Windows PowerShell 5.1** as administrator — preferably in Windows Terminal — on an Exchange 2019 server or on the admin workstation, go to the framework folder, and run one command.

| I want… | Command |
|---|---|
| See the 26 steps | `.\Deploy-Exchange2019.ps1 -Step List` |
| Read the state of everything | `.\Deploy-Exchange2019.ps1 -Step All -Mode Inventory` |
| Preview the preparation | `.\Deploy-Exchange2019.ps1 -Step (1..9) -Mode Simulate` |
| Apply one step | `.\Deploy-Exchange2019.ps1 -Step 3 -Mode Apply` |
| Apply several steps | `.\Deploy-Exchange2019.ps1 -Step 10,11 -Mode Apply` |
| Run a step by its name | `.\Deploy-Exchange2019.ps1 -Step PrepareMigration -Mode Simulate` |
| Continue after an incident | `.\Deploy-Exchange2019.ps1 -Step All -Mode Apply -Resume` |
| Work from an admin workstation | add `-ExchangeServer exch201901.contoso.local -Credential (Get-Credential)` |
| Read the full help | `Get-Help .\Deploy-Exchange2019.ps1 -Full` |

**Inventory** reads, **Simulate** previews the same actions with `-WhatIf`, **Apply** changes the servers and asks for a typed confirmation first (`-Force` skips it). Write the step numbers without a leading zero: `-Step 3`, not `-Step 03`.

> [!IMPORTANT]
> **Apply changes production servers.** Always run an Inventory or a Simulate of the same steps first. `-Step All -Mode Apply` never runs the migration steps **18 to 26** — they are called one by one, in the order of chapter 5.

> [!TIP]
> **Interrupted?** <kbd>Ctrl</kbd>+<kbd>C</kbd> stops the run; the step in progress stays `InProgress` in `DeploymentState.json`. Every action is idempotent: run the same command again, or add `-Resume`.

<!-- icon: layers -->
## 4. Build the platform — Steps 01 to 17

Steps 01 to 17 prepare the Exchange 2019 servers. Apply them **phase by phase**, in a maintenance window, before the 2019 servers carry users.

```steps
Preview | `.\Deploy-Exchange2019.ps1 -Step (1..9) -Mode Simulate` — read the console and the CSV files: every line must be what you expect.
Phases 1 and 2 | `-Step (1..9) -Mode Apply` — system preparation and the DAG; then `-Step 10,11 -Mode Apply` — databases and copies. Step 04 needs a **reboot** when it configures the page file.
Phase 3 | `-Step (12..17) -Mode Apply` — runtime services: `X-Forwarded-For`, anti-malware check, event log size, Kerberos ASA, MAPI over HTTP, IIS log maintenance.
Check | `.\Deploy-Exchange2019.ps1 -Step All -Mode Inventory` — everything already compliant shows `DONE`.
```

> [!WARNING]
> Several steps restart services or need a reboot: Step 01 (`MSExchangeIS`), Step 03 (IIS reset, one server at a time), Step 04 (page file — **reboot required**), Steps 05 and 06 (transport services), Steps 10 and 11 (`MSExchangeIS`). Only Steps **15** and **25** change the legacy servers, and they say so in red.

![Console — a step in Inventory mode](images/console-run.png)

<!-- icon: refresh -->
## 5. The migration runbook — Steps 18 to 26

The migration steps are called **one at a time**, by an operator who reads the reports between two commands. This is the tested sequence.

```flow
people | Step 18 | System mailboxes
arrow | |
layers | Step 19 | Batch plan
arrow | |
play | Step 20 | Batches created and synced
arrow | Follow |
check | Step 21 | Completion
arrow | |
refresh | Steps 22-24 | Cleanup, quotas, health
arrow | |
search | Steps 26 · 25 · 26 | Legacy traffic, SCP, final check
```

### 5.1 Before the migration — Steps 18 and 19

```powershell
.\Deploy-Exchange2019.ps1 -Step 18 -Mode Apply    # system mailboxes
.\Deploy-Exchange2019.ps1 -Step 19 -Mode Apply    # the batch plan
```

**Check after Step 18:** open `Reports\SystemMailboxPlan.html` — the arbitration, audit and discovery mailboxes move to 2019 and complete on their own. They carry DLP, OAB generation and eDiscovery: they must be on 2019 before the users.

**Check after Step 19:** open `Reports\MigrationPlan.html` — expected scope? balanced batches? archives attached to their mailbox? If needed, change `Migration.BatchCount` and run Step 19 again, or move rows between batches in `MigrationPlan.csv`. A total of **0 archives** when archives are expected usually means the `ArchiveDatabase` property is empty on the legacy mailboxes.

### 5.2 Launch and follow — Step 20

```powershell
.\Deploy-Exchange2019.ps1 -Step 20 -Mode Apply                  # create the batches
.\Deploy-Exchange2019.ps1 -Step 20 -Follow -FollowInterval 15   # follow, in a second window
```

One migration batch per plan batch, started and **suspended before completion**; they are visible in the EAC, under *Migration*. The follow mode submits nothing: it regenerates `MigrationStatus.html` every `-FollowInterval` minutes and opens it in the browser.

**Check:** the tiles of the status report — **Quarantined** > 0 → inspect now · **Failed** > 0 → find the cause (often `BadItemLimit`) · **Synced** > 0 → a batch is ready to complete.

> [!CAUTION]
> **Never run Step 20 in Apply while a migration is in progress**: it removes every batch and move request and starts again from zero. During the migration, use `-Follow` only; to restart a single mailbox, work on its move request with the Exchange cmdlets.

A move request can be stalled or **quarantined** by the Mailbox Replication Service while its status stays `Synced` or `Syncing`:

```powershell
Get-MoveRequestStatistics john.smith@contoso.com -IncludeReport | Format-List
Resume-MoveRequest john.smith@contoso.com
```

`StalledDueToTarget_DiskLatency` means the target is saturated: watch the DAG and the disks before resuming more moves.

### 5.3 Completion, batch by batch — Step 21

| Need | Command |
|---|---|
| See the state of every batch | `.\Deploy-Exchange2019.ps1 -Step 21` (Inventory) |
| Complete one batch now | `.\Deploy-Exchange2019.ps1 -Step 21 -Mode Apply -BatchNumber 02 -Force` |
| Complete several batches | `… -BatchNumber '06','07','08'` |
| Complete every `Synced` batch | `… -BatchNumber ALL` |
| Complete tonight, console closed | `… -BatchNumber ALL -ScheduledCompletionTime '2026-05-15 23:00'` — **recommended for night windows** |
| Complete and watch | `… -BatchNumber 01 -FollowAfter -FollowInterval 2` |

**Check:** `Reports\MigrationBatchStatus.csv` (Step 21 without `-BatchNumber`) gives `Status`, `TotalCount`, `SyncedCount`, `FinalizedCount` and `FailedCount` per batch. After a **scheduled** completion the batch stays `Synced` in the EAC even when every mailbox is done — the real progress is in the move requests, and Step 22 cleans the batch up.

> [!WARNING]
> With `-FollowAfter`, the console stays busy until <kbd>Ctrl</kbd>+<kbd>C</kbd> or the final status. To close the console and let Exchange finish alone, use `-ScheduledCompletionTime` without `-FollowAfter`.

### 5.4 After the migration — Steps 22 to 24

```steps
Check the failures | Read the `FailedCount` column of `Reports\MigrationBatchStatus.csv` and fix each failed mailbox.
Clean up | `-Step 22` (Inventory, to see what would be removed), then `.\Deploy-Exchange2019.ps1 -Step 22 -Mode Apply -Force`. `$env:CLEANUP_INCLUDE_FAILED = '1'` also removes the failed batches.
Production quotas | `.\Deploy-Exchange2019.ps1 -Step 23 -Mode Apply -Force` — the high migration quotas existed so that no move was blocked.
Health | `.\Deploy-Exchange2019.ps1 -Step 24 -Mode Apply -Force`, then read `HealthChecker-Issues.html`.
```

> [!TIP]
> Run Step 16 again after the migration: it enables MAPI over HTTP and blocks RPC over HTTP for the mailboxes that are now on 2019.

### 5.5 Prove the cut-over — Steps 26, 25, 26

```powershell
.\Deploy-Exchange2019.ps1 -Step 26                       # remaining legacy traffic
.\Deploy-Exchange2019.ps1 -Step 25 -Mode Apply -Force    # empty the legacy Autodiscover SCP
$env:LOG_HOURS = 168 ; .\Deploy-Exchange2019.ps1 -Step 26   # final check, one week of logs
```

Step 26 reads the IIS and SMTP logs of every server, legacy and 2019, and counts the traffic of **real users** only — probes, system mailboxes and anonymous requests are filtered out. `ProtocolLogsAnalysis.html` opens by itself.

**Check:** the `Autodiscover` row is the key indicator. Legacy > 0 means clients still reach a legacy SCP: run Step 25, then Step 26 again a few hours later. **Legacy = 0 for IIS and SMTP** over a representative window (a full working week) means the legacy servers can be decommissioned.

> [!NOTE]
> A connector with protocol logging `None` gives no SMTP data: run `-Step 26 -Mode Apply` once (it sets **Verbose**), then read the analysis again after a few hours. Clean the environment variables after use: `Remove-Item Env:BATCH, Env:SCHEDULE_TIME, Env:FOLLOW, Env:FOLLOW_INTERVAL, Env:CLEANUP_INCLUDE_FAILED, Env:LOG_HOURS -ErrorAction SilentlyContinue`.

<!-- icon: chart -->
## 6. Results

Everything the framework writes goes under `Reports\`. Plans and status files shared between runs stay at the root; each run gets its own timestamped folder with one CSV per step, a global CSV, the run report, the run log and one transcript per step.

| Report | Written by | Content |
|---|---|---|
| `Report_GLOBAL_<Mode>.html` | every run | Run report: tiles per status, one accordion per step — open when it has a failure |
| `SystemMailboxPlan.html` | Step 18 | The system mailboxes, their source server and database, their move request |
| `MigrationPlan.html` · `MigrationPlan.csv` | Step 19 | Primaries, archives, batches, total GB, and the mailboxes of each batch |
| `MigrationStatus*.html` | Step 20 | Follow report, refreshes itself: tiles, batches, mailboxes |
| `MigrationBatchStatus.csv` | Step 21 | One row per batch, with its counters |
| `HealthChecker-Issues.html` | Step 24 | Warnings and errors only, per server and category |
| `ProtocolLogsAnalysis.html` | Step 26 | Legacy versus 2019 traffic, per virtual directory, connector and server |

![Run report](images/report-run.png)

All the reports are **self-contained** HTML files with a light and a dark theme: they can be mailed or archived as they are.

| Console status | Report status | Meaning |
|---|---|---|
| `OK` | `Success` | Change applied and checked |
| `DONE` | `AlreadyDone` | Already compliant, nothing to do |
| `INV` | `Inventoried` | Current state read |
| `SIM` | `Simulated` | Change previewed, not applied |
| `SKIP` | `Skipped` | Nothing to do or not applicable |
| `FAILED` | `Failed` | Error — the message is in the `ErrorMessage` column |

| Exit code | Meaning |
|---|---|
| `0` | All selected steps completed — also returned when the operator declines the Apply confirmation |
| `1` | At least one step failed |
| `2` | Every step completed, but some actions failed: the final card is yellow — open the run report |
| `10` | Module `Modules\Exchange2019.Common.psd1` missing |
| `11` | Unknown step number or name |
| `12` · `13` · `14` | A step-specific parameter without its step: `-PfxPassword` (step 2), `-Follow` or `-FollowInterval` (step 20), `-FollowAfter` (step 21) |

# Part III · Troubleshoot

<!-- icon: lifebuoy -->
## 7. If something goes wrong

```steps
Read | Open the CSV of the step (or the run report) and read `ErrorMessage` and `Detail` of the `FAILED` lines; the transcript of the step has the full output.
Fix | Fix the cause — configuration, rights, prerequisite step — not the symptom.
Check | Run the step again in **Inventory**, then **Simulate**: the lines must now be what you expect.
Apply | Run the step in Apply, alone or with `-Resume`. Everything already done is `DONE`.
```

| Symptom | Cause and what to do |
|---|---|
| Exit `11` | Unknown value in `-Step` (a typo, or `03` written as a name): run `-Step List` and use numbers without a leading zero, or the exact names. |
| Exit `12`, `13` or `14` | A step-specific parameter without its step: add the step to `-Step`, or remove the parameter. |
| A step returns `2` | Configuration file missing, key missing or placeholder left: fix `Deployment.config.psd1`, then run the step in Inventory. |
| Cmdlets not found, access denied over WinRM | No Exchange shell and no remote session: run on a 2019 server, or pass `-ExchangeServer` and `-Credential` (1.1). |
| Squares or `?` instead of icons | Console without emoji support: `$env:EXM_ICONS = 'Ascii'` (or `Symbols`), or use Windows Terminal. |
| Step 21 completes the wrong batch, Step 20 refreshes at an odd interval | A leftover `$env:BATCH`, `$env:SCHEDULE_TIME` or `$env:FOLLOW_INTERVAL` wins over the parameters: remove them (5.5). |
| Step 20 returns `1` | `MigrationPlan.csv` missing (Step 19 not run in Simulate or Apply), or no 2019 database: run Step 19, run Step 10. |
| Every batch disappeared | Step 20 was run again in Apply during the migration: rebuild with Step 20 Apply, then use `-Follow` only (5.2). |
| Tile **Quarantined** > 0 | MRS stalled or quarantined the move: `Get-MoveRequestStatistics <id> -IncludeReport`, then `Resume-MoveRequest` or `Set-MoveRequest -SkipAutomaticResume:$false`. |
| Mailbox `Failed` | Most often too many corrupted items: read the report of the move request, then `Set-MoveRequest <id> -BadItemLimit <n>` and `Resume-MoveRequest <id>`. |
| `SyncedCount` = 0 on a `Syncing` batch | Exchange updates this counter only when the whole batch is `Synced`: read the status report, which counts per mailbox. |
| Step 21: *no eligible batch* | The selected batches are not `Synced` yet: wait for `Synced`, or follow them with `-FollowAfter`. |
| Batch still `Synced` in the EAC after completion | Normal after a scheduled completion: it works on the move requests, not on the batch. Step 22 removes it. |
| Step 25: Autodiscover legacy > 0 in Step 26 | Clients still read the legacy SCP, or a DNS record or a hard-coded URL points to legacy: run Step 25, check internal DNS, Outlook profiles and scripts, then run Step 26 again hours later. |
| Step 26: no SMTP data for a server | Protocol logging `None` on its connectors: run Step 26 in Apply, then Inventory after a few hours. |

Anything else — every step in detail, every configuration key and the full troubleshooting tables: [developer guide, Annex A](Exchange2019Migration-Guide.md#annex-a--troubleshooting) and [chapter 8](Exchange2019Migration-Guide.md#8-the-26-steps).
