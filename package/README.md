# Exchange 2013/2016 to 2019 Migration

Deploys Exchange Server 2019 next to the legacy servers, moves every mailbox in controlled batches and proves the cut-over.

This folder contains everything needed to run the tool: `Deploy-Exchange2019.ps1`, `Manage-IISLogs.ps1`, the step scripts and their module, the configuration, the HealthChecker script and the guide. Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows. Unblock them once, from this folder:
>
> ```powershell
> Get-ChildItem . -Recurse -File | Unblock-File
> ```

## Requirements
- Windows Server 2019 or 2022.
- Windows PowerShell 5.1 - Exchange Management Shell 2019, local or through WinRM.
- Exchange Organization Management and local administrator rights on the Exchange 2019 servers.
- Active Directory rights for Steps 07 and 15.
- WinRM (5985) to the Exchange 2019 servers.

## Quick start
```powershell
notepad .\Configs\Deployment.config.psd1                     # URLs, domain, product key, databases...
notepad .\Configs\DiskLayout.csv ; notepad .\Configs\DAGInfo.csv

.\Deploy-Exchange2019.ps1 -Step List                         # the 26 steps
.\Deploy-Exchange2019.ps1 -Step All -Mode Inventory          # current state, changes nothing
.\Deploy-Exchange2019.ps1 -Step 1,2,3 -Mode Simulate         # preview
.\Deploy-Exchange2019.ps1 -Step 1,2,3 -Mode Apply            # apply (asks YES)
.\Deploy-Exchange2019.ps1 -Step 20 -Follow -FollowInterval 15   # follow the migration batches
```

## Content
| Item | Role |
|---|---|
| `Deploy-Exchange2019.ps1` | Entry script and orchestrator. |
| `Manage-IISLogs.ps1` | Standalone IIS log compression and purge script. |
| `Configs\` | Configuration, CSV examples and HealthChecker script. |
| `Docs\` | User and developer guides, Markdown and single-file HTML, with images. |
| `Modules\` | Runtime module and manifest. |
| `Steps\` | The 26 migration step scripts. |
| `LICENSE` | MIT licence. |
| `README.md` | This short package readme. |

## Documentation
- [User guide](Docs/Exchange2019Migration-UserGuide.md) - also `Docs/Exchange2019Migration-UserGuide.html`, a single file to open locally
- [Developer guide](Docs/Exchange2019Migration-Guide.md) - also `Docs/Exchange2019Migration-Guide.html`

Project page, releases and change log: https://github.com/Nico77600/Exchange2013-2016-to-2019-Migration

License: [MIT](LICENSE).