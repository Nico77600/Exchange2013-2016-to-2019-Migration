<#
.SYNOPSIS
    Step 04 - Disk initialization, formatting and assignment for Exchange 2019.

.DESCRIPTION
    Initializes, formats and assigns the Exchange 2019 disks.

    Source: Configs\DiskLayout.csv
    Columns:
        ServerName   : target server name (case-insensitive)
        Role         : Swap | Queue | Databases
        DriveLetter  : target letter (X, Q, M) - without ':'
        DiskNumber   : physical disk number (optional - if set, it takes precedence over the size)
        MinSizeGB    : minimum disk size in GB (for identification by size)
        MaxSizeGB    : maximum disk size in GB (99999 for unlimited)

    Target disk identification:
        - If DiskNumber is set (integer > 0): Get-Disk -Number $DiskNumber
        - Otherwise: search within the MinSizeGB-MaxSizeGB range (first free disk, smallest)

    Formatting rules (Microsoft Exchange 2019 recommendations):
        Swap  (X:) : NTFS, 4,096 bytes - skipped if X: is already formatted as NTFS
        Queue (Q:) : NTFS, 65,536 bytes (64 KB)
        Databases  : ReFS, 65,536 bytes (64 KB), IntegrityStreams disabled

    Drives C: and D: are always excluded.
    No mount points - drive letters only.

    Once M: is formatted:
        M:\Databases\<Prefix><N>\          (e.g. DB01)
        M:\Databases\<Prefix><N>\Logs\

    Idempotent: compares the file system and the allocation unit size before any action.
    All operations are run through Invoke-Command on the target server.

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

    Write-StepBanner -StepName '04' -Title 'Disk initialization and formatting' -Mode $Mode -Actions @(
        'Source: DiskLayout.csv - identification by DiskNumber (takes precedence) OR MinSizeGB/MaxSizeGB range',
        'C: and D: always excluded',
        'Swap  (X:) : NTFS 4 KB - skipped if X: is already NTFS',
        'Queue (Q:) : NTFS 64 KB (MS Exchange recommendation)',
        'Databases (M:) : ReFS 64 KB IntegrityStreams=Off (MS Exchange recommendation)',
        'After formatting M: - creation of M:\Databases\<DB>\Logs\ for each database',
        'PageFile: centralized on X:\pagefile.sys (Init=Max=25% RAM), AutoManaged=Off, removes any pagefile != X: (reboot required)'
    )

    # =========================================================================
    # Interactive prompt: are the disks already prepared?
    # Double confirmation YES/YES -> skip Init/Partition/Format
    # The creation of the M:\Databases\<DB>\Logs\ folders is still performed.
    # In Inventory and Simulate modes: no prompt (the script can run
    # NonInteractive). Formatting is not applied in Simulate (-WhatIf)
    # anyway; in Inventory mode we only take inventory.
    # =========================================================================
    $skipDiskFormat = $false
    $isInteractive  = -not [Console]::IsInputRedirected
    if ($Mode -eq 'Inventory') {
        Write-Host ''
        Write-ExItem Info 'Inventory mode: Init/Partition/Format inventoried as usual (no prompt).' -Icon 'Sub'
        Write-Host ''
    }
    elseif ($Mode -eq 'Simulate' -or -not $isInteractive) {
        Write-Host ''
        Write-ExItem Info 'Non-interactive mode: Init/Partition/Format simulated/executed as usual (no prompt).' -Icon 'Sub'
        Write-Host ''
    }
    else {
        Write-Host ''
        Write-ExItem Warn 'Are the disks already initialized, formatted and mounted?'
        Write-ExItem Dim 'If YES/YES: Init/Partition/Format skipped for all servers' -Indent 9
        Write-ExItem Dim 'The creation of the M:\Databases\<DB>\Logs\ folders is still performed' -Indent 9
        $ans1 = Read-Host '     Answer (YES/NO)'
        if ($ans1 -ieq 'YES') {
            $ans2 = Read-Host '     Are you sure they are ready? Confirmation (YES/NO)'
            if ($ans2 -ieq 'YES') {
                $skipDiskFormat = $true
                Write-ExItem Skip 'Init/Partition/Format SKIPPED - folder creation kept.'
            } else {
                Write-ExItem Info 'Negative confirmation - Init/Partition/Format will be executed.' -Icon 'Sub'
            }
        } else {
            Write-ExItem Info 'Init/Partition/Format will be executed as usual.' -Icon 'Sub'
        }
        Write-Host ''
    }

    Initialize-ExchangeShell -Credential $Credential
    Reset-Report

    $csvPath = Join-Path $CsvFolder 'DiskLayout.csv'
    $layout  = Import-ConfigCsv -Path $csvPath -Delimiter ';' | Where-Object { $_.ServerName }
    if (-not $layout) {
        Write-Log "DiskLayout.csv is empty or not found: $csvPath" -Level Warning
        Save-Report -StepName 'Step04-Disks'
        return 0
    }

    # DB names built from DatabaseLayout (e.g. DB01, DB02, DB03, DB04)
    $dl      = $Config.DatabaseLayout
    $dbNames = @(
        for ($i = [int]$dl.DatabaseStartIndex; $i -lt ([int]$dl.DatabaseStartIndex + [int]$dl.DatabaseCount); $i++) {
            '{0}{1}' -f $dl.DatabasePrefix, $i.ToString("D$([int]$dl.DatabaseDigits)")
        }
    )

    # FS/AUS settings per role (MS recommendations)
    $roleCfg = @{
        Swap      = @{ FS = 'NTFS'; AUS = 4096;  Label = 'Swap';      Desc = 'NTFS 4KB'      }
        Queue     = @{ FS = 'NTFS'; AUS = 65536; Label = 'Queue';     Desc = 'NTFS 64KB'     }
        Databases = @{ FS = 'ReFS'; AUS = 65536; Label = 'Databases'; Desc = 'ReFS 64KB NoIS' }
    }
    $roleOrder = @{ Swap = 0; Queue = 1; Databases = 2 }

    $servers = Get-Exchange2019Servers
    if (-not $servers) {
        Write-Log "No Exchange 2019 server detected." -Level Warning
        Save-Report -StepName 'Step04-Disks'
        return 0
    }

    foreach ($srv in $servers) {
        $serverName = $srv.Name
        $srvRows    = @($layout | Where-Object { $_.ServerName -ieq $serverName })

        if (-not $srvRows) {
            Write-Log "[$serverName] Not found in DiskLayout.csv - ignored." -Level Warning
            Add-Report -Step '04-Disks' -Target $serverName -Action 'CheckCsv' -Status Skipped `
                       -Detail 'No row in DiskLayout.csv'
            continue
        }

        $icArgs = @{ ComputerName = $srv.Fqdn; ErrorAction = 'Stop' }
        if ($Credential) { $icArgs['Credential'] = $Credential }

        # Processing order: Swap -> Queue -> Databases
        $orderedRows = $srvRows | Sort-Object { $roleOrder[$_.Role] }

        if ($skipDiskFormat) {
            foreach ($row in $orderedRows) {
                $letter = ($row.DriveLetter -replace ':','').ToUpper()
                Add-Report -Step '04-Disks' -Target $serverName `
                           -Action ("DiskSetup {0} ({1}:)" -f $row.Role, $letter) `
                           -Status Skipped `
                           -Detail "Disks declared as already prepared by the operator (Init/Partition/Format skipped)"
            }
        }
        else { foreach ($row in $orderedRows) {
            $role     = $row.Role
            $letter   = ($row.DriveLetter -replace ':','').ToUpper()
            $minGB    = [double]$row.MinSizeGB
            $maxGB    = [double]$row.MaxSizeGB
            $diskNum  = 0
            if ($row.PSObject.Properties['DiskNumber'] -and ([string]$row.DiskNumber) -match '^\d+$') {
                $diskNum = [int]$row.DiskNumber
            }
            $rc       = $roleCfg[$role]
            $expFS    = $rc.FS
            $expAUS   = [int]$rc.AUS
            $volLabel = $rc.Label
            $isSwap   = ($role -eq 'Swap')
            $isReFS   = ($expFS -eq 'ReFS')
            $idMode   = if ($diskNum -gt 0) { "Disk#$diskNum" } else { "range ${minGB}-${maxGB} GB" }

            Invoke-Action -Step '04-Disks' -Target $serverName -Action "DiskSetup $role (${letter}:)" `
                -Detail "$($rc.Desc) | $idMode" `
                -InventoryScript {
                    try {
                        Invoke-Command @icArgs -ScriptBlock {
                            param($ltr, $minGB, $maxGB, $diskNum)
                            $excluded = @('C','D')
                            $vol = Get-Volume -DriveLetter $ltr -EA SilentlyContinue
                            if ($vol) {
                                $wmiVol = Get-CimInstance -ClassName Win32_Volume `
                                              -Filter ("DriveLetter='{0}:'" -f $ltr) -EA SilentlyContinue
                                $aus = if ($wmiVol) { [int]$wmiVol.BlockSize } else { 0 }
                                $info = "${ltr}: FS=$($vol.FileSystem) AUS=$aus Size=$([math]::Round($vol.Size/1GB,1))GB"
                                if ($ltr -eq 'X') {
                                    $info += " Pagefile=$(Test-Path 'X:\pagefile.sys')"
                                }
                                return $info
                            }
                            # Identification: by number (takes precedence) or by size range
                            if ($diskNum -gt 0) {
                                $disk = Get-Disk -Number $diskNum -EA SilentlyContinue
                                if (-not $disk) { return ("Disk#{0} not found | {1}: absent" -f $diskNum, $ltr) }
                                $ltrs = (Get-Partition -DiskNumber $disk.Number -EA SilentlyContinue |
                                            Where-Object DriveLetter).DriveLetter
                                if ($ltrs | Where-Object { $excluded -contains $_ }) {
                                    return ("Disk#{0} contains {1}: - excluded" -f $diskNum, ($ltrs -join ','))
                                }
                                return ('Disk#{0} {1}GB {2} (by number) -> will be {3}:' -f `
                                    $disk.Number, [math]::Round($disk.Size/1GB,1), $disk.PartitionStyle, $ltr)
                            }
                            $disk = Get-Disk | Where-Object {
                                ($_.Size / 1GB) -ge $minGB -and ($_.Size / 1GB) -le $maxGB
                            } | Where-Object {
                                $ltrs = (Get-Partition -DiskNumber $_.Number -EA SilentlyContinue |
                                            Where-Object DriveLetter).DriveLetter
                                -not ($ltrs | Where-Object { $excluded -contains $_ })
                            } | Sort-Object Size | Select-Object -First 1
                            if ($disk) {
                                return ('Disk#{0} {1}GB {2} (by size) -> will be {3}:' -f `
                                    $disk.Number, [math]::Round($disk.Size/1GB,1), $disk.PartitionStyle, $ltr)
                            }
                            return ("No disk in {0}-{1}GB | {2}: absent" -f $minGB, $maxGB, $ltr)
                        } -ArgumentList $letter, $minGB, $maxGB, $diskNum
                    } catch { "Inventory error: $($_.Exception.Message)" }
                } `
                -PreCheckScript {
                    try {
                        Invoke-Command @icArgs -ScriptBlock {
                            param($ltr, $expFS, $expAUS, $swapRole)
                            $vol = Get-Volume -DriveLetter $ltr -EA SilentlyContinue
                            if (-not $vol -or $vol.FileSystem -ne $expFS) { return $false }
                            # Swap: an existing NTFS volume is enough (an existing NTFS drive is not reformatted)
                            if ($swapRole) { return $true }
                            $wmiVol = Get-CimInstance -ClassName Win32_Volume `
                                          -Filter ("DriveLetter='{0}:'" -f $ltr) -EA SilentlyContinue
                            $aus = if ($wmiVol) { [int]$wmiVol.BlockSize } else { 0 }
                            return ($aus -eq $expAUS)
                        } -ArgumentList $letter, $expFS, $expAUS, $isSwap
                    } catch { $false }
                } `
                -ActionScript {
                    Invoke-Command @icArgs -ScriptBlock {
                        param($ltr, $minGB, $maxGB, $diskNum, $expFS, $expAUS, $lbl, $useReFS, $isWhatIf)
                        $ErrorActionPreference = 'Stop'
                        $excluded = @('C','D')

                        # Identification: by number (takes precedence) or by size range
                        if ($diskNum -gt 0) {
                            $disk = Get-Disk -Number $diskNum -EA SilentlyContinue
                            if (-not $disk) { throw ("Disk#{0} not found" -f $diskNum) }
                            $ltrs = (Get-Partition -DiskNumber $disk.Number -EA SilentlyContinue |
                                        Where-Object DriveLetter).DriveLetter
                            if ($ltrs | Where-Object { $excluded -contains $_ }) {
                                throw ("Disk#{0} contains {1}: (C/D) - refused for safety" -f $diskNum, ($ltrs -join ','))
                            }
                        } else {
                            $disk = Get-Disk | Where-Object {
                                ($_.Size / 1GB) -ge $minGB -and ($_.Size / 1GB) -le $maxGB
                            } | Where-Object {
                                $ltrs = (Get-Partition -DiskNumber $_.Number -EA SilentlyContinue |
                                            Where-Object DriveLetter).DriveLetter
                                -not ($ltrs | Where-Object { $excluded -contains $_ })
                            } | Sort-Object Size | Select-Object -First 1
                            if (-not $disk) {
                                throw ("No disk available in {0}-{1}GB" -f $minGB, $maxGB)
                            }
                        }

                        if ($isWhatIf) {
                            return ('[WhatIf] Disk#{0} {1}GB -> {2}: {3} AUS={4}' -f `
                                $disk.Number, [math]::Round($disk.Size/1GB,1), $ltr, $expFS, $expAUS)
                        }

                        if ($disk.IsOffline)  { Set-Disk -Number $disk.Number -IsOffline  $false }
                        if ($disk.IsReadOnly) { Set-Disk -Number $disk.Number -IsReadOnly $false }

                        if ($disk.PartitionStyle -eq 'RAW') {
                            Initialize-Disk -Number $disk.Number -PartitionStyle GPT
                        }

                        # Remove non-system partitions
                        Get-Partition -DiskNumber $disk.Number -EA SilentlyContinue |
                            Where-Object { $_.Type -notin @('System','Reserved') } |
                            Remove-Partition -Confirm:$false -EA SilentlyContinue

                        $part = New-Partition -DiskNumber $disk.Number `
                                    -DriveLetter ([char]$ltr) -UseMaximumSize
                        Start-Sleep -Seconds 2

                        $fmtParams = @{
                            Partition          = $part
                            FileSystem         = $expFS
                            NewFileSystemLabel = $lbl
                            AllocationUnitSize = $expAUS
                            Confirm            = $false
                            Force              = $true
                        }
                        if ($useReFS) { $fmtParams['SetIntegrityStreams'] = $false }
                        Format-Volume @fmtParams | Out-Null

                        ('{0}: {1} AUS={2} OK - Disk#{3} {4}GB' -f `
                            $ltr, $expFS, $expAUS, $disk.Number, [math]::Round($disk.Size/1GB,1))
                    } -ArgumentList $letter, $minGB, $maxGB, $diskNum, $expFS, $expAUS, $volLabel, $isReFS, ([bool]($Mode -eq 'Simulate'))
                }
        } }

        # ---------------------------------------------------------------
        # Directory structure M:\Databases\<DB>\Logs\
        # ---------------------------------------------------------------
        if ($srvRows | Where-Object { $_.Role -eq 'Databases' }) {

            $capturedDbNames = $dbNames

            Invoke-Action -Step '04-Disks' -Target $serverName `
                -Action 'CreateDatabaseDirectories (M:)' `
                -Detail ('M:\Databases\{0}..{1}\Logs\' -f $capturedDbNames[0], $capturedDbNames[-1]) `
                -InventoryScript {
                    try {
                        Invoke-Command @icArgs -ScriptBlock {
                            param([string[]]$dbs)
                            $status = foreach ($db in $dbs) {
                                $logPath = 'M:\Databases\' + $db + '\Logs'
                                $dbPath  = 'M:\Databases\' + $db
                                $s = if     (Test-Path $logPath) { 'OK'      }
                                     elseif (Test-Path $dbPath)  { 'NoLogs'  }
                                     else                        { 'Missing' }
                                '{0}={1}' -f $db, $s
                            }
                            $status -join ' | '
                        } -ArgumentList (,$capturedDbNames)
                    } catch { "Error: $($_.Exception.Message)" }
                } `
                -PreCheckScript {
                    try {
                        Invoke-Command @icArgs -ScriptBlock {
                            param([string[]]$dbs)
                            foreach ($db in $dbs) {
                                if (-not (Test-Path ('M:\Databases\' + $db + '\Logs'))) { return $false }
                            }
                            return $true
                        } -ArgumentList (,$capturedDbNames)
                    } catch { $false }
                } `
                -ActionScript {
                    Invoke-Command @icArgs -ScriptBlock {
                        param([string[]]$dbs, [bool]$isWhatIf)
                        $ErrorActionPreference = 'Stop'

                        # Safeguard: detect the excessive comma-wrap bug on a multi-arg
                        # -ArgumentList, or any case where the DB names would be badly serialized.
                        if (-not $dbs -or $dbs.Count -lt 1) {
                            throw "dbs is empty or null - parameter badly serialized"
                        }
                        foreach ($d in $dbs) {
                            if ($d -match '\s') {
                                throw "Invalid DB name (whitespace detected): '$d' - $($dbs.Count) elements received: $($dbs -join ' | ')"
                            }
                        }

                        $paths = foreach ($db in $dbs) { 'M:\Databases\' + $db + '\Logs' }
                        if (-not $isWhatIf) {
                            foreach ($p in $paths) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
                            'Created: ' + ($paths -join ', ')
                        } else {
                            '[WhatIf] Would create: ' + ($paths -join ', ')
                        }
                    } -ArgumentList $capturedDbNames, ([bool]($Mode -eq 'Simulate'))
                }
        }

        # ---------------------------------------------------------------
        # PageFile: moves the Windows pagefile to X:\pagefile.sys
        #   - Disables AutomaticManagedPagefile on Win32_ComputerSystem
        #   - Creates/adjusts Win32_PageFileSetting Name='X:\pagefile.sys'
        #   - InitialSize = MaximumSize = 25% of RAM (MS Exchange 2019 recommendation)
        #   - Removes any other pagefile (notably C:\pagefile.sys)
        # Idempotent: if everything is already correct on X:, the action is Skipped.
        # Note: takes effect at the next server restart.
        # ---------------------------------------------------------------
        if ($srvRows | Where-Object { $_.Role -eq 'Swap' }) {

            Invoke-Action -Step '04-Disks' -Target $serverName `
                -Action 'ConfigurePageFile (X:)' `
                -Detail 'AutoManaged=Off, X:\pagefile.sys Init=Max=25% RAM, removes any pagefile != X: [REBOOT REQUIRED]' `
                -InventoryScript {
                    try {
                        Invoke-Command @icArgs -ScriptBlock {
                            $cs       = Get-CimInstance Win32_ComputerSystem -EA SilentlyContinue
                            if (-not $cs) { return 'Win32_ComputerSystem unavailable' }
                            $ramMB    = [int]($cs.TotalPhysicalMemory / 1MB)
                            $targetMB = [int]($ramMB * 0.25)
                            $auto     = $cs.AutomaticManagedPagefile

                            # Persistent config (registry) - applied at the next reboot
                            $pfCfg    = @(Get-CimInstance Win32_PageFileSetting -EA SilentlyContinue)
                            # Live state - pagefiles actually active on the system
                            $pfRun    = @(Get-CimInstance Win32_PageFileUsage   -EA SilentlyContinue)

                            # Search for the physical pagefile.sys file on all drive letters.
                            # We use [IO.File]::Exists() (Win32 API) because Test-Path does not see the hidden
                            # system files locked by the kernel (pagefile.sys, hiberfil.sys, etc.).
                            $pfFiles  = @(Get-Volume -EA SilentlyContinue | Where-Object { $_.DriveLetter } | ForEach-Object {
                                if ([System.IO.File]::Exists(('{0}:\pagefile.sys' -f $_.DriveLetter))) { [string]$_.DriveLetter }
                            } | Sort-Object)
                            $filesStr = if ($pfFiles.Count -gt 0) { 'pagefile.sys[' + ($pfFiles -join ',') + ':]' } else { 'no pagefile.sys' }

                            $cfgInfo  = if ($pfCfg.Count -gt 0) {
                                ($pfCfg | ForEach-Object { '{0} Init={1}MB Max={2}MB' -f $_.Name, $_.InitialSize, $_.MaximumSize }) -join ' ; '
                            } else { 'none (auto-managed)' }

                            $runInfo  = if ($pfRun.Count -gt 0) {
                                ($pfRun | ForEach-Object { '{0} Allocated={1}MB' -f $_.Name, $_.AllocatedBaseSize }) -join ' ; '
                            } else { 'none active' }

                            # REBOOT PENDING detection: the config differs from the live state.
                            # We compare the set of paths (case-insensitive).
                            # If pfCfg is empty (AutoManaged or blank config), nothing is pending - Windows is in control.
                            $cfgSet   = @($pfCfg | ForEach-Object { $_.Name.ToLower() } | Sort-Object) -join '|'
                            $runSet   = @($pfRun | ForEach-Object { $_.Name.ToLower() } | Sort-Object) -join '|'
                            $pending  = ($pfCfg.Count -gt 0) -and ($cfgSet -ne $runSet)

                            # Tag placed first so it stays visible even if the console truncates at 117 chars
                            $prefix   = if ($pending) {
                                '[REBOOT PENDING live=' + (($pfRun | ForEach-Object { $_.Name }) -join ',') +
                                ' cfg=' + (($pfCfg | ForEach-Object { $_.Name }) -join ',') + '] '
                            } else { '' }

                            "${prefix}RAM=${ramMB}MB Target=${targetMB}MB AutoManaged=$auto | $filesStr | Cfg: $cfgInfo | Live: $runInfo"
                        }
                    } catch { "Pagefile inventory error: $($_.Exception.Message)" }
                } `
                -PreCheckScript {
                    try {
                        Invoke-Command @icArgs -ScriptBlock {
                            $cs = Get-CimInstance Win32_ComputerSystem -EA SilentlyContinue
                            if (-not $cs) { return $false }
                            if ($cs.AutomaticManagedPagefile) { return $false }
                            $ramMB    = [int]($cs.TotalPhysicalMemory / 1MB)
                            $targetMB = [int]($ramMB * 0.25)
                            $pfList   = @(Get-CimInstance Win32_PageFileSetting -EA SilentlyContinue)
                            if ($pfList.Count -ne 1) { return $false }
                            $pf = $pfList[0]
                            if ($pf.Name -ine 'X:\pagefile.sys') { return $false }
                            # tolerance +/-10 MB to absorb WMI rounding
                            if ([math]::Abs([int]$pf.InitialSize - $targetMB) -gt 10) { return $false }
                            if ([math]::Abs([int]$pf.MaximumSize - $targetMB) -gt 10) { return $false }
                            return $true
                        }
                    } catch { $false }
                } `
                -ActionScript {
                    Invoke-Command @icArgs -ScriptBlock {
                        param($isWhatIf)
                        $ErrorActionPreference = 'Stop'

                        $cs = Get-CimInstance Win32_ComputerSystem
                        if (-not $cs) { throw 'Win32_ComputerSystem unavailable' }
                        $ramMB    = [int]($cs.TotalPhysicalMemory / 1MB)
                        $targetMB = [int]($ramMB * 0.25)

                        # Check that the X: volume exists before configuring the pagefile
                        $volX = Get-Volume -DriveLetter X -EA SilentlyContinue
                        if (-not $volX) { throw 'Volume X: not found - the pagefile cannot be placed there' }

                        if ($isWhatIf) {
                            return ('[WhatIf] AutoManaged->False, would remove pagefiles other than X:, X:\pagefile.sys Init={0}MB Max={0}MB (RAM={1}MB)' -f $targetMB, $ramMB)
                        }

                        # 1. Disable automatic pagefile management
                        if ($cs.AutomaticManagedPagefile) {
                            Set-CimInstance -InputObject $cs -Property @{ AutomaticManagedPagefile = $false }
                        }

                        # 2. Remove all pagefile settings other than X:
                        Get-CimInstance Win32_PageFileSetting -EA SilentlyContinue |
                            Where-Object { $_.Name -ine 'X:\pagefile.sys' } |
                            ForEach-Object { Remove-CimInstance -InputObject $_ }

                        # 3. Create or adjust the pagefile on X:
                        $x = Get-CimInstance Win32_PageFileSetting -Filter "Name='X:\\pagefile.sys'" -EA SilentlyContinue
                        if (-not $x) {
                            New-CimInstance -ClassName Win32_PageFileSetting -Property @{
                                Name        = 'X:\pagefile.sys'
                                InitialSize = [uint32]$targetMB
                                MaximumSize = [uint32]$targetMB
                            } | Out-Null
                        } else {
                            Set-CimInstance -InputObject $x -Property @{
                                InitialSize = [uint32]$targetMB
                                MaximumSize = [uint32]$targetMB
                            }
                        }

                        'OK: X:\pagefile.sys Init={0}MB Max={0}MB (RAM={1}MB) - reboot required' -f $targetMB, $ramMB
                    } -ArgumentList ([bool]($Mode -eq 'Simulate'))
                }
        }
    }

    Save-Report -StepName 'Step04-Disks'
    return 0
}
