<#
.SYNOPSIS
    Step 02 - Export the certificate from Exchange 2013/2016 and import it on the 2019 servers.

.DESCRIPTION
    Exports the Exchange certificate from the legacy servers, imports it on every Exchange 2019
    server and enables the required services.

    Logic:
      - Walks through ALL Exchange 2013/2016 Mailbox servers (not Edge)
      - Identifies the non-self-signed certificate that has ALL the required services
        (IIS, SMTP, POP, IMAP - read from Config.CertificateServices)
      - Keeps the certificate with the latest expiration date
      - Exports it to a PFX (password protected) if not already done
      - Imports it on each Exchange 2019 server and enables the services
      - Idempotent: checked through the Thumbprint on each target

    Technical note: Get-ExchangeCertificate through a PSRemoting proxy returns objects
    with Services/IsSelfSigned = null (double serialization). The Get-CertSafe helper
    runs the command INSIDE the Exchange session (before serialization) and converts
    Services and IsSelfSigned to simple types (string / bool).

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

    $exportPath       = $Config.CertificateExportPath
    $services         = $Config.CertificateServices
    $requiredServices = ($services -split '[,\s]+') | Where-Object { $_ }

    Write-StepBanner -StepName '02' -Title 'Certificate export and import' -Mode $Mode -Actions @(
        ('MANUAL mode: if the PFX already exists at {0}, the step reads the thumbprint from the PFX' -f $exportPath),
        '             (skip 2013 discovery + skip export; use case: non-exportable key on the 2013)',
        'AUTO mode: otherwise, walks through the Exchange 2013/2016 Mailbox servers to find the certificate,',
        ('           identifies the non-self-signed cert with the {0} services, then Export-ExchangeCertificate' -f $services),
        'Import of the PFX on each Exchange 2019 server (PrivateKeyExportable=True)',
        ('Activation for the services: {0}' -f $services),
        'Idempotency via Thumbprint (skip if already installed and enabled)'
    )

    Initialize-ExchangeShell -Credential $Credential

    # No dedicated helper: the global Exchange proxies are used.
    # WARNING: the Exchange 2013 PS endpoints are in NoLanguage mode -> Invoke-Command
    # with a scriptblock is rejected. The global proxies serialize Services/IsSelfSigned
    # as null. The filter below handles both cases:
    #   - Services available : filter on the required services (normal behavior)
    #   - Services = null    : fallback filter on Subject (excludes the Exchange self-signed
    #                          certs whose Subject matches the server name)

    if (-not $exportPath -or -not $services) {
        Write-Log "Incomplete certificate configuration (CertificateExportPath / CertificateServices)." -Level Error
        Add-Report -Step '02-Certificate' -Target '<config>' -Action 'CheckConfig' -Status Failed `
                   -ErrorMessage 'CertificateExportPath or CertificateServices missing in Deployment.config.psd1'
        return 2
    }

    # ----- 1. PFX password --------------------------------------------------
    # Possible sources (in priority order):
    #   1. $Global:PfxPassword : passed with -PfxPassword on the orchestrator (CI, Posh-ACME)
    #   2. Read-Host : interactive prompt (Apply mode only)
    #   3. Dummy SecureString : Simulate mode without a provided password (the cmdlets
    #      are called with -WhatIf, the password is never actually used)
    #   4. $null : Inventory mode (no action executed)
    $certPassword = $null
    $passwordSource = 'none'
    if ($Global:PfxPassword) {
        $certPassword = $Global:PfxPassword
        $passwordSource = 'CLI (-PfxPassword)'
        Write-Log "PFX password provided on the CLI - interactive prompt bypassed." -Level Sub
    }
    elseif ($Mode -eq 'Apply') {
        if ([Console]::IsInputRedirected -or $Host.Name -notmatch 'ConsoleHost') {
            throw "Apply mode requires the PFX password interactively: rerun in an interactive PowerShell console, or pass -PfxPassword on the CLI."
        }
        $certPassword = Read-Host "PFX certificate password" -AsSecureString
        $passwordSource = 'Read-Host'
    }
    elseif ($Mode -eq 'Simulate') {
        $certPassword = ConvertTo-SecureString -String 'SimulateOnlyDummyPassword' -AsPlainText -Force
        $passwordSource = 'Simulate-dummy'
    }
    $isWhatIf = ([bool]($Mode -eq 'Simulate'))

    # ----- 2. PFX provisioning mode: MANUAL vs AUTO ---------------------------
    # MANUAL: if the PFX is already present at $exportPath, the step:
    #   - reads the thumbprint and the private key directly from the PFX (Apply only)
    #   - skips the certificate discovery on the Exchange 2013/2016 servers
    #   - skips Export-ExchangeCertificate (marked AlreadyDone)
    # Use case: non-exportable private key (PrivateKeyExportable=False) on the 2013,
    # renewal with a new certificate, or simply a cache of a previous run.
    #
    # AUTO: otherwise, walks through the Exchange 2013/2016 Mailbox servers, identifies
    # the non-self-signed certificate with all the required services, then exports it to a PFX.
    $thumbprint   = $null
    $sourceCert   = $null
    $sourceServer = $null
    $manualMode   = (Test-Path $exportPath)

    if ($manualMode) {
        # ----- 2a. MANUAL mode: validation of the pre-positioned PFX ---------
        $sourceServer = '<manual-pfx>'
        Write-Log "MANUAL mode: PFX already present at $exportPath - discovery + export skipped." -Level Info

        # Validation is only possible with a real password:
        #   - Apply (Read-Host)
        #   - Apply or Simulate with -PfxPassword passed on the CLI
        $canValidate = $passwordSource -in @('CLI (-PfxPassword)', 'Read-Host')
        if ($canValidate) {
            # Validation: open the PFX with the password + check the private key.
            # Extract the thumbprint from the PFX (authoritative source).
            try {
                $plainPwd = [System.Net.NetworkCredential]::new('', $certPassword).Password
                $pfxCert  = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 `
                                -ArgumentList $exportPath, $plainPwd, 'Exportable'
                if (-not $pfxCert.HasPrivateKey) {
                    $pfxCert.Dispose()
                    throw 'The PFX does not contain a private key (HasPrivateKey=False)'
                }
                $thumbprint = $pfxCert.Thumbprint
                $pfxSubject = $pfxCert.Subject
                $pfxNotAft  = $pfxCert.NotAfter
                $pfxCert.Dispose()

                Add-Report -Step '02-Certificate' -Target $sourceServer -Action 'ProvidePfxManually' -Status Inventoried `
                           -Phase 'Before' `
                           -BeforeValue ("Thumbprint={0}; Subject={1}; NotAfter={2}; HasPrivateKey=True; PasswordSource={3}" -f `
                               $thumbprint, $pfxSubject, $pfxNotAft, $passwordSource) `
                           -Detail "Source: pre-positioned PFX ($exportPath)"
            } catch {
                $msg = "Validation of the manual PFX failed: $($_.Exception.Message). Check the password (source=$passwordSource) and that the file really contains the private key."
                Write-Log $msg -Level Error
                Add-Report -Step '02-Certificate' -Target $sourceServer -Action 'ProvidePfxManually' -Status Failed `
                           -ErrorMessage $msg -Detail "PFX: $exportPath"
                return 4
            }
        }
        else {
            # Inventory or Simulate without -PfxPassword: no usable password.
            # The presence of the PFX is reported. The per-server actions do not change anything in these modes.
            $hint = if ($Mode -eq 'Inventory') {
                'validation impossible in Inventory - no password (pass -PfxPassword to validate)'
            } else {
                'validation skipped in Simulate - dummy password (pass -PfxPassword to validate)'
            }
            Add-Report -Step '02-Certificate' -Target $sourceServer -Action 'ProvidePfxManually' -Status Inventoried `
                       -Detail "Manual PFX detected: $exportPath ($hint)"
        }

        Add-Report -Step '02-Certificate' -Target $sourceServer -Action 'Export-ExchangeCertificate' -Status AlreadyDone `
                   -Detail "Source: manual pre-positioned PFX ($exportPath)"
    }
    else {
        # ----- 2b. AUTO mode: discovery + export from Exchange 2013/2016 ---

        # Discovery of the Exchange 2013/2016 Mailbox servers
        $serversLegacy = Get-LegacyExchangeServers | Where-Object { $_.ServerRole -notlike '*Edge*' }

        if (-not $serversLegacy) {
            $msg = "No Exchange 2013/2016 Mailbox server detected. To provide a manual PFX, place the file at: $exportPath"
            Write-Log $msg -Level Warning
            Add-Report -Step '02-Certificate' -Target '<discovery>' -Action 'Select2013Server' -Status Skipped `
                       -Detail $msg
            return 0
        }

        $legacyNames = ($serversLegacy | ForEach-Object { $_.Name }) -join ', '
        Add-Report -Step '02-Certificate' -Target '<discovery>' -Action 'Select2013Server' -Status Inventoried `
                   -Detail "Available 2013/2016 servers: $legacyNames"

        # Search for the certificate on ALL legacy servers
        foreach ($srv in $serversLegacy) {
            $srvName = $srv.Name
            try {
                $candidates = Get-ExchangeCertificate -Server $srvName -ErrorAction Stop |
                              Where-Object {
                                  # IsSelfSigned can be null through the Exchange 2013 proxy (NoLanguage)
                                  if ($_.IsSelfSigned -eq $true) { return $false }

                                  $subj   = [string]$_.Subject
                                  $issuer = [string]$_.Issuer

                                  # Subject == Issuer => self-signed (WMSvc, Auth, default, Azure)
                                  if ($subj -and $issuer -and $subj -eq $issuer) { return $false }

                                  $svc = [string]$_.Services
                                  if ($svc) {
                                      # Services available: check the required services
                                      foreach ($r in $requiredServices) {
                                          if ($svc -notmatch $r) { return $false }
                                      }
                                  } else {
                                      # Services = null: additional exclusion patterns
                                      if ($subj -match "(?i)^CN=$([regex]::Escape($srvName))[\s,\.\$]?" -or
                                          $subj -match '(?i)DC=Windows Azure' -or
                                          $subj -match '(?i)CN=WMSvc' -or
                                          $subj -match '(?i)CN=Microsoft Exchange Server Auth') {
                                          return $false
                                      }
                                  }
                                  return $true
                              }
                if ($candidates) {
                    $best = $candidates | Sort-Object NotAfter -Descending | Select-Object -First 1
                    if (-not $sourceCert -or $best.NotAfter -gt $sourceCert.NotAfter) {
                        $sourceCert   = $best
                        $sourceServer = $srvName
                    }
                }
            } catch {
                Write-Log "Unable to query $srvName for the certificate: $($_.Exception.Message)" -Level Warning
            }
        }

        if (-not $sourceCert) {
            $msg = "No non-self-signed certificate with the services '$services' found on: $legacyNames. To provide a manual PFX, place the file at: $exportPath"
            Write-Log $msg -Level Error
            Add-Report -Step '02-Certificate' -Target '<discovery>' -Action 'IdentifySourceCert' -Status Failed `
                       -ErrorMessage $msg
            return 3
        }

        Write-Log "Certificate found on $sourceServer (Thumbprint=$($sourceCert.Thumbprint))" -Level Info
        Add-Report -Step '02-Certificate' -Target $sourceServer -Action 'IdentifySourceCert' -Status Inventoried `
                   -Phase 'Before' -BeforeValue ("Thumbprint={0}; Subject={1}; Services={2}; NotAfter={3}" -f `
                       $sourceCert.Thumbprint, $sourceCert.Subject, [string]$sourceCert.Services, $sourceCert.NotAfter)

        $thumbprint = $sourceCert.Thumbprint

        # ----- 3. Export the PFX from the source 2013/2016 --------------------
        # Note: $manualMode=$false here, so the PFX does not exist at $exportPath.
        # If Export-ExchangeCertificate fails (typically a non-exportable key),
        # the user can manually pre-position the PFX at $exportPath
        # and rerun: the step then automatically switches to MANUAL mode.
        Invoke-Action -Step '02-Certificate' -Target $sourceServer -Action 'Export-ExchangeCertificate' `
            -Detail "Export to $exportPath" `
            -InventoryScript { "PFX not found - will be generated: $exportPath" } `
            -PreCheckScript  { Test-Path $exportPath } `
            -ActionScript    {
                if ($isWhatIf) {
                    Write-Log "WhatIf: Export-ExchangeCertificate $thumbprint -> $exportPath" -Level Sub
                    return
                }
                $bin = Export-ExchangeCertificate -Server $sourceServer -Thumbprint $thumbprint `
                            -BinaryEncoded -Password $certPassword -ErrorAction Stop
                [System.IO.File]::WriteAllBytes($exportPath, $bin.FileData)
            }
    }

    # ----- 4. Import and enable on each 2019 server -------------------------
    $servers = @(Get-Exchange2019Servers)

    # Safeguard: in MANUAL Simulate, the thumbprint could not be extracted from the PFX
    # (dummy password). The per-server actions require -Thumbprint (mandatory
    # parameter) - skip cleanly instead of letting the parameter binding fail.
    if (-not $thumbprint -and $Mode -eq 'Simulate') {
        foreach ($srv in $servers) {
            Add-Report -Step '02-Certificate' -Target $srv.Name -Action 'Import-ExchangeCertificate' -Status Skipped `
                       -Detail 'Thumbprint unknown in MANUAL Simulate (PFX validation impossible without a real password). Apply will perform the actual import.'
            Add-Report -Step '02-Certificate' -Target $srv.Name -Action 'Enable-ExchangeCertificate' -Status Skipped `
                       -Detail 'Thumbprint unknown in MANUAL Simulate. Apply will perform the actual activation.'
        }
        Save-Report -StepName 'Step02-Certificate'
        return 0
    }

    foreach ($srv in $servers) {
        $server = $srv.Name

        Invoke-Action -Step '02-Certificate' -Target $server -Action 'Import-ExchangeCertificate' -Detail $thumbprint `
            -InventoryScript {
                $existing = Get-ExchangeCertificate -Server $server -ErrorAction SilentlyContinue
                if ($existing) {
                    ($existing | ForEach-Object { "TP={0}; Svc={1}; NotAfter={2}" -f $_.Thumbprint, [string]$_.Services, $_.NotAfter }) -join ' | '
                } else {
                    'No certificate returned'
                }
            } `
            -PreCheckScript {
                $existing = Get-ExchangeCertificate -Server $server -ErrorAction SilentlyContinue |
                            Where-Object { $_.Thumbprint -eq $thumbprint }
                return [bool]$existing
            } `
            -ActionScript {
                if ($isWhatIf) {
                    Write-Log "WhatIf: Import-ExchangeCertificate $thumbprint on $server" -Level Sub
                    return
                }
                $bytes = [System.IO.File]::ReadAllBytes($exportPath)
                Import-ExchangeCertificate -Server $server -FileData $bytes -Password $certPassword `
                    -PrivateKeyExportable $true -ErrorAction Stop | Out-Null
            }

        Invoke-Action -Step '02-Certificate' -Target $server -Action 'Enable-ExchangeCertificate' -Detail "Services=$services" `
            -InventoryScript {
                $c = Get-ExchangeCertificate -Server $server -ErrorAction SilentlyContinue |
                     Where-Object { $_.Thumbprint -eq $thumbprint }
                if ($c) { "Services=$([string]$c.Services)" } else { '<certificate not found>' }
            } `
            -PreCheckScript {
                $existing = Get-ExchangeCertificate -Server $server -ErrorAction SilentlyContinue |
                            Where-Object { $_.Thumbprint -eq $thumbprint }
                if (-not $existing) { return $false }
                # Services can be null through the Exchange proxy -> impossible to verify
                # Enable-ExchangeCertificate is left to run to guarantee the activation
                $svc = [string]$existing.Services
                if (-not $svc) { return $false }
                foreach ($r in $requiredServices) {
                    if ($svc -notmatch $r) { return $false }
                }
                return $true
            } `
            -ActionScript {
                Enable-ExchangeCertificate -Server $server -Thumbprint $thumbprint -Services $services `
                    -WhatIf:$isWhatIf -ErrorAction Stop -Confirm:$false
            }
    }

    Save-Report -StepName 'Step02-Certificate'
    return 0
}
