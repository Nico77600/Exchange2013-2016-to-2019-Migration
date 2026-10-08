<#
.SYNOPSIS
    Step 15 - Kerberos authentication for load-balanced Client Access servers.

.DESCRIPTION
    Sets up Kerberos authentication for the load-balanced Client Access
    servers (Microsoft procedure) with an Alternate Service Account (ASA).

    Reference:
       https://learn.microsoft.com/en-us/exchange/architecture/client-access/kerberos-auth-for-load-balanced-client-access

    Official procedure, in 7 phases:

      1. Creation of the ASA computer account in AD
         New-ADComputer -Name <name> -AccountPassword <SecureString> -Description ... -Enabled:$true -SamAccountName <name>

      2. Enabling AES encryption on the computer account
         Set-ADComputer <name> -add @{"msDS-SupportedEncryptionTypes"="28"}
         (28 = RC4-HMAC + AES128-CTS-HMAC-SHA1-96 + AES256-CTS-HMAC-SHA1-96)

      3+4. Deployment of the ASA credential on the Exchange 2019 CAS
           (MANUAL Microsoft procedure, without RollAlternateServiceAccountPassword.ps1
            which enumerates ALL the CAS of the organization and is blocked in
            2013/2016 coexistence by an internal version check).
           a. Generation of a random password (32 chars)
           b. Set-ADAccountPassword -Identity <ASA> -NewPassword <pwd> -Reset
           c. For each 2019 CAS:
              Set-ClientAccessService -Identity <srv> -AlternateServiceAccountCredential <PSCredential>

      5. Check that the target SPNs are not associated with any other account
         setspn -F -Q <SPN>

      6. Association of the SPNs with the ASA account
         setspn -S <SPN> DOMAIN\ACCT$

      7. Enabling Kerberos on the client services
         a. Outlook Anywhere : Set-OutlookAnywhere -InternalClientAuthenticationMethod Negotiate
         b. MAPI over HTTP   : Set-MapiVirtualDirectory -IISAuthenticationMethods Ntlm,Negotiate
            (already done by step 03 - checked again when the step is re-run)

      8. Final verification via Get-ClientAccessService -IncludeAlternateServiceAccountCredentialStatus

    Prerequisites:
        - ActiveDirectory module (RSAT) installed locally
        - Execution account with "Create Computer Objects" rights on the target OU
          (otherwise the account must be pre-created by the AD team)

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

    $asa            = $Config.KerberosASA
    $accountName    = $asa.AccountName.TrimEnd('$')          # short name without $
    $samAccount     = $accountName                            # SamAccount = name for a computer
    $netbiosDomain  = $asa.DomainNetBios
    $asaQualifiedUN = "$netbiosDomain\$accountName`$"         # CONTOSO\EXCHKRB$
    $ouPath         = $asa.OUPath
    $spnList        = @($asa.SPNs)

    Write-StepBanner -StepName '15' -Title 'Kerberos authentication for load-balanced Client Access' -Mode $Mode -Actions @(
        "1. Create the ASA computer account '$accountName' in AD if absent",
        '2. Enable AES (msDS-SupportedEncryptionTypes=28) on the ASA account',
        '3. RollAlternateServiceAccountPassword.ps1 -ToSpecificServer <FIRST_2019> -GenerateNewPasswordFor',
        '4. RollAlternateServiceAccountPassword.ps1 -ToSpecificServer <OTHERS> -CopyFrom <FIRST>',
        '   WARNING: the -CopyFrom propagation ALSO targets the Exchange 2013/2016 servers',
        '   so that all CAS share the same ASA credential (required for Kerberos).',
        '5. setspn -F -Q on each SPN (check that there is no conflict)',
        "6. setspn -S on each SPN ($($spnList -join ', '))",
        '7. Set-OutlookAnywhere -InternalClientAuthenticationMethod Negotiate (Microsoft procedure)',
        '8. Verification: Get-ClientAccessService -IncludeAlternateServiceAccountCredentialStatus (ALL CAS)'
    ) -IncludeLegacyServers `
       -LegacyImpactNote "The ASA credential is shared by ALL CAS (2019 + 2013/2016): Set-ClientAccessService -AlternateServiceAccountCredential will also apply to the legacy servers."

    Reset-Report
    Initialize-ExchangeShell -Credential $Credential
    $isWhatIf = ([bool]($Mode -eq 'Simulate'))

    # ==========================================================================
    # Prerequisite check: ActiveDirectory module
    # ==========================================================================
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        Write-Log "ActiveDirectory module not found. Install RSAT: Install-WindowsFeature RSAT-AD-PowerShell." -Level Error
        Add-Report -Step '15-ASA' -Target 'Prerequisites' -Action 'Import-Module ActiveDirectory' `
                   -Status Failed -ErrorMessage $_.Exception.Message
        Save-Report -StepName 'Step15-KerberosASA'
        return 1
    }

    # List of the CAS impacted by the ASA:
    #   - 2019: all detected servers
    #   - 2013/2016: all EXCEPT the Edge role (excluded from CAS load balancing)
    # The ASA credential is SHARED by all the CAS that answer the same
    # load-balanced SPNs: the propagation must therefore include the legacy servers.
    $servers       = Get-Exchange2019Servers
    $legacyServers = @(Get-LegacyExchangeServers | Where-Object { $_.ServerRole -notlike '*Edge*' })

    if (-not $servers) {
        Write-Log "No Exchange 2019 server detected." -Level Warning
        Save-Report -StepName 'Step15-KerberosASA'
        return 0
    }

    if ($legacyServers.Count -gt 0) {
        Write-Log "Legacy 2013/2016 servers included in the ASA propagation: $(($legacyServers.Name) -join ', ')" -Level Warning
    } else {
        Write-Log "No legacy 2013/2016 server detected - ASA propagation limited to 2019." -Level Sub
    }

    # Local helper: generate a complex password
    function New-StrongPassword {
        param([int]$Length = 32)
        $chars = ([char[]]([char]'A'..[char]'Z')) + ([char[]]([char]'a'..[char]'z')) + `
                 ([char[]]([char]'0'..[char]'9')) + ('!','@','#','$','%','^','&','*','-','_','+','=')
        -join (1..$Length | ForEach-Object { $chars | Get-Random })
    }

    # WARNING - Exchange remoting pitfall: through implicit remoting, the property
    # $cas.AlternateServiceAccountConfiguration is SERIALIZED as [System.String]
    # (the native ToString), NOT as a rich object. Therefore:
    #   $cas.AlternateServiceAccountConfiguration.LatestCredentials   <-- null on a string
    # while the string CONTAINS 'Latest: <date>, <UserName>'.
    # The string is therefore parsed with a regex using the helpers below.

    # Local helper: extract UserName + LastUpdate from the ASA string via regex
    function Get-AsaFromString {
        param([string]$AsaString)
        if (-not $AsaString) { return $null }
        $m = [regex]::Match($AsaString, '(?im)^Latest:\s+([^,]+),\s+(.+?)\s*$')
        if ($m.Success) {
            return @{ LastUpdate = $m.Groups[1].Value.Trim(); UserName = $m.Groups[2].Value.Trim() }
        }
        return $null
    }

    # Local helper: check whether a CAS has an ASA configured (parses the string)
    function Test-CasHasAsa {
        param([Parameter(Mandatory)] [string]$ServerName)
        $cas = Get-ClientAccessService -Identity $ServerName `
                   -IncludeAlternateServiceAccountCredentialStatus -ErrorAction SilentlyContinue
        if (-not $cas) { return $false }
        return [bool](Get-AsaFromString -AsaString ([string]$cas.AlternateServiceAccountConfiguration))
    }

    # Local helper: format the ASA string for the report (UserName + LastUpdate)
    function Format-CasAsa {
        param([Parameter(Mandatory)] [string]$ServerName)
        $cas = Get-ClientAccessService -Identity $ServerName `
                   -IncludeAlternateServiceAccountCredentialStatus -ErrorAction SilentlyContinue
        if (-not $cas) { return '<CAS not found>' }
        $asaStr = [string]$cas.AlternateServiceAccountConfiguration
        $parsed = Get-AsaFromString -AsaString $asaStr
        if ($parsed) { "ASA=$($parsed.UserName) | LastUpdate=$($parsed.LastUpdate)" }
        else { 'No ASA configured' }
    }

    # ==========================================================================
    # PHASE 1 - Creation of the ASA computer account
    # ==========================================================================
    Invoke-Action -Step '15-ASA' -Target $accountName -Action 'New-ADComputer' `
        -Detail "OU=$ouPath" `
        -InventoryScript {
            $c = Get-ADComputer -Filter "SamAccountName -eq '$samAccount`$'" -Properties Description, msDS-SupportedEncryptionTypes -ErrorAction SilentlyContinue
            if (-not $c) { 'ABSENT' }
            else { "Present | DN=$($c.DistinguishedName) | Enabled=$($c.Enabled) | EncTypes=$($c.'msDS-SupportedEncryptionTypes')" }
        } `
        -PreCheckScript {
            $c = Get-ADComputer -Filter "SamAccountName -eq '$samAccount`$'" -ErrorAction SilentlyContinue
            return [bool]$c
        } `
        -ActionScript {
            $pwdPlain = New-StrongPassword -Length 32
            $pwd = ConvertTo-SecureString -String $pwdPlain -AsPlainText -Force
            $params = @{
                Name              = $accountName
                SamAccountName    = $samAccount
                AccountPassword   = $pwd
                Enabled           = $true
                Description       = $asa.Description
                ErrorAction       = 'Stop'
            }
            if ($ouPath) { $params['Path'] = $ouPath }
            New-ADComputer @params -WhatIf:$isWhatIf
            if (-not $isWhatIf) {
                Write-Log "Computer account '$accountName' created (the password will be regenerated by RollAlternateServiceAccountPassword.ps1 in phase 3)." -Level Success
            }
        }

    # ==========================================================================
    # PHASE 2 - Enabling AES (msDS-SupportedEncryptionTypes = 28)
    # ==========================================================================
    Invoke-Action -Step '15-ASA' -Target $accountName -Action 'Set-ADComputer msDS-SupportedEncryptionTypes=28' `
        -Detail 'AES128 + AES256 + RC4-HMAC' `
        -InventoryScript {
            $c = Get-ADComputer -Filter "SamAccountName -eq '$samAccount`$'" -Properties msDS-SupportedEncryptionTypes -ErrorAction SilentlyContinue
            if ($c) { "EncTypes=$($c.'msDS-SupportedEncryptionTypes')" } else { 'Account not found' }
        } `
        -PreCheckScript {
            $c = Get-ADComputer -Filter "SamAccountName -eq '$samAccount`$'" -Properties msDS-SupportedEncryptionTypes -ErrorAction SilentlyContinue
            if (-not $c) { return $false }
            return ([int]$c.'msDS-SupportedEncryptionTypes' -eq 28)
        } `
        -ActionScript {
            # In Simulate: if the AD account does not exist yet (just
            # "created" by New-ADComputer -WhatIf), Set-ADComputer -WhatIf
            # fails anyway because the cmdlet resolves -Identity first.
            # It is bypassed in that case.
            if ($isWhatIf) {
                $exists = [bool](Get-ADComputer -Filter "SamAccountName -eq '$samAccount`$'" -ErrorAction SilentlyContinue)
                if (-not $exists) {
                    Write-Log "  [Simulate] Set-ADComputer $samAccount EncTypes=28 (account will be created by Step15 Apply)" -Level Sub
                    return
                }
            }
            Set-ADComputer -Identity $samAccount -Replace @{ 'msDS-SupportedEncryptionTypes' = 28 } -ErrorAction Stop -WhatIf:$isWhatIf
        }

    # ==========================================================================
    # PHASE 3 + 4 - Deployment of the ASA credential via RollAlternateServiceAccountPassword.ps1
    # ==========================================================================
    $rollScript = Join-Path $env:ExchangeInstallPath 'Scripts\RollAlternateServiceAccountPassword.ps1'

    if (-not (Test-Path $rollScript)) {
        Write-Log "Microsoft script not found: $rollScript" -Level Error
        Add-Report -Step '15-ASA' -Target 'RollAlternateServiceAccountPassword' -Action 'Test-Path' `
                   -Status Failed -ErrorMessage "Script not found: $rollScript"
        Save-Report -StepName 'Step15-KerberosASA'
        return 1
    }

    # Global pre-check: do all the CAS (2019 + 2013/2016) already have an ASA configured?
    # (Test-CasHasAsa handles the remoting String vs Object pitfall - see helpers above.)
    $allCasForCheck = @($servers) + @($legacyServers)
    $allDeployed = $true
    foreach ($s in $allCasForCheck) {
        if (-not (Test-CasHasAsa -ServerName $s.Name)) {
            $allDeployed = $false
            break
        }
    }

    # Phase 3: the password generation starts on a 2019 server.
    # Phase 4: the -CopyFrom propagation ALSO targets the 2013/2016 servers
    # (the ASA credential must be shared by all the CAS that answer the
    # same load-balanced SPNs).
    $firstServer  = $servers | Select-Object -First 1
    $otherServers = @($servers | Select-Object -Skip 1) + @($legacyServers)

    # PHASE 3 - First server (-GenerateNewPasswordFor)
    # In Simulate: NOTHING is executed (the Microsoft script has no -WhatIf and
    # contains a ShouldContinue() Y/N prompt that blocks in NonInteractive).
    Invoke-Action -Step '15-ASA' -Target $firstServer.Name -Action 'RollAlternateServiceAccountPassword (first/-GenerateNewPasswordFor)' `
        -Detail "ASA=$asaQualifiedUN" `
        -InventoryScript { Format-CasAsa -ServerName $firstServer.Name } `
        -PreCheckScript  { Test-CasHasAsa -ServerName $firstServer.Name } `
        -ActionScript {
            if ($Mode -eq 'Simulate') {
                Write-Log "  [Simulate] RollAlternateServiceAccountPassword -ToSpecificServer $($firstServer.Fqdn) -GenerateNewPasswordFor $asaQualifiedUN" -Level Sub
                return
            }
            # Apply: real execution. The Microsoft script asks for an interactive
            # confirmation (ShouldContinue) -> requires an interactive console.
            Push-Location (Split-Path $rollScript -Parent)
            try {
                $prevConfirm = $ConfirmPreference
                $ConfirmPreference = 'None'
                & $rollScript -ToSpecificServer $firstServer.Fqdn -GenerateNewPasswordFor $asaQualifiedUN -Verbose 4>&1 | ForEach-Object { Write-Log "  [Roll] $_" -Level Sub }
                $ConfirmPreference = $prevConfirm
            } finally {
                Pop-Location
            }
        }

    # PHASE 4 - Other servers (-CopyFrom)
    foreach ($s in $otherServers) {
        Invoke-Action -Step '15-ASA' -Target $s.Name -Action 'RollAlternateServiceAccountPassword (-CopyFrom)' `
            -Detail "Source=$($firstServer.Name)" `
            -InventoryScript { Format-CasAsa -ServerName $s.Name } `
            -PreCheckScript  { Test-CasHasAsa -ServerName $s.Name } `
            -ActionScript {
                if ($Mode -eq 'Simulate') {
                    Write-Log "  [Simulate] RollAlternateServiceAccountPassword -ToSpecificServer $($s.Fqdn) -CopyFrom $($firstServer.Fqdn)" -Level Sub
                    return
                }
                Push-Location (Split-Path $rollScript -Parent)
                try {
                    & $rollScript -ToSpecificServer $s.Fqdn -CopyFrom $firstServer.Fqdn -Verbose 4>&1 | ForEach-Object { Write-Log "  [Roll] $_" -Level Sub }
                } finally {
                    Pop-Location
                }
            }
    }

    # ==========================================================================
    # PHASE 5 - Check that the SPNs are not associated with any other account
    # IMPORTANT (Microsoft doc): do not associate the SPNs until the ASA
    # credential has been deployed on AT LEAST ONE server.
    # ==========================================================================
    foreach ($spn in $spnList) {
        Invoke-Action -Step '15-ASA' -Target $spn -Action 'setspn -F -Q (conflict check)' `
            -InventoryScript {
                $out = & setspn.exe -F -Q $spn 2>&1 | Out-String
                $out.Trim()
            } `
            -PreCheckScript {
                # No pre-check: the check is always run, it is non-destructive.
                return $false
            } `
            -ActionScript {
                $out = & setspn.exe -F -Q $spn 2>&1 | Out-String
                # The command returns something like:
                # "Checking forest DC=contoso,DC=local
                #  No such SPN found."
                # or, in case of conflict:
                # "CN=OtherAccount,...   http/mail.contoso.com"
                if ($out -notmatch 'No such SPN found' -and $out -match 'CN=') {
                    # If a line contains "CN=" and is NOT the one of our ASA account,
                    # there is a conflict.
                    $cnLines = $out -split "`r?`n" | Where-Object { $_ -match 'CN=' }
                    $conflict = $cnLines | Where-Object { $_ -notmatch ('CN=' + [regex]::Escape($accountName) + '\b') }
                    if ($conflict) {
                        throw "SPN $spn already associated with another account: $($conflict -join ' / ')"
                    }
                }
            }
    }

    # ==========================================================================
    # PHASE 6 - Association of the SPNs with the ASA account via setspn -S
    # ==========================================================================
    foreach ($spn in $spnList) {
        Invoke-Action -Step '15-ASA' -Target "$spn -> $asaQualifiedUN" -Action 'setspn -S' `
            -InventoryScript {
                $comp = Get-ADComputer -Filter "SamAccountName -eq '$samAccount`$'" -ErrorAction SilentlyContinue
                if (-not $comp) { return "Account $asaQualifiedUN not found - SPN not applicable" }
                try {
                    $out = & setspn.exe -L $asaQualifiedUN 2>&1 | Out-String
                    if ($out.Trim()) { $out.Trim() } else { "No SPN registered for $asaQualifiedUN" }
                } catch {
                    "setspn -L failed: $($_.Exception.Message.Trim())"
                }
            } `
            -PreCheckScript {
                $out = & setspn.exe -L $asaQualifiedUN 2>&1 | Out-String
                return ($out -match [regex]::Escape($spn))
            } `
            -ActionScript {
                # setspn -S has no native -WhatIf: bypassed in Simulate.
                if ($isWhatIf) {
                    Write-Log "  [Simulate] setspn -S $spn $asaQualifiedUN" -Level Sub
                    return
                }
                # setspn -S does not create duplicates: it checks first and fails on conflict.
                $out = & setspn.exe -S $spn $asaQualifiedUN 2>&1 | Out-String
                if ($out -match 'Duplicate SPN found, aborting operation') {
                    throw "Duplicate SPN $spn in the forest. setspn refuses the addition."
                }
                Write-Log "  [setspn -S] $($out.Trim())" -Level Sub
            }
    }

    # ==========================================================================
    # PHASE 7a - Enabling Kerberos on Outlook Anywhere
    #            Set-OutlookAnywhere -InternalClientAuthenticationMethod Negotiate
    # ==========================================================================
    foreach ($s in $servers) {
        $oaIdentity = "$($s.Name)\Rpc (Default Web Site)"

        Invoke-Action -Step '15-ASA' -Target $oaIdentity -Action 'Set-OutlookAnywhere Negotiate' `
            -InventoryScript {
                $oa = Get-OutlookAnywhere -Identity $oaIdentity -ErrorAction Stop
                "InternalAuth=$($oa.InternalClientAuthenticationMethod) | IISAuth=$($oa.IISAuthenticationMethods -join '+')"
            } `
            -PreCheckScript {
                $oa = Get-OutlookAnywhere -Identity $oaIdentity -ErrorAction Stop
                return ($oa.InternalClientAuthenticationMethod -eq 'Negotiate')
            } `
            -ActionScript {
                Set-OutlookAnywhere -Identity $oaIdentity `
                    -InternalClientAuthenticationMethod Negotiate `
                    -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
            }
    }

    # ==========================================================================
    # PHASE 7b - MAPI over HTTP check: Ntlm + Negotiate
    #            (already configured by step 03, checked just in case)
    # ==========================================================================
    foreach ($s in $servers) {
        $mapiIdentity = "$($s.Name)\mapi (Default Web Site)"

        Invoke-Action -Step '15-ASA' -Target $mapiIdentity -Action 'Set-MapiVirtualDirectory Ntlm,Negotiate' `
            -InventoryScript {
                $m = Get-MapiVirtualDirectory -Identity $mapiIdentity -ErrorAction Stop
                "IISAuth=$($m.IISAuthenticationMethods -join '+')"
            } `
            -PreCheckScript {
                $m = Get-MapiVirtualDirectory -Identity $mapiIdentity -ErrorAction Stop
                return (($m.IISAuthenticationMethods -contains 'Ntlm') -and ($m.IISAuthenticationMethods -contains 'Negotiate'))
            } `
            -ActionScript {
                Set-MapiVirtualDirectory -Identity $mapiIdentity `
                    -IISAuthenticationMethods Ntlm,Negotiate `
                    -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
            }
    }

    # ==========================================================================
    # PHASE 8 - Final verification per server (2019 + 2013/2016)
    # ==========================================================================
    # After phases 3/4 (RollAlternateServiceAccountPassword), the AD attribute
    # msExchAlternateServiceAccountConfiguration is written on the DC of the
    # Exchange server's site. An immediate check through Get-ClientAccessService
    # may query a different DC that has not received the replication yet,
    # hence the need to:
    #   1. Target a DC of the SAME site as the Exchange server being tested
    #   2. Retry several times with a delay to absorb the network latency

    # Local helper: checks the ASA on a server with retry (legitimate replication delay).
    # Uses Get-AsaFromString defined above (see the remoting String vs Object pitfall).
    function Get-CasAsaStatus {
        param(
            [Parameter(Mandatory)] [string]$ServerName,
            [int]$MaxTry = 6,
            [int]$DelaySec = 10
        )
        for ($i = 1; $i -le $MaxTry; $i++) {
            $cas = Get-ClientAccessService -Identity $ServerName `
                       -IncludeAlternateServiceAccountCredentialStatus -ErrorAction SilentlyContinue
            $asaStr = [string]$cas.AlternateServiceAccountConfiguration
            $parsed = Get-AsaFromString -AsaString $asaStr
            if ($parsed) {
                return @{ Found = $true; Tries = $i; UserName = $parsed.UserName; LastUpdate = $parsed.LastUpdate; Raw = $asaStr }
            }
            if ($i -lt $MaxTry) { Start-Sleep -Seconds $DelaySec }
        }
        return @{ Found = $false; Tries = $MaxTry; UserName = $null; LastUpdate = $null; Raw = $asaStr }
    }

    $allCasForVerify = @($servers) + @($legacyServers)
    foreach ($s in $allCasForVerify) {
        Invoke-Action -Step '15-ASA' -Target $s.Name -Action 'Verify Get-ClientAccessService ASA status' `
            -Detail "Parsing of the ASA string + 6x10s retry in case of AD replication delay" `
            -InventoryScript {
                $cas = Get-ClientAccessService -Identity $s.Name `
                           -IncludeAlternateServiceAccountCredentialStatus -ErrorAction SilentlyContinue
                if (-not $cas) { return 'CAS not found' }
                $asaStr = [string]$cas.AlternateServiceAccountConfiguration
                $parsed = Get-AsaFromString -AsaString $asaStr
                if ($parsed) { "ASA={0} | LastUpdate={1}" -f $parsed.UserName, $parsed.LastUpdate }
                else { "No ASA (raw='$asaStr')" }
            } `
            -PreCheckScript {
                $cas = Get-ClientAccessService -Identity $s.Name `
                           -IncludeAlternateServiceAccountCredentialStatus -ErrorAction SilentlyContinue
                if (-not $cas) { return $false }
                $parsed = Get-AsaFromString -AsaString ([string]$cas.AlternateServiceAccountConfiguration)
                return [bool]$parsed
            } `
            -ActionScript {
                if ($Mode -eq 'Simulate') {
                    Write-Log "  [Simulate] Verify ASA on $($s.Name) (deployed at the next Apply)" -Level Sub
                    return
                }
                # Retry to absorb possible AD replication delays
                $r = Get-CasAsaStatus -ServerName $s.Name -MaxTry 6 -DelaySec 10
                if ($r.Found) {
                    Write-Log "  $($s.Name): ASA visible after $($r.Tries) attempt(s) (ASA=$($r.UserName) LastUpdate=$($r.LastUpdate))" -Level Sub
                    return
                }
                throw ("ASA not visible on {0} after {1} attempt(s) (60s max). Check the output of RollAlternateServiceAccountPassword.ps1 in the transcript." -f `
                    $s.Name, $r.Tries)
            }
    }

    Save-Report -StepName 'Step15-KerberosASA'
    return 0
}
