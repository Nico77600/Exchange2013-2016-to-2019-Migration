<#
.SYNOPSIS
    Step 03 - Configure URLs and authentication on the Exchange 2019 virtual directories.

.DESCRIPTION
    Configures the URLs and the authentication mechanisms of the Exchange 2019 virtual
    directories.

    The URLs and the authentication mechanisms are read from a reference
    Exchange 2013/2016 server:
        - SourceServerForVDirs in Deployment.config.psd1 (priority)
        - Otherwise: first Exchange 2013/2016 Mailbox server detected automatically
        - Fallback on UrlInterne / UrlExterne + standard Exchange auth values if absent

    Configured for each 2019 server (URLs + Auth):
        - SCP      : AutoDiscoverServiceInternalUri
        - OWA      : InternalUrl, ExternalUrl, FormsAuthentication, LogonFormat
        - ECP      : InternalUrl, ExternalUrl, FormsAuthentication
        - AutoDisc : InternalUrl, ExternalUrl, BasicAuthentication,
                     WindowsAuthentication, WSSecurityAuthentication
        - OAB      : InternalUrl, ExternalUrl, BasicAuthentication,
                     WindowsAuthentication, RequireSSL
        - EWS      : InternalUrl, ExternalUrl, BasicAuthentication,
                     WindowsAuthentication, OAuthAuthentication
        - EAS      : InternalUrl, ExternalUrl, BasicAuthEnabled, WindowsAuthEnabled
        - OA       : ExternalHostname, InternalHostname,
                     ExternalClientAuthenticationMethod,
                     InternalClientAuthenticationMethod,
                     IISAuthenticationMethods, SSLOffloading
        - MAPI     : InternalUrl, IISAuthenticationMethods
        - EPA      : ExtendedProtectionTokenChecking on all 2019 VDirs
                     (OWA/ECP/MAPI=Require | OA=Require | EWS/OAB/EAS=Allow | AutoDisc=None)
                     Fixed target values - nothing is read from the legacy source.

    No action is taken on the Exchange 2013/2016 servers in coexistence.
    Idempotent: compares URL + Auth + EPA before any change.

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

    Initialize-ExchangeShell -Credential $Credential

    # --- 0. Source server resolution ---
    $sourceServer  = $null
    $sourceVersion = ''

    if ($Config.SourceServerForVDirs -and ($Config.SourceServerForVDirs -ne '')) {
        $sourceServer  = $Config.SourceServerForVDirs
        $sourceVersion = '(config)'
    } else {
        $legacy = Get-LegacyExchangeServers |
                  Where-Object { $_.ServerRole -notlike '*Edge*' } |
                  Select-Object -First 1
        if ($legacy) {
            $sourceServer  = $legacy.Name
            $sourceVersion = if ($legacy.AdminDisplayVersion -like 'Version 15.1*') { '[Exchange 2016]' } else { '[Exchange 2013]' }
        }
    }

    $sourceDesc   = if ($sourceServer) { "$sourceServer $sourceVersion" } else { "config ($($Config.UrlInterne))" }
    $sourceTarget = if ($sourceServer) { $sourceServer } else { '<config>' }

    Write-StepBanner -StepName '03' -Title 'URL and Auth configuration of the Virtual Directories' -Mode $Mode -Actions @(
        "Source URLs + Auth: $sourceDesc",
        'Configure SCP / Autodiscover Service Internal URI',
        'Configure InternalUrl / ExternalUrl + Auth mechanisms: OWA, ECP, Autodiscover, OAB, EWS, EAS, OA, MAPI',
        'Auth taken from the source server (2013/2016) - fallback on standard Exchange values',
        'Verify and configure Extended Protection: OWA/ECP/MAPI=Require | OA=Require | EWS/OAB/EAS=Allow | AutoDisc=None',
        'Sequential IIS reset (one server at a time) to apply the SCP + VDir + ExtendedProtection changes',
        'No action on the Exchange 2013/2016 servers in coexistence'
    )

    Reset-Report

    # --- 1a. Reference values - URLs (fallback = config) ---
    $urlInt = $Config.UrlInterne
    $urlExt = $Config.UrlExterne

    $refScpInt    = "https://$urlInt/Autodiscover/Autodiscover.xml"
    $refOwaInt    = "https://$urlInt/OWA"
    $refOwaExt    = "https://$urlExt/OWA"
    $refEcpInt    = "https://$urlInt/ECP"
    $refEcpExt    = "https://$urlExt/ECP"
    $refAutoInt   = "https://$urlInt/Autodiscover"
    $refAutoExt   = "https://$urlExt/Autodiscover"
    $refOabInt    = "https://$urlInt/OAB"
    $refOabExt    = "https://$urlExt/OAB"
    $refEwsInt    = "https://$urlInt/EWS/Exchange.asmx"
    $refEwsExt    = "https://$urlExt/EWS/Exchange.asmx"
    $refEasInt    = "https://$urlInt/Microsoft-Server-ActiveSync"
    $refEasExt    = "https://$urlExt/Microsoft-Server-ActiveSync"
    $refOaIntHost = $urlInt
    $refOaExtHost = $urlExt
    $refMapiInt   = "https://$urlInt/mapi"

    # --- 1b. Reference values - Auth (fallback = standard Exchange values) ---
    $refOwaFormsAuth  = $true
    $refOwaLogonFmt   = 'FullDomain'
    $refOwaDefDomain  = ''      # required if LogonFormat=UserName
    $refEcpFormsAuth  = $true
    $refAutoBasicAuth = $true
    $refAutoWinAuth   = $true
    $refAutoWSSecAuth = $true
    $refOabBasicAuth  = $false
    $refOabWinAuth    = $true
    $refEwsBasicAuth  = $false
    $refEwsWinAuth    = $true
    $refEwsOAuthAuth  = $true
    $refEasBasicAuth  = $true
    $refEasWinAuth    = $false
    $refOaExtAuth     = 'Negotiate'
    $refOaIntAuth     = 'Ntlm'
    $refOaIISAuth     = 'Basic,Ntlm,Negotiate'
    $refOaSSLOff      = $false
    $refMapiIISAuth   = 'Ntlm,OAuth,Negotiate'

    # --- 2. Read URLs + Auth from the source server ---
    if ($sourceServer) {
        try {
            $svc = Get-ClientAccessService $sourceServer -ErrorAction Stop
            if ($svc.AutoDiscoverServiceInternalUri) { $refScpInt = [string]$svc.AutoDiscoverServiceInternalUri }
        } catch { Write-Log "Cannot read SCP on $sourceServer - $($_.Exception.Message)" -Level Warning }

        try {
            $vd = Get-OwaVirtualDirectory -Server $sourceServer -ErrorAction Stop | Select-Object -First 1
            if ($vd.InternalUrl)                    { $refOwaInt      = [string]$vd.InternalUrl }
            if ($vd.ExternalUrl)                    { $refOwaExt      = [string]$vd.ExternalUrl }
            if ($null -ne $vd.FormsAuthentication)  { $refOwaFormsAuth = [bool]$vd.FormsAuthentication }
            if ($vd.LogonFormat)                    { $refOwaLogonFmt  = [string]$vd.LogonFormat }
            if ($vd.DefaultDomain)                  { $refOwaDefDomain = [string]$vd.DefaultDomain }
        } catch { Write-Log "Cannot read OWA on $sourceServer - $($_.Exception.Message)" -Level Warning }

        try {
            $vd = Get-EcpVirtualDirectory -Server $sourceServer -ErrorAction Stop | Select-Object -First 1
            if ($vd.InternalUrl)                   { $refEcpInt      = [string]$vd.InternalUrl }
            if ($vd.ExternalUrl)                   { $refEcpExt      = [string]$vd.ExternalUrl }
            if ($null -ne $vd.FormsAuthentication) { $refEcpFormsAuth = [bool]$vd.FormsAuthentication }
        } catch { Write-Log "Cannot read ECP on $sourceServer - $($_.Exception.Message)" -Level Warning }

        try {
            $vd = Get-AutodiscoverVirtualDirectory -Server $sourceServer -ErrorAction Stop | Select-Object -First 1
            if ($vd.InternalUrl)                       { $refAutoInt      = [string]$vd.InternalUrl }
            if ($vd.ExternalUrl)                       { $refAutoExt      = [string]$vd.ExternalUrl }
            if ($null -ne $vd.BasicAuthentication)     { $refAutoBasicAuth = [bool]$vd.BasicAuthentication }
            if ($null -ne $vd.WindowsAuthentication)   { $refAutoWinAuth   = [bool]$vd.WindowsAuthentication }
            if ($null -ne $vd.WSSecurityAuthentication){ $refAutoWSSecAuth  = [bool]$vd.WSSecurityAuthentication }
        } catch { Write-Log "Cannot read Autodiscover VDir on $sourceServer - $($_.Exception.Message)" -Level Warning }

        try {
            $vd = Get-OabVirtualDirectory -Server $sourceServer -ErrorAction Stop | Select-Object -First 1
            if ($vd.InternalUrl)                   { $refOabInt      = [string]$vd.InternalUrl }
            if ($vd.ExternalUrl)                   { $refOabExt      = [string]$vd.ExternalUrl }
            if ($null -ne $vd.BasicAuthentication) { $refOabBasicAuth = [bool]$vd.BasicAuthentication }
            if ($null -ne $vd.WindowsAuthentication){ $refOabWinAuth  = [bool]$vd.WindowsAuthentication }
        } catch { Write-Log "Cannot read OAB on $sourceServer - $($_.Exception.Message)" -Level Warning }

        try {
            $vd = Get-WebServicesVirtualDirectory -Server $sourceServer -ErrorAction Stop | Select-Object -First 1
            if ($vd.InternalUrl)                   { $refEwsInt      = [string]$vd.InternalUrl }
            if ($vd.ExternalUrl)                   { $refEwsExt      = [string]$vd.ExternalUrl }
            if ($null -ne $vd.BasicAuthentication) { $refEwsBasicAuth = [bool]$vd.BasicAuthentication }
            if ($null -ne $vd.WindowsAuthentication){ $refEwsWinAuth  = [bool]$vd.WindowsAuthentication }
            if ($null -ne $vd.OAuthAuthentication) { $refEwsOAuthAuth = [bool]$vd.OAuthAuthentication }
        } catch { Write-Log "Cannot read EWS on $sourceServer - $($_.Exception.Message)" -Level Warning }

        try {
            $vd = Get-ActiveSyncVirtualDirectory -Server $sourceServer -ErrorAction Stop | Select-Object -First 1
            if ($vd.InternalUrl)                { $refEasInt      = [string]$vd.InternalUrl }
            if ($vd.ExternalUrl)                { $refEasExt      = [string]$vd.ExternalUrl }
            if ($null -ne $vd.BasicAuthEnabled) { $refEasBasicAuth = [bool]$vd.BasicAuthEnabled }
            if ($null -ne $vd.WindowsAuthEnabled){ $refEasWinAuth  = [bool]$vd.WindowsAuthEnabled }
        } catch { Write-Log "Cannot read ActiveSync on $sourceServer - $($_.Exception.Message)" -Level Warning }

        try {
            $oa = Get-OutlookAnywhere -Server $sourceServer -ErrorAction Stop | Select-Object -First 1
            if ($oa.InternalHostname)                         { $refOaIntHost = [string]$oa.InternalHostname }
            if ($oa.ExternalHostname)                         { $refOaExtHost = [string]$oa.ExternalHostname }
            if ($oa.ExternalClientAuthenticationMethod)       { $refOaExtAuth = [string]$oa.ExternalClientAuthenticationMethod }
            if ($oa.InternalClientAuthenticationMethod)       { $refOaIntAuth = [string]$oa.InternalClientAuthenticationMethod }
            $oaIIS = [string]$oa.IISAuthenticationMethods
            if ($oaIIS)                                       { $refOaIISAuth = $oaIIS -replace '\s+', ',' }
            if ($null -ne $oa.SSLOffloading)                  { $refOaSSLOff  = [bool]$oa.SSLOffloading }
        } catch { Write-Log "Cannot read OutlookAnywhere on $sourceServer - $($_.Exception.Message)" -Level Warning }

        try {
            $vd = Get-MapiVirtualDirectory -Server $sourceServer -ErrorAction Stop | Select-Object -First 1
            if ($vd.InternalUrl) { $refMapiInt = [string]$vd.InternalUrl }
            $mapiIIS = [string]$vd.IISAuthenticationMethods
            if ($mapiIIS) { $refMapiIISAuth = $mapiIIS -replace '\s+', ',' }
        } catch { Write-Log "Cannot read MAPI on $sourceServer - $($_.Exception.Message)" -Level Warning }
    }

    # If LogonFormat=UserName but no DefaultDomain was read => fallback on DomainNetBios from the config
    if ($refOwaLogonFmt -eq 'UserName' -and -not $refOwaDefDomain) {
        $refOwaDefDomain = [string]$Config.DomainNetBios
        Write-Log "OWA LogonFormat=UserName: DefaultDomain not read from the source, fallback on DomainNetBios=$refOwaDefDomain" -Level Warning
    }

    Write-Log "Reference URLs (source: $sourceDesc)" -Level Sub
    Write-Log "  SCP=$refScpInt" -Level Sub
    Write-Log "  OWA Int=$refOwaInt Ext=$refOwaExt DefaultDomain=$refOwaDefDomain" -Level Sub
    Write-Log "  ECP Int=$refEcpInt Ext=$refEcpExt" -Level Sub
    Write-Log "  Auth: OWA Forms=$refOwaFormsAuth Logon=$refOwaLogonFmt | AutoDisc Basic=$refAutoBasicAuth Win=$refAutoWinAuth WSSec=$refAutoWSSecAuth" -Level Sub
    Write-Log "  Auth: EWS Basic=$refEwsBasicAuth Win=$refEwsWinAuth OAuth=$refEwsOAuthAuth | EAS Basic=$refEasBasicAuth Win=$refEasWinAuth" -Level Sub
    Write-Log "  Auth: OA ExtAuth=$refOaExtAuth IntAuth=$refOaIntAuth IISAuth=$refOaIISAuth SSLOff=$refOaSSLOff | MAPI IISAuth=$refMapiIISAuth" -Level Sub

    Add-Report -Step '03-VDir' -Target $sourceTarget -Action 'ResolveSourceUrls' -Status Inventoried `
               -Detail "SCP=$refScpInt; OWA=$refOwaInt/$refOwaExt; OA Int=$refOaIntHost Ext=$refOaExtHost"

    Add-Report -Step '03-VDir' -Target $sourceTarget -Action 'ResolveSourceAuth' -Status Inventoried `
               -Detail ("OWA Forms={0} Logon={1} | EWS Basic={2} Win={3} OAuth={4} | OA ExtAuth={5} IntAuth={6} IISAuth={7} | MAPI IISAuth={8}" -f `
                   $refOwaFormsAuth, $refOwaLogonFmt, $refEwsBasicAuth, $refEwsWinAuth, $refEwsOAuthAuth, `
                   $refOaExtAuth, $refOaIntAuth, $refOaIISAuth, $refMapiIISAuth)

    # --- 3. Apply on each 2019 server ---
    $servers = @(Get-Exchange2019Servers)

    foreach ($srv in $servers) {
        $server = $srv.Name

        # --- SCP ---
        $autodiscoverUri = $refScpInt
        Invoke-Action -Step '03-VDir' -Target $server -Action 'Set-ClientAccessService SCP' -Detail $autodiscoverUri `
            -InventoryScript {
                $svc = Get-ClientAccessService $server -ErrorAction Stop
                "AutoDiscoverServiceInternalUri={0}" -f $svc.AutoDiscoverServiceInternalUri
            } `
            -PreCheckScript {
                $svc = Get-ClientAccessService $server -ErrorAction Stop
                return ([string]$svc.AutoDiscoverServiceInternalUri -eq $autodiscoverUri)
            } `
            -ActionScript {
                Set-ClientAccessService -Identity $server -AutoDiscoverServiceInternalUri $autodiscoverUri `
                    -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
            }

        # --- OWA ---
        $owaIdentity   = "$server\OWA (Default Web Site)"
        $owaIntUrl     = $refOwaInt
        $owaExtUrl     = $refOwaExt
        $owaFormsAuth  = $refOwaFormsAuth
        $owaLogonFmt   = $refOwaLogonFmt
        $owaDefDomain  = $refOwaDefDomain
        $owaDetail     = if ($owaLogonFmt -eq 'UserName') { "$owaIntUrl | $owaExtUrl | Forms=$owaFormsAuth Logon=$owaLogonFmt DefaultDomain=$owaDefDomain" }
                         else                              { "$owaIntUrl | $owaExtUrl | Forms=$owaFormsAuth Logon=$owaLogonFmt" }
        Invoke-Action -Step '03-VDir' -Target $server -Action 'Set-OwaVirtualDirectory URLs+Auth' `
            -Detail $owaDetail `
            -InventoryScript {
                $vd = Get-OwaVirtualDirectory -Server $server -ErrorAction Stop | Select-Object -First 1
                "Int={0}; Ext={1}; IISAuth={2}; Forms={3}; Logon={4}; DefaultDomain={5}" -f `
                    $vd.InternalUrl, $vd.ExternalUrl, ([string]$vd.IISAuthenticationMethods), $vd.FormsAuthentication, $vd.LogonFormat, $vd.DefaultDomain
            } `
            -PreCheckScript {
                $vd = Get-OwaVirtualDirectory -Identity $owaIdentity -ErrorAction Stop
                $okBase = ([string]$vd.InternalUrl -eq $owaIntUrl -and `
                           [string]$vd.ExternalUrl -eq $owaExtUrl -and `
                           [bool]$vd.FormsAuthentication -eq $owaFormsAuth -and `
                           [string]$vd.LogonFormat -eq $owaLogonFmt)
                if (-not $okBase) { return $false }
                if ($owaLogonFmt -eq 'UserName') { return ([string]$vd.DefaultDomain -eq $owaDefDomain) }
                return $true
            } `
            -ActionScript {
                $isWhatIf = ([bool]($Mode -eq 'Simulate'))
                if ($owaLogonFmt -eq 'UserName') {
                    Set-OwaVirtualDirectory -Identity $owaIdentity `
                        -InternalUrl $owaIntUrl -ExternalUrl $owaExtUrl `
                        -FormsAuthentication:$owaFormsAuth -LogonFormat $owaLogonFmt -DefaultDomain $owaDefDomain `
                        -WhatIf:$isWhatIf -ErrorAction Stop
                } else {
                    Set-OwaVirtualDirectory -Identity $owaIdentity `
                        -InternalUrl $owaIntUrl -ExternalUrl $owaExtUrl `
                        -FormsAuthentication:$owaFormsAuth -LogonFormat $owaLogonFmt `
                        -WhatIf:$isWhatIf -ErrorAction Stop
                }
            }

        # --- ECP ---
        $ecpIdentity  = "$server\ECP (Default Web Site)"
        $ecpIntUrl    = $refEcpInt
        $ecpExtUrl    = $refEcpExt
        $ecpFormsAuth = $refEcpFormsAuth
        Invoke-Action -Step '03-VDir' -Target $server -Action 'Set-EcpVirtualDirectory URLs+Auth' `
            -Detail "$ecpIntUrl | $ecpExtUrl | Forms=$ecpFormsAuth" `
            -InventoryScript {
                $vd = Get-EcpVirtualDirectory -Identity $ecpIdentity -ErrorAction Stop
                "Int={0}; Ext={1}; IISAuth={2}; Forms={3}" -f `
                    $vd.InternalUrl, $vd.ExternalUrl, ([string]$vd.IISAuthenticationMethods), $vd.FormsAuthentication
            } `
            -PreCheckScript {
                $vd = Get-EcpVirtualDirectory -Identity $ecpIdentity -ErrorAction Stop
                return ([string]$vd.InternalUrl -eq $ecpIntUrl -and `
                        [string]$vd.ExternalUrl -eq $ecpExtUrl -and `
                        [bool]$vd.FormsAuthentication -eq $ecpFormsAuth)
            } `
            -ActionScript {
                Set-EcpVirtualDirectory -Identity $ecpIdentity `
                    -InternalUrl $ecpIntUrl -ExternalUrl $ecpExtUrl `
                    -FormsAuthentication:$ecpFormsAuth `
                    -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
            }

        # --- Autodiscover VDir ---
        # NB: Set-AutodiscoverVirtualDirectory does NOT expose InternalUrl/ExternalUrl.
        # These Autodiscover URLs are published via the SCP (Set-ClientAccessService -AutoDiscoverServiceInternalUri).
        # Here only the authentication mechanisms are configured.
        $autoIdentity   = "$server\Autodiscover (Default Web Site)"
        $autoBasicAuth  = $refAutoBasicAuth
        $autoWinAuth    = $refAutoWinAuth
        $autoWSSecAuth  = $refAutoWSSecAuth
        Invoke-Action -Step '03-VDir' -Target $server -Action 'Set-AutodiscoverVirtualDirectory Auth' `
            -Detail "Basic=$autoBasicAuth Win=$autoWinAuth WSSec=$autoWSSecAuth" `
            -InventoryScript {
                $vd = Get-AutodiscoverVirtualDirectory -Identity $autoIdentity -ErrorAction Stop
                "Basic={0}; Win={1}; WSSec={2}" -f `
                    $vd.BasicAuthentication, $vd.WindowsAuthentication, $vd.WSSecurityAuthentication
            } `
            -PreCheckScript {
                $vd = Get-AutodiscoverVirtualDirectory -Identity $autoIdentity -ErrorAction Stop
                return ([bool]$vd.BasicAuthentication -eq $autoBasicAuth -and `
                        [bool]$vd.WindowsAuthentication -eq $autoWinAuth -and `
                        [bool]$vd.WSSecurityAuthentication -eq $autoWSSecAuth)
            } `
            -ActionScript {
                Set-AutodiscoverVirtualDirectory -Identity $autoIdentity `
                    -BasicAuthentication:$autoBasicAuth `
                    -WindowsAuthentication:$autoWinAuth `
                    -WSSecurityAuthentication:$autoWSSecAuth `
                    -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
            }

        # --- OAB ---
        $oabIdentity  = "$server\OAB (Default Web Site)"
        $oabIntUrl    = $refOabInt
        $oabExtUrl    = $refOabExt
        $oabBasicAuth = $refOabBasicAuth
        $oabWinAuth   = $refOabWinAuth
        Invoke-Action -Step '03-VDir' -Target $server -Action 'Set-OabVirtualDirectory URLs+Auth' `
            -Detail "$oabIntUrl | $oabExtUrl | Basic=$oabBasicAuth Win=$oabWinAuth RequireSSL=True" `
            -InventoryScript {
                $vd = Get-OabVirtualDirectory -Identity $oabIdentity -ErrorAction Stop
                "Int={0}; Ext={1}; Basic={2}; Win={3}; RequireSSL={4}" -f `
                    $vd.InternalUrl, $vd.ExternalUrl, $vd.BasicAuthentication, $vd.WindowsAuthentication, $vd.RequireSSL
            } `
            -PreCheckScript {
                $vd = Get-OabVirtualDirectory -Identity $oabIdentity -ErrorAction Stop
                return ([string]$vd.InternalUrl -eq $oabIntUrl -and `
                        [string]$vd.ExternalUrl -eq $oabExtUrl -and `
                        [bool]$vd.BasicAuthentication -eq $oabBasicAuth -and `
                        [bool]$vd.WindowsAuthentication -eq $oabWinAuth -and `
                        $vd.RequireSSL -eq $true)
            } `
            -ActionScript {
                Set-OabVirtualDirectory -Identity $oabIdentity `
                    -InternalUrl $oabIntUrl -ExternalUrl $oabExtUrl `
                    -BasicAuthentication:$oabBasicAuth -WindowsAuthentication:$oabWinAuth -RequireSSL:$true `
                    -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
            }

        # --- EWS ---
        $ewsIdentity  = "$server\EWS (Default Web Site)"
        $ewsIntUrl    = $refEwsInt
        $ewsExtUrl    = $refEwsExt
        $ewsBasicAuth = $refEwsBasicAuth
        $ewsWinAuth   = $refEwsWinAuth
        $ewsOAuthAuth = $refEwsOAuthAuth
        Invoke-Action -Step '03-VDir' -Target $server -Action 'Set-WebServicesVirtualDirectory URLs+Auth' `
            -Detail "$ewsIntUrl | $ewsExtUrl | Basic=$ewsBasicAuth Win=$ewsWinAuth OAuth=$ewsOAuthAuth" `
            -InventoryScript {
                $vd = Get-WebServicesVirtualDirectory -Identity $ewsIdentity -ErrorAction Stop
                "Int={0}; Ext={1}; Basic={2}; Win={3}; OAuth={4}" -f `
                    $vd.InternalUrl, $vd.ExternalUrl, $vd.BasicAuthentication, $vd.WindowsAuthentication, $vd.OAuthAuthentication
            } `
            -PreCheckScript {
                $vd = Get-WebServicesVirtualDirectory -Identity $ewsIdentity -ErrorAction Stop
                return ([string]$vd.InternalUrl -eq $ewsIntUrl -and `
                        [string]$vd.ExternalUrl -eq $ewsExtUrl -and `
                        [bool]$vd.BasicAuthentication -eq $ewsBasicAuth -and `
                        [bool]$vd.WindowsAuthentication -eq $ewsWinAuth -and `
                        [bool]$vd.OAuthAuthentication -eq $ewsOAuthAuth)
            } `
            -ActionScript {
                Set-WebServicesVirtualDirectory -Identity $ewsIdentity `
                    -InternalUrl $ewsIntUrl -ExternalUrl $ewsExtUrl `
                    -BasicAuthentication:$ewsBasicAuth -WindowsAuthentication:$ewsWinAuth -OAuthAuthentication:$ewsOAuthAuth `
                    -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
            }

        # --- ActiveSync ---
        $easIdentity  = "$server\Microsoft-Server-ActiveSync (Default Web Site)"
        $easIntUrl    = $refEasInt
        $easExtUrl    = $refEasExt
        $easBasicAuth = $refEasBasicAuth
        $easWinAuth   = $refEasWinAuth
        Invoke-Action -Step '03-VDir' -Target $server -Action 'Set-ActiveSyncVirtualDirectory URLs+Auth' `
            -Detail "$easIntUrl | $easExtUrl | Basic=$easBasicAuth Win=$easWinAuth" `
            -InventoryScript {
                $vd = Get-ActiveSyncVirtualDirectory -Identity $easIdentity -ErrorAction Stop
                "Int={0}; Ext={1}; Basic={2}; Win={3}" -f `
                    $vd.InternalUrl, $vd.ExternalUrl, $vd.BasicAuthEnabled, $vd.WindowsAuthEnabled
            } `
            -PreCheckScript {
                $vd = Get-ActiveSyncVirtualDirectory -Identity $easIdentity -ErrorAction Stop
                return ([string]$vd.InternalUrl -eq $easIntUrl -and `
                        [string]$vd.ExternalUrl -eq $easExtUrl -and `
                        [bool]$vd.BasicAuthEnabled -eq $easBasicAuth -and `
                        [bool]$vd.WindowsAuthEnabled -eq $easWinAuth)
            } `
            -ActionScript {
                Set-ActiveSyncVirtualDirectory -Identity $easIdentity `
                    -InternalUrl $easIntUrl -ExternalUrl $easExtUrl `
                    -BasicAuthEnabled:$easBasicAuth -WindowsAuthEnabled:$easWinAuth `
                    -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
            }

        # --- Outlook Anywhere ---
        # IISAuthenticationMethods must be passed as an ARRAY (not as a CSV string).
        # ExternalHostname requires ExternalClientsRequireSsl ($true since SSL is used).
        $oaIdentity   = "$server\Rpc (Default Web Site)"
        $oaIntHost    = $refOaIntHost
        $oaExtHost    = $refOaExtHost
        $oaExtAuth    = $refOaExtAuth
        $oaIntAuth    = $refOaIntAuth
        $oaIISAuthArr = @($refOaIISAuth -split '[\s,]+' | Where-Object { $_ })
        $oaSSLOff     = $refOaSSLOff
        Invoke-Action -Step '03-VDir' -Target $server -Action 'Set-OutlookAnywhere URLs+Auth' `
            -Detail "Int=$oaIntHost Ext=$oaExtHost | ExtAuth=$oaExtAuth IntAuth=$oaIntAuth IISAuth=$($oaIISAuthArr -join ',') SSLOff=$oaSSLOff RequireSsl=True" `
            -InventoryScript {
                $oa = Get-OutlookAnywhere -Identity $oaIdentity -ErrorAction Stop
                "ExtHost={0}; IntHost={1}; ExtAuth={2}; IntAuth={3}; IISAuth={4}; SSLOff={5}; ExtSsl={6}; IntSsl={7}" -f `
                    $oa.ExternalHostname, $oa.InternalHostname, `
                    $oa.ExternalClientAuthenticationMethod, $oa.InternalClientAuthenticationMethod, `
                    ([string]$oa.IISAuthenticationMethods), $oa.SSLOffloading, $oa.ExternalClientsRequireSsl, $oa.InternalClientsRequireSsl
            } `
            -PreCheckScript {
                $oa = Get-OutlookAnywhere -Identity $oaIdentity -ErrorAction Stop
                return ([string]$oa.InternalHostname -eq $oaIntHost -and `
                        [string]$oa.ExternalHostname -eq $oaExtHost -and `
                        [string]$oa.ExternalClientAuthenticationMethod -eq $oaExtAuth -and `
                        [string]$oa.InternalClientAuthenticationMethod -eq $oaIntAuth -and `
                        [bool]$oa.SSLOffloading -eq $oaSSLOff -and `
                        [bool]$oa.ExternalClientsRequireSsl -eq $true -and `
                        [bool]$oa.InternalClientsRequireSsl -eq $true)
            } `
            -ActionScript {
                Set-OutlookAnywhere -Identity $oaIdentity `
                    -ExternalHostname $oaExtHost -InternalHostname $oaIntHost `
                    -ExternalClientsRequireSsl $true -InternalClientsRequireSsl $true `
                    -ExternalClientAuthenticationMethod $oaExtAuth `
                    -InternalClientAuthenticationMethod $oaIntAuth `
                    -IISAuthenticationMethods $oaIISAuthArr `
                    -SSLOffloading:$oaSSLOff `
                    -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
            }

        # --- MAPI ---
        # IISAuthenticationMethods must be passed as an ARRAY (not as a CSV string).
        $mapiIdentity   = "$server\mapi (Default Web Site)"
        $mapiIntUrl     = $refMapiInt
        $mapiIISAuthArr = @($refMapiIISAuth -split '[\s,]+' | Where-Object { $_ })
        Invoke-Action -Step '03-VDir' -Target $server -Action 'Set-MapiVirtualDirectory URL+Auth' `
            -Detail "Int=$mapiIntUrl | IISAuth=$($mapiIISAuthArr -join ',')" `
            -InventoryScript {
                $vd = Get-MapiVirtualDirectory -Identity $mapiIdentity -ErrorAction Stop
                "Int={0}; IISAuth={1}; MapiHttpEnabled={2}" -f `
                    $vd.InternalUrl, ([string]$vd.IISAuthenticationMethods), $vd.MapiHttpEnabled
            } `
            -PreCheckScript {
                $vd = Get-MapiVirtualDirectory -Identity $mapiIdentity -ErrorAction Stop
                return ([string]$vd.InternalUrl -eq $mapiIntUrl)
            } `
            -ActionScript {
                Set-MapiVirtualDirectory -Identity $mapiIdentity `
                    -InternalUrl $mapiIntUrl `
                    -IISAuthenticationMethods $mapiIISAuthArr `
                    -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
            }

        # --- Extended Protection ---
        # Fixed Exchange 2019 targets: nothing is read from the legacy source.
        # OWA/ECP/MAPI=Require | OA=Require | EWS/OAB/EAS=Allow | AutoDisc=None
        Invoke-Action -Step '03-VDir' -Target $server -Action 'Set-ExtendedProtection' `
            -Detail 'OWA/ECP/MAPI=Require | OA=Require | EWS/OAB/EAS=Allow | AutoDisc=None' `
            -InventoryScript {
                $parts = @()
                try { $v = Get-OwaVirtualDirectory          -Identity $owaIdentity  -EA Stop; $parts += "OWA=$($v.ExtendedProtectionTokenChecking)"  } catch {}
                try { $v = Get-EcpVirtualDirectory          -Identity $ecpIdentity  -EA Stop; $parts += "ECP=$($v.ExtendedProtectionTokenChecking)"  } catch {}
                try { $v = Get-AutodiscoverVirtualDirectory  -Identity $autoIdentity -EA Stop; $parts += "Auto=$($v.ExtendedProtectionTokenChecking)" } catch {}
                try { $v = Get-OabVirtualDirectory          -Identity $oabIdentity  -EA Stop; $parts += "OAB=$($v.ExtendedProtectionTokenChecking)"  } catch {}
                try { $v = Get-WebServicesVirtualDirectory  -Identity $ewsIdentity  -EA Stop; $parts += "EWS=$($v.ExtendedProtectionTokenChecking)"  } catch {}
                try { $v = Get-ActiveSyncVirtualDirectory   -Identity $easIdentity  -EA Stop; $parts += "EAS=$($v.ExtendedProtectionTokenChecking)"  } catch {}
                try { $v = Get-OutlookAnywhere              -Identity $oaIdentity   -EA Stop; $parts += "OA=$($v.ExtendedProtectionTokenChecking)"   } catch {}
                try { $v = Get-MapiVirtualDirectory         -Identity $mapiIdentity -EA Stop; $parts += "MAPI=$($v.ExtendedProtectionTokenChecking)" } catch {}
                $parts -join '; '
            } `
            -PreCheckScript {
                try {
                    $v = Get-OwaVirtualDirectory          -Identity $owaIdentity  -EA Stop; if ([string]$v.ExtendedProtectionTokenChecking -ne 'Require') { return $false }
                    $v = Get-EcpVirtualDirectory          -Identity $ecpIdentity  -EA Stop; if ([string]$v.ExtendedProtectionTokenChecking -ne 'Require') { return $false }
                    $v = Get-AutodiscoverVirtualDirectory  -Identity $autoIdentity -EA Stop; if ([string]$v.ExtendedProtectionTokenChecking -ne 'None')    { return $false }
                    $v = Get-OabVirtualDirectory          -Identity $oabIdentity  -EA Stop; if ([string]$v.ExtendedProtectionTokenChecking -ne 'Allow')   { return $false }
                    $v = Get-WebServicesVirtualDirectory  -Identity $ewsIdentity  -EA Stop; if ([string]$v.ExtendedProtectionTokenChecking -ne 'Allow')   { return $false }
                    $v = Get-ActiveSyncVirtualDirectory   -Identity $easIdentity  -EA Stop; if ([string]$v.ExtendedProtectionTokenChecking -ne 'Allow')   { return $false }
                    $v = Get-OutlookAnywhere              -Identity $oaIdentity   -EA Stop; if ([string]$v.ExtendedProtectionTokenChecking -ne 'Require') { return $false }
                    $v = Get-MapiVirtualDirectory         -Identity $mapiIdentity -EA Stop; if ([string]$v.ExtendedProtectionTokenChecking -ne 'Require') { return $false }
                    return $true
                } catch { return $false }
            } `
            -ActionScript {
                $isWhatIf = ([bool]($Mode -eq 'Simulate'))
                Set-OwaVirtualDirectory         -Identity $owaIdentity  -ExtendedProtectionTokenChecking Require -WhatIf:$isWhatIf -ErrorAction Stop
                Set-EcpVirtualDirectory         -Identity $ecpIdentity  -ExtendedProtectionTokenChecking Require -WhatIf:$isWhatIf -ErrorAction Stop
                Set-AutodiscoverVirtualDirectory -Identity $autoIdentity -ExtendedProtectionTokenChecking None    -WhatIf:$isWhatIf -ErrorAction Stop
                Set-OabVirtualDirectory         -Identity $oabIdentity  -ExtendedProtectionTokenChecking Allow   -WhatIf:$isWhatIf -ErrorAction Stop
                Set-WebServicesVirtualDirectory -Identity $ewsIdentity  -ExtendedProtectionTokenChecking Allow   -WhatIf:$isWhatIf -ErrorAction Stop
                Set-ActiveSyncVirtualDirectory  -Identity $easIdentity  -ExtendedProtectionTokenChecking Allow   -WhatIf:$isWhatIf -ErrorAction Stop
                Set-OutlookAnywhere             -Identity $oaIdentity   -ExtendedProtectionTokenChecking Require -WhatIf:$isWhatIf -ErrorAction Stop
                Set-MapiVirtualDirectory        -Identity $mapiIdentity -ExtendedProtectionTokenChecking Require -WhatIf:$isWhatIf -ErrorAction Stop
            }
    }

    # --- 4. Sequential IIS restart (one server at a time) ---------------
    # Without iisreset, the Set-ClientAccessService (SCP), Set-*VirtualDirectory
    # (URLs/Auth), Set-OutlookAnywhere and ExtendedProtection changes do NOT take effect:
    # Exchange writes to AD/metabase but the IIS worker keeps running with the old config
    # until the AppPools are reloaded.
    # We operate server by server (sequential) to preserve high availability: while
    # one server resets IIS (~15-30s), the other 3 serve the traffic through the LB.
    foreach ($srv in $servers) {
        $server = $srv.Name

        Invoke-Action -Step '03-VDir' -Target $server -Action 'Restart-IIS (iisreset /noforce)' `
            -Detail 'iisreset /noforce + wait for W3SVC + WAS Running (sequential)' `
            -InventoryScript {
                try {
                    Invoke-Command -ComputerName $server -ScriptBlock {
                        'W3SVC={0}; WAS={1}' -f `
                            (Get-Service W3SVC -EA SilentlyContinue).Status, `
                            (Get-Service WAS   -EA SilentlyContinue).Status
                    } -ErrorAction Stop
                } catch { "Error reading IIS services: $($_.Exception.Message)" }
            } `
            -PreCheckScript $null `
            -ActionScript {
                $isWhatIf = ([bool]($Mode -eq 'Simulate'))
                if ($isWhatIf) {
                    Write-Log "WhatIf: iisreset /noforce on $server (sequential)" -Level Sub
                    return
                }

                # iisreset /noforce runs on the remote server through WinRM.
                # Tolerant strategy:
                #   1. Up to 2 attempts (1 initial + 1 retry if exit != 0)
                #   2. The iisreset output is NOT the success criterion: we rely on the actual
                #      state of the W3SVC + WAS services after the attempts.
                #   3. Typical tolerated case: exit 1053 (SCM timeout on stop) while the
                #      services eventually come back Running through the SCM dependency
                #      chain. iisreset.exe gave up but the service is UP.
                $maxAttempts = 2
                $attempt     = 0
                $lastResult  = $null
                $allOutputs  = @()
                $hadNonZero  = $false

                while ($attempt -lt $maxAttempts) {
                    $attempt++
                    $lastResult = Invoke-Command -ComputerName $server -ScriptBlock {
                        $output = & iisreset /noforce 2>&1
                        [PSCustomObject]@{ ExitCode = $LASTEXITCODE; Output = ($output -join ' ; ') }
                    } -ErrorAction Stop
                    $allOutputs += ("[try {0}/{1} exit={2}] {3}" -f $attempt, $maxAttempts, $lastResult.ExitCode, $lastResult.Output)
                    Write-Log ("iisreset {0} attempt {1}/{2}: exit={3}" -f $server, $attempt, $maxAttempts, $lastResult.ExitCode) -Level Sub
                    if ($lastResult.ExitCode -ne 0) { $hadNonZero = $true }
                    if ($lastResult.ExitCode -eq 0) { break }
                    if ($attempt -lt $maxAttempts) {
                        Write-Log ("iisreset {0} returned exit {1} - retrying in 5s..." -f $server, $lastResult.ExitCode) -Level Warning
                        Start-Sleep -Seconds 5
                    }
                }

                # Effective success criterion: W3SVC + WAS services Running (60s timeout).
                # Independent of the iisreset exit code. The actual state is what matters.
                $deadline = (Get-Date).AddSeconds(60)
                $w3 = $was = $null
                do {
                    Start-Sleep -Seconds 2
                    $status = Invoke-Command -ComputerName $server -ScriptBlock {
                        '{0}|{1}' -f `
                            (Get-Service W3SVC -EA SilentlyContinue).Status, `
                            (Get-Service WAS   -EA SilentlyContinue).Status
                    } -ErrorAction SilentlyContinue
                    if ($status) {
                        $w3, $was = $status -split '\|'
                        if ($w3 -eq 'Running' -and $was -eq 'Running') { break }
                    }
                } while ((Get-Date) -lt $deadline)

                if ($w3 -ne 'Running' -or $was -ne 'Running') {
                    throw ("Services not Running after {0} iisreset attempt(s) (60s timeout) on {1}: W3SVC={2}; WAS={3}. Attempts: {4}" -f $attempt, $server, $w3, $was, ($allOutputs -join ' | '))
                }

                # Services Running. If an exit != 0 was observed (typical SCM timeout 1053
                # case), a warning is logged but the step is NOT failed:
                # the goal is reached (services reloaded, SCP/VDir/EP changes active).
                if ($hadNonZero) {
                    Write-Log ("iisreset {0}: non-zero exit tolerated - W3SVC + WAS Running. Attempts: {1}" -f $server, ($allOutputs -join ' | ')) -Level Warning
                }
            }
    }

    Save-Report -StepName 'Step03-VirtualDirectories'
    return 0
}
