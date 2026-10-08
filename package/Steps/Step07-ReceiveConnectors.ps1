<#
.SYNOPSIS
    Step 07 - Copy the Receive Connectors from an Exchange 2013 or 2016 server to the 2019 servers.

.DESCRIPTION
    Copies the custom Receive Connectors of a legacy Exchange 2013/2016 server to every
    Exchange 2019 server.

    Source: resolved in the following priority order:
        1. Config.SourceServer2016ForConnectors  (if defined and not empty)
        2. Config.SourceServer2013ForConnectors  (if defined and not empty)
        3. Auto-detection: first Exchange 2016 server detected in the organization
        4. Auto-detection: first Exchange 2013 server detected in the organization
    Targets: all Exchange 2019 servers (filtered through AdminDisplayVersion).

    For each source connector, the script copies:
        - The main parameters (Bindings, RemoteIPRanges, AuthMechanism, PermissionGroups,
          MaxMessageSize, ConnectionTimeout, AdvertiseClientSettings, BannerString,
          MessageRateLimit, MaxInboundConnection, etc.)
        - The CUSTOM, NON-inherited, explicitly defined AD permissions (Get-ADPermission ... |
          Where-Object { $_.IsInherited -eq $false }) - filtered to exclude the standard
          Microsoft permissions.

    Idempotency:
        - If a Receive Connector with the same name already exists on the target and its
          configuration is identical: AlreadyDone
        - Otherwise: delete / re-create with identical parameters + add the AD permissions

    Important note: never touches the connectors of the source 2013/2016 servers.

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

    Initialize-ExchangeShell

    # Source server resolution: priority 2016 > 2013 > auto-detection
    $sourceServer = $null
    $sourceVersion = $null

    if ($Config.SourceServer2016ForConnectors) {
        $sourceServer  = $Config.SourceServer2016ForConnectors
        $sourceVersion = '2016'
    } elseif ($Config.SourceServer2013ForConnectors) {
        $sourceServer  = $Config.SourceServer2013ForConnectors
        $sourceVersion = '2013'
    } else {
        # Auto-detection: Exchange 2016 first, then 2013
        $detected2016 = Get-Exchange2016Servers | Select-Object -First 1
        if ($detected2016) {
            $sourceServer  = $detected2016.Name
            $sourceVersion = '2016 (auto-detected)'
        } else {
            $detected2013 = Get-Exchange2013Servers | Select-Object -First 1
            if ($detected2013) {
                $sourceServer  = $detected2013.Name
                $sourceVersion = '2013 (auto-detected)'
            }
        }
    }

    Write-StepBanner -StepName '07' -Title 'Copy of the Receive Connectors from Exchange 2013/2016' -Mode $Mode -Actions @(
        ('Source: {0} [Exchange {1}]' -f $sourceServer, $sourceVersion),
        'For each non-default source Receive Connector:',
        '  - Read its full parameters (Bindings, RemoteIPRanges, Auth, PermissionGroups, ...)',
        '  - Retrieve the Get-ADPermission -InheritedObjectAccess, only non-inherited / explicit ones',
        '  - For each 2019 server: create the connector if missing (otherwise compare then Set-)',
        '  - Re-apply the custom non-inherited Add-ADPermission entries',
        'No connector of the source 2013/2016 server will be modified.'
    )

    if (-not $sourceServer) {
        Write-Log "No Exchange 2013/2016 server found (config and auto-detection)" -Level Error
        Add-Report -Step '07-RC' -Target '<discovery>' -Action 'ResolveSourceServer' -Status Failed `
                   -ErrorMessage 'SourceServer2016ForConnectors and SourceServer2013ForConnectors not defined, no legacy server detected'
        return 2
    }

    Write-Log "Receive Connectors source server: $sourceServer [Exchange $sourceVersion]" -Level Info

    # Gets the connectors of the 2013 server, excluding the default connectors (Default <SRV>, Client Frontend, etc.
    # kept by Exchange) - the short name "Default" is often intentionally not copied to avoid conflicts.
    # We choose to copy ALL custom Receive Connectors, i.e. those that are not named
    # exactly "Default", "Default Frontend", "Client Frontend", "Outbound Proxy Frontend", "Client Proxy".
    $defaultNames = @('Default','Default Frontend','Client Frontend','Outbound Proxy Frontend','Client Proxy')

    $sourceRcs = Get-ReceiveConnector -Server $sourceServer -ErrorAction Stop |
                 Where-Object {
                    # ShortName of the connector after the backslash
                    $name = $_.Identity.ToString().Split('\')[-1].Trim()
                    $isDefault = $false
                    foreach ($d in $defaultNames) {
                        if ($name -like "$d*") { $isDefault = $true; break }
                    }
                    -not $isDefault
                 }

    if (-not $sourceRcs -or $sourceRcs.Count -eq 0) {
        Write-Log "No custom Receive Connector found on $sourceServer" -Level Warning
        Add-Report -Step '07-RC' -Target $sourceServer -Action 'EnumerateConnectors' -Status Skipped `
                   -Detail 'No custom connector'
        return 0
    }

    Add-Report -Step '07-RC' -Target $sourceServer -Action 'EnumerateConnectors' -Status Inventoried `
               -Phase 'Before' -BeforeValue ("Connectors to copy: " + ($sourceRcs.Name -join ', '))

    # Targets: 2019 only
    $targets = Get-Exchange2019Servers
    if (-not $targets -or $targets.Count -eq 0) {
        Write-Log "No Exchange 2019 server detected." -Level Warning
        return 0
    }

    foreach ($rc in $sourceRcs) {
        $rcName = $rc.Name

        # Gets the non-inherited AD permissions, excluding the system/Exchange principals
        # managed automatically by the schema. WE KEEP 'NT AUTHORITY\ANONYMOUS LOGON' which
        # is the typical pattern of a custom anonymous relay (cf. Apply-LabConfig.ps1 prod).
        # We ONLY EXCLUDE:
        #   - the built-in Exchange groups (Exchange Trusted Subsystem, Mailbox Servers...)
        #   - the Windows system principals (SYSTEM, NETWORK SERVICE, BUILTIN\Administrators)
        #   - 'Authenticated Users' (generic system group, never custom)
        # IMPORTANT: via Exchange remoting, $rc.Identity is serialized as a string
        # "Server\Name" which is NOT resolvable by Get-ADPermission (AD searches in
        # the Domain partition whereas the Receive Connectors are in the Configuration NC).
        # Use $rc.DistinguishedName (full DN in the Configuration NC), which works.
        $rcAdIdentity = if ($rc.DistinguishedName) { $rc.DistinguishedName } else { $rc.Identity }
        $customPerms = @()
        try {
            $customPerms = Get-ADPermission -Identity $rcAdIdentity -ErrorAction Stop |
                           Where-Object {
                               -not $_.IsInherited -and
                               $_.User -notlike '*\Exchange Servers' -and
                               $_.User -notlike '*\Exchange Trusted Subsystem' -and
                               $_.User -notlike '*\Exchange Online-ApplicationAccount' -and
                               $_.User -notlike '*\Organization Management' -and
                               $_.User -notlike '*\Hub Transport Servers' -and
                               $_.User -notlike '*\Edge Transport Servers' -and
                               $_.User -notlike '*\Mailbox Servers' -and
                               $_.User -notlike '*\Managed Availability Servers' -and
                               $_.User -notlike '*\Externally Secured Servers' -and
                               $_.User -notlike '*\Authenticated Users' -and
                               $_.User -ne     'NT AUTHORITY\SYSTEM' -and
                               $_.User -ne     'NT AUTHORITY\NETWORK SERVICE' -and
                               $_.User -notlike 'BUILTIN\*' -and
                               $_.User -notlike 'S-1-*'
                           }
        } catch { $customPerms = @() }

        $permsSummary = if ($customPerms.Count -eq 0) {
            '0 custom permissions'
        } else {
            ($customPerms | ForEach-Object { "{0}:{1}" -f $_.User, ($_.ExtendedRights -join '|') }) -join '; '
        }

        foreach ($target in $targets) {
            $serverName = $target.Name

            Invoke-Action -Step '07-RC' -Target "$serverName\$rcName" -Action 'Recreate ReceiveConnector' -Detail $permsSummary `
                -InventoryScript {
                    $existing = Get-ReceiveConnector -Identity "$serverName\$rcName" -ErrorAction SilentlyContinue
                    if ($existing) {
                        "Bindings={0}; RemoteIP={1}; AuthMech={2}; PermGroups={3}" -f `
                            ($existing.Bindings -join ','), ($existing.RemoteIPRanges -join ','), $existing.AuthMechanism, ($existing.PermissionGroups -join '+')
                    } else {
                        '<absent>'
                    }
                } `
                -PreCheckScript {
                    $existing = Get-ReceiveConnector -Identity "$serverName\$rcName" -ErrorAction SilentlyContinue
                    if (-not $existing) { return $false }
                    # Broad but correct comparison
                    return (
                        ((@($existing.Bindings | Sort-Object) -join ',') -eq (@($rc.Bindings | Sort-Object) -join ',')) -and
                        ((@($existing.RemoteIPRanges | Sort-Object) -join ',') -eq (@($rc.RemoteIPRanges | Sort-Object) -join ',')) -and
                        ($existing.AuthMechanism -eq $rc.AuthMechanism) -and
                        ((@($existing.PermissionGroups | Sort-Object) -join ',') -eq (@($rc.PermissionGroups | Sort-Object) -join ','))
                    )
                } `
                -ActionScript {
                    # Remove the existing connector if present (to start clean)
                    $existing = Get-ReceiveConnector -Identity "$serverName\$rcName" -ErrorAction SilentlyContinue
                    if ($existing) {
                        Remove-ReceiveConnector -Identity $existing.Identity -Confirm:$false -WhatIf:([bool]($Mode -eq 'Simulate')) -ErrorAction Stop
                    }

                    # PermissionGroups: 'Custom' is NOT acceptable for
                    # New-ReceiveConnector -PermissionGroups (the Custom value
                    # is set automatically by Exchange when custom Add-ADPermission
                    # entries are added). Custom/None are filtered out.
                    # Note: via Exchange remoting, PermissionGroups can come back
                    # as a CSV string "Custom, AnonymousUsers" rather than an array,
                    # hence the explicit split/trim.
                    $rawPermGroups = @()
                    foreach ($pg in @($rc.PermissionGroups)) {
                        $rawPermGroups += ([string]$pg -split '[,;\s]+')
                    }
                    $cleanPermGroups = @(
                        $rawPermGroups |
                            ForEach-Object { $_.Trim() } |
                            Where-Object { $_ -and $_ -ne 'Custom' -and $_ -ne 'None' } |
                            Select-Object -Unique
                    )

                    # Creation of the new connector with the same parameters
                    $newParams = @{
                        Name             = $rcName
                        Server           = $serverName
                        Bindings         = $rc.Bindings
                        RemoteIPRanges   = $rc.RemoteIPRanges
                        AuthMechanism    = $rc.AuthMechanism
                        TransportRole    = $rc.TransportRole
                        Usage            = 'Custom'
                        ErrorAction      = 'Stop'
                    }
                    if ($cleanPermGroups.Count -gt 0) {
                        $newParams['PermissionGroups'] = $cleanPermGroups
                    }
                    $new = New-ReceiveConnector @newParams -WhatIf:([bool]($Mode -eq 'Simulate'))

                    # Application of the other relevant parameters (after creation)
                    if ($Mode -ne 'Simulate') {
                        Set-ReceiveConnector -Identity "$serverName\$rcName" `
                            -Banner                       $rc.Banner `
                            -ConnectionTimeout            $rc.ConnectionTimeout `
                            -ConnectionInactivityTimeout  $rc.ConnectionInactivityTimeout `
                            -MaxInboundConnection         $rc.MaxInboundConnection `
                            -MaxInboundConnectionPercentagePerSource $rc.MaxInboundConnectionPercentagePerSource `
                            -MaxInboundConnectionPerSource $rc.MaxInboundConnectionPerSource `
                            -MaxMessageSize               $rc.MaxMessageSize `
                            -MaxRecipientsPerMessage      $rc.MaxRecipientsPerMessage `
                            -MessageRateLimit             $rc.MessageRateLimit `
                            -SizeEnabled                  $rc.SizeEnabled `
                            -AdvertiseClientSettings      $rc.AdvertiseClientSettings `
                            -RequireTLS                   $rc.RequireTLS `
                            -EnableAuthGSSAPI             $rc.EnableAuthGSSAPI `
                            -ProtocolLoggingLevel         $rc.ProtocolLoggingLevel `
                            -ErrorAction SilentlyContinue
                    }

                    # Re-application of the non-inherited custom AD permissions
                    foreach ($perm in $customPerms) {
                        Add-ADPermission -Identity "$serverName\$rcName" `
                            -User $perm.User `
                            -ExtendedRights $perm.ExtendedRights `
                            -ErrorAction SilentlyContinue `
                            -WhatIf:([bool]($Mode -eq 'Simulate')) | Out-Null
                    }
                }
        }
    }

    return 0
}
