<#
.SYNOPSIS
    Step 16 - Migrate to MAPI over HTTPS only for the users hosted on Exchange 2019.

.DESCRIPTION
    How it works:
        - The organization already has MapiHttpEnabled=$true (otherwise it is
          enabled).
        - ONLY the mailboxes whose database is on an Exchange 2019 server are
          targeted: Set-CASMailbox -MapiHttpEnabled $true
                    + Set-CASMailbox -MAPIBlockOutlookRpcHttp $true
        - Mailboxes still on Exchange 2013/2016 are not modified.

    Idempotency: if MapiHttpEnabled=$true and MAPIBlockOutlookRpcHttp=$true
                 -> AlreadyDone.

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

    Write-StepBanner -StepName '16' -Title 'MAPI over HTTPS - Exchange 2019 users' -Mode $Mode -Actions @(
        "Set-OrganizationConfig -MapiHttpEnabled `$true (if not already enabled)",
        'Identify the mailboxes hosted on the Exchange 2019 servers',
        'Set-CASMailbox -MapiHttpEnabled $true on these mailboxes',
        'Set-CASMailbox -MAPIBlockOutlookRpcHttp $true (force MAPI/HTTP)',
        'Do NOT touch the mailboxes still on Exchange 2013/2016'
    )

    Reset-Report
    Initialize-ExchangeShell -Credential $Credential

    # --- 1) MapiHttpEnabled at organization level ------------------------
    Invoke-Action -Step '16-MAPI' -Target 'OrganizationConfig' -Action 'Set-OrganizationConfig -MapiHttpEnabled $true' `
        -InventoryScript {
            $o = Get-OrganizationConfig -ErrorAction Stop
            "MapiHttpEnabled=$($o.MapiHttpEnabled)"
        } `
        -PreCheckScript {
            $o = Get-OrganizationConfig -ErrorAction Stop
            return [bool]$o.MapiHttpEnabled
        } `
        -ActionScript {
            Set-OrganizationConfig -MapiHttpEnabled $true -ErrorAction Stop
        }

    # --- 2) Retrieval of the mailboxes hosted on 2019 --------------------
    $ex2019Names = (Get-Exchange2019Servers).Name
    if (-not $ex2019Names) {
        Write-Log "No Exchange 2019 server detected." -Level Warning
        Save-Report -StepName 'Step16-MAPIOverHTTPS'
        return 0
    }

    Write-Log "Retrieving the mailboxes hosted on: $($ex2019Names -join ', ')" -Level Sub
    $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
                 Where-Object { $_.ServerName -in $ex2019Names }

    Write-Log "Total mailboxes to configure: $($mailboxes.Count)" -Level Sub

    foreach ($m in $mailboxes) {
        $alias = $m.Alias

        Invoke-Action -Step '16-MAPI' -Target $alias -Action 'Set-CASMailbox MAPI/HTTP only' `
            -Detail "Server=$($m.ServerName)" `
            -InventoryScript {
                $c = Get-CASMailbox -Identity $alias -ErrorAction Stop
                "MapiHttpEnabled=$($c.MapiHttpEnabled) | BlockRpcHttp=$($c.MAPIBlockOutlookRpcHttp)"
            } `
            -PreCheckScript {
                $c = Get-CASMailbox -Identity $alias -ErrorAction Stop
                return ($c.MapiHttpEnabled -eq $true -and $c.MAPIBlockOutlookRpcHttp -eq $true)
            } `
            -ActionScript {
                Set-CASMailbox -Identity $alias `
                    -MapiHttpEnabled $true `
                    -MAPIBlockOutlookRpcHttp $true `
                    -ErrorAction Stop
            }
    }

    Save-Report -StepName 'Step16-MAPIOverHTTPS'
    return 0
}
