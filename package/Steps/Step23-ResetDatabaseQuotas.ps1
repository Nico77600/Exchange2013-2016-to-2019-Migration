<#
.SYNOPSIS
    Step 23 - Reset mailbox quotas after migration.

.DESCRIPTION
    The databases were created with large quotas to ease the migration
    (Step10: 148/149/150 GB). This step reapplies the standard production
    quotas defined in Config.PostMigrationQuotas.

    Applied to all Exchange 2019 databases of the organization.
    Idempotent: the InventoryScript reads the current values for audit.

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

    $quotas = $Config.PostMigrationQuotas
    $warn   = $quotas.IssueWarningQuota
    $send   = $quotas.ProhibitSendQuota
    $recv   = $quotas.ProhibitSendReceiveQuota
    $gc     = $Config.PreferredGCs | Select-Object -First 1

    Write-StepBanner -StepName '23' -Title 'Reset mailbox quotas' -Mode $Mode -Actions @(
        "IssueWarningQuota        : $warn",
        "ProhibitSendQuota        : $send",
        "ProhibitSendReceiveQuota : $recv",
        "Applies to all Exchange 2019 databases in the organization"
    )

    Initialize-ExchangeShell -Credential $Credential

    $dbs = Get-MailboxDatabase -ErrorAction Stop |
           Where-Object { Test-IsExchange2019Server -ServerName $_.Server.Name }

    if (-not $dbs) {
        Write-Log "No Exchange 2019 database found" -Level Warning
        Save-Report -StepName 'Step23-ResetDatabaseQuotas'
        return 0
    }

    foreach ($db in $dbs) {
        $dbName = $db.Name

        Invoke-Action -Step '23-Quotas' -Target $dbName -Action 'Set-MailboxDatabase (quotas)' `
            -Detail "Warn=$warn | Send=$send | Recv=$recv" `
            -InventoryScript {
                $b = Get-MailboxDatabase -Identity $dbName -ErrorAction Stop
                "Warn={0}; Send={1}; Recv={2}" -f $b.IssueWarningQuota, $b.ProhibitSendQuota, $b.ProhibitSendReceiveQuota
            } `
            -PreCheckScript $null `
            -ActionScript {
                Set-MailboxDatabase -Identity $dbName `
                    -IssueWarningQuota        $warn `
                    -ProhibitSendQuota        $send `
                    -ProhibitSendReceiveQuota $recv `
                    -DomainController         $gc `
                    -ErrorAction Stop -WhatIf:([bool]($Mode -eq 'Simulate'))
            }
    }

    Save-Report -StepName 'Step23-ResetDatabaseQuotas'
    return 0
}
