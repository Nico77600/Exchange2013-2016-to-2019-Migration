#
#  Exchange 2013/2016 to 2019 Migration - configuration
#  -----------------------------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : 2.0.0
#
#  Read by Deploy-Exchange2019.ps1 (Import-PowerShellDataFile) and passed to every step as $Config.
#  Everything specific to an environment is here, with DiskLayout.csv and DAGInfo.csv: moving from
#  the lab to production (or to another organisation) never needs a change in the scripts.
#  PowerShell data file: text between quotes, $true / $false, numbers, @( ) for lists, # for
#  comments. Guide, chapter 6.
#
@{
    # --- External and internal URLs (virtual directories, Step 03) ---------------------------------
    UrlExterne     = ''
    UrlInterne     = ''

    # --- NetBIOS domain name used by OWA ------------------------------------------------------------
    DomainNetBios  = ''

    # --- Exchange 2019 product key (Step 01) --------------------------------------------------------
    LicenseKey     = ''

    # --- Certificate exported from Exchange 2013/2016 (Step 02) ------------------------------------
    # AUTO mode: the source server is picked among the legacy Mailbox servers of the organisation;
    # the certificate kept is the one bound to all the services IIS, SMTP, POP and IMAP.
    # MANUAL mode: a PFX file placed beforehand at CertificateExportPath.
    CertificateExportPath       = 'D:\sources\exchange2019\exportcert2013a.pfx'
    # The password of the PFX is asked interactively (or given with -PfxPassword).

    CertificateServices         = 'IIS,POP,SMTP,IMAP'

    # --- Source server of the receive connectors (2013 or 2016, Step 07) ---------------------------
    # Step 07 uses SourceServer2016ForConnectors first when set, then SourceServer2013ForConnectors,
    # then the first legacy server found.
    SourceServer2013ForConnectors = ''
    SourceServer2016ForConnectors = ''

    # --- Source server of the virtual directories (2013 or 2016, Step 03) --------------------------
    # Step 03 reads the URLs (InternalUrl / ExternalUrl) of this model server.
    # Empty: the first Exchange 2013/2016 Mailbox server found.
    # No legacy server available: UrlInterne / UrlExterne above.
    SourceServerForVDirs = ''

    # --- Transport queue folder (Step 05) -----------------------------------------------------------
    TransportQueueRoot          = 'Q:\Queue'

    # --- Exchange 2019 log paths (Step 06) ----------------------------------------------------------
    LogPaths = @{
        FrontEndTransportLogs   = 'D:\Logs\FrontendTransport'
        TransportLogs           = 'D:\Logs\Transport'
        MailboxDeliveryLogs     = 'D:\Logs\MailboxDelivery'
        IISLogsDefaultWebSite   = 'D:\Logs\IIS\DefaultWebSite'
        IISLogsBackEnd          = 'D:\Logs\IIS\ExchangeBackEnd'
        Pop3Logs                = 'D:\Logs\Pop3'
        Imap4Logs               = 'D:\Logs\Imap4'
    }

    # --- Kerberos alternate service account (Step 15) -----------------------------------------------
    # Microsoft procedure:
    # https://learn.microsoft.com/en-us/exchange/architecture/client-access/kerberos-auth-for-load-balanced-client-access
    KerberosASA = @{
        # Short name of the computer account (without the final $)
        AccountName     = ''
        # Active Directory domain (FQDN)
        Domain          = ''
        # NetBIOS name of the domain (for DOMAIN\ACCOUNT$)
        DomainNetBios   = ''
        # OU of the computer account (distinguished name). Empty: the default Computers container
        # (CN=Computers,DC=contoso,DC=local). Recommended: an OU for the Exchange service accounts.
        OUPath          = ''
        # Active Directory description
        Description     = 'Alternate Service Account credentials for Exchange 2019'
        # SPNs of the account: the shared names behind the load balancer (not the server names).
        # IMPORTANT: these SPNs must not be registered on ANY other account.
        SPNs            = @()
    }

    # --- Migration quota ----------------------------------------------------------------------------
    MigrationQuotaGB     = 150

    # --- Mailbox database layout (generator of Step 10) ---------------------------------------------
    # Step 10 builds the database names and paths from the pattern below instead of reading
    # MailboxDatabases.csv.
    #
    # Mode = 'Generator' : the settings below are used
    # Mode = 'Csv'       : MailboxDatabases.csv is read (historical mode)
    #
    # In both modes Step 10 creates the target folders REMOTELY (Invoke-Command on the active server)
    # BEFORE New-MailboxDatabase.
    DatabaseLayout = @{
        Mode                = 'Generator'                       # 'Generator' or 'Csv'

        # --- Names ----------------------------------------------------------------------------------
        DatabasePrefix      = 'DB'                              # DB01, DB02, ...
        DatabaseDigits      = 2                                 # 2 -> DB01..DB99 (3 -> DB001..)
        DatabaseStartIndex  = 1                                 # First number
        DatabaseCount       = 4                                 # Number of databases to create

        # --- Paths ----------------------------------------------------------------------------------
        # Templates: {Drive} -> DriveLetter, {Name} -> database name (e.g. DB01)
        DriveLetter         = 'M'
        EdbPathTemplate     = '{Drive}:\Databases\{Name}\{Name}.edb'  # M:\Databases\DB01\DB01.edb
        LogPathTemplate     = '{Drive}:\Databases\{Name}\Logs'        # M:\Databases\DB01\Logs

        # Distribution of the databases on the Exchange 2019 servers:
        #   empty => round-robin on every Exchange 2019 server found
        #   list  => round-robin on this explicit list (NetBIOS name or FQDN)
        Servers             = @()

        # --- Properties set on every generated database --------------------------------------------
        DefaultProperties = @{
            ProhibitSendReceiveQuota          = '150GB'
            ProhibitSendQuota                 = '149GB'
            IssueWarningQuota                 = '148GB'
            DeletedItemRetention              = '30.00:00:00'
            MailboxRetention                  = '90.00:00:00'
            OfflineAddressBook                = $null            # null -> default OAB
            IndexEnabled                      = $true
            BackgroundDatabaseMaintenance     = $true
            MountAtStartup                    = $true
            AllowFileRestore                  = $false
            RetainDeletedItemsUntilBackup     = $false
            IsExcludedFromInitialProvisioning = $false
            IsExcludedFromProvisioning        = $false
            IsSuspendedFromProvisioning       = $false
            AutoDagExcludeFromMonitoring      = $false
            EventHistoryRetentionPeriod       = '7.00:00:00'
            CalendarLoggingQuota              = '6GB'
        }
    }

    # --- Migration (Steps 19 to 21) -----------------------------------------------------------------
    Migration = @{
        BatchCount         = 12
        BadItemLimit       = 20
        LargeItemLimit     = 0
        AcceptLargeDataLoss = $false
    }

    # --- Anti-malware check (Step 13) ---------------------------------------------------------------
    Antimalware = @{
        EventLogName    = 'Application'
        UpdateOK_EventID = 6033          # FIPFS update succeeded
        UpdateKO_EventID = 6027          # update attempted without success
        EventSource      = 'Microsoft-Filtering-FIPFS'
        SearchHours      = 24
    }

    # --- Event logs (Step 14) -----------------------------------------------------------------------
    EventLogs = @{
        TargetSizeBytes = 1GB
        Logs            = @('Application','System')
    }

    # --- Domain controllers -------------------------------------------------------------------------
    PreferredGCs = @()

    # --- DAG (Step 08) ------------------------------------------------------------------------------
    # Fixed values applied to every DAG.
    # SafetyNetHoldTime must cover the ReplayLagTime of the lagged copies.
    DAG = @{
        DACMode            = 'DagOnly'   # DAC mandatory (IP-less DAG)
        NetworkEncryption  = 'Enabled'   # encryption within and between sites
        NetworkCompression = 'Enabled'   # compression within and between sites
        SafetyNetHoldTime  = 8           # days (>= ReplayLagTime of the copies)
    }

    # --- Database copies (Step 11) ------------------------------------------------------------------
    # Settings of the lagged copies (preference 4, other site, next index).
    # The passive copies (preferences 2 and 3) have no lag.
    DatabaseCopies = @{
        ReplayLagTime     = '7.00:00:00'   # replay lag (1 week)
        TruncationLagTime = '0.00:00:00'
        ReplayLagMaxDelay = '1.00:00:00'   # safety-net play-down (1 day)
    }

    # --- IIS log management (Step 17) ---------------------------------------------------------------
    # Deploys Manage-IISLogs.ps1 and a scheduled task on every Exchange 2019 server.
    # LogPath = PARENT folder of both IIS paths (IISLogsDefaultWebSite / IISLogsBackEnd),
    # e.g. D:\Logs\IIS. The real root is detected at run time; this value is the fallback.
    IISLogsManagement = @{
        LogPath                = 'D:\Logs\IIS'
        ScheduledTaskName      = 'Manage IIS Logs - Exchange'
        ScriptDeployPath       = 'C:\Scripts\Manage-IISLogs.ps1'
        RunDay                 = 'Sunday'
        RunTime                = '02:00'
        RetainUncompressedDays = 7
        PurgeAfterDays         = 180
    }

    # --- HealthChecker (Step 24) --------------------------------------------------------------------
    # Runs the Microsoft Exchange HealthChecker.ps1 on every Exchange 2019 server, then builds a
    # report limited to the warnings and errors (the official HealthChecker report keeps the OK items).
    #
    # ScriptPath        : path of HealthChecker.ps1 (relative to the tool folder, or absolute).
    # SkipVersionCheck  : passes -SkipVersionCheck (no online update check).
    # BuildOfficialHtml : also runs -BuildHtmlServersReport to build the official HealthChecker
    #                     report (every status) next to the custom report.
    HealthCheckerReport = @{
        ScriptPath        = 'Configs\HealthChecker.ps1'
        SkipVersionCheck  = $true
        BuildOfficialHtml = $true
    }

    # --- Quotas after the migration (Step 23) -------------------------------------------------------
    # Applied by Step 23 once the migration is complete. The databases are created with large
    # migration quotas (Step 10) so that no move is blocked.
    PostMigrationQuotas = @{
        IssueWarningQuota        = '4GB'    # 4 GB: warning
        ProhibitSendQuota        = '4608MB' # 4.5 GB: send blocked
        ProhibitSendReceiveQuota = '5GB'    # 5 GB: send and receive blocked
    }

    # --- Protocol-log analysis (Step 26) ------------------------------------------------------------
    # Step26-AnalyzeProtocolLogs reads the IIS logs and the SMTP protocol logs (Receive + Send, NOT
    # message tracking) to compare the real-user traffic of the legacy servers and of Exchange 2019,
    # and to confirm that the clients have moved to Exchange 2019.
    #
    # DefaultHoursWindow : rolling window in hours, ending now. Override for one run with
    #   $env:LOG_HOURS = <n> (whole number).
    # IISVdirsToTrack    : IIS virtual directories counted (match on the first segment of
    #   cs-uri-stem, case-insensitive).
    # SmtpAutoEnableProtocolLogs : in Apply mode, a receive/send connector with
    #   ProtocolLoggingLevel 'None' is set to 'Verbose'. $false: only reported.
    #   Inventory never changes anything.
    LogAnalysis = @{
        DefaultHoursWindow         = 24
        IISVdirsToTrack            = @(
            'Mapi','OWA','EWS','OAB','ECP','RPC','Autodiscover',
            'Microsoft-Server-ActiveSync','PowerShell'
        )
        SmtpAutoEnableProtocolLogs = $true

        # --- Filters: Managed Availability probes, monitoring and system objects -------------------
        # Only REAL user traffic is counted: no Exchange probe (Health Manager, Managed
        # Availability), no system mailbox (HealthMailbox*, SystemMailbox{...}, ...). Annex B.

        # IIS: a request whose cs-uri-stem matches one of these regular expressions is excluded.
        IISExcludeUriPatterns        = @(
            '/healthcheck\.htm$'           # health check page of every virtual directory
        )
        # IIS: a request whose cs-username matches one of these regular expressions is excluded.
        # Tested on the full value, on the part after '\' (DOMAIN\User -> User) and before '@'.
        IISExcludeUserPatterns       = @(
            '^HealthMailbox'               # Managed Availability mailboxes
            '^extest_'                     # extest mailboxes (external probe)
            '^DiscoverySearchMailbox'      # discovery search mailbox
            '^SystemMailbox\{'             # system mailboxes
            '^FederatedEmail\.'            # federation mailbox
            '^Migration\.[0-9a-fA-F]'      # migration mailbox
            '^(OABGen|OfflineAddressBook)'  # OAB generation objects
            '^Microsoft\.Exchange'
            '^IISAPPPOOL\\'                # IIS application pool identity
            '^IIS APPPOOL\\'              # IIS application pool identity
            '^IUSR($|_)'                   # IIS anonymous account
            '^(AMProbe|ManagedAvailability|Microsoft\.Exchange\.Monitoring)'
            '^(HealthService|MailboxLoadBalancer|MSExchange|SMTPSVC|W3SVC)'
            '^(LOCAL SYSTEM|LOCAL SERVICE|NETWORK SERVICE|ANONYMOUS LOGON)$'
        )
        # IIS: strict mode, a request without an authenticated identity is never user traffic.
        IISRequireAuthenticatedUser  = $true
        # IIS: a request whose cs(User-Agent) contains one of these strings is excluded
        # (case-insensitive "contains", not a regular expression).
        IISExcludeUserAgentPatterns  = @(
            'AMProbe'                       # Active Monitoring probe
            'MSRPC'                         # internal Exchange RPC tests
            'ASProxy'                       # internal ActiveSync proxy
            'ManagedAvailability'           # Managed Availability framework
            'MailboxLoadBalancer'           # internal load balancer health checks
            'Microsoft.Exchange.Monitoring' # monitoring framework
            'HealthService'                 # SCOM / Health Service
            'Health Manager'                # MSExchange Health Manager
            'ExchangeActiveMonitoring'
            'Test-OAuthConnectivity'
            'Test-MapiConnectivity'
            'Test-EwsConnectivity'
            'MonitoringProbe'
        )

        # SMTP: a session is "valid" only when the client sent at least one 'MAIL FROM:' command
        # (otherwise it is a probe: EHLO/QUIT, NOOP...). The report shows total and valid sessions.
        SmtpRequireMailFromForValid = $true
        # SMTP: a session whose MAIL FROM is a system mailbox, a monitoring agent or empty is not
        # user activity.
        SmtpExcludeSenderPatterns = @(
            '^HealthMailbox'
            '^SystemMailbox\{'
            '^FederatedEmail\.'
            '^DiscoverySearchMailbox'
            '^Migration\.[0-9a-fA-F]'
            '^(OABGen|OfflineAddressBook)'
            '^Microsoft\.Exchange'
            '^(AMProbe|ManagedAvailability|Microsoft\.Exchange\.Monitoring)'
            '^(HealthService|MSExchange|SMTPSVC|W3SVC)'
            '^(postmaster|mailer-daemon)(@|$)'
        )
        SmtpRequireSenderAddress = $true

        # IIS: only the requests with an sc-status 2xx (authentication completed) are counted.
        # Without this filter, the 401 challenges of the Negotiate/NTLM round trip are counted as
        # user traffic: one real Outlook operation = 2 IIS lines (401 then 200), counted twice.
        # Worse: the HealthMailbox / AMProbe probes send a 401 WITHOUT cs-username, with the user
        # agent MapiHttpClient (a real MAPI client), which passes every other filter.
        # Default: $true. $false shows every status.
        IISRequireSuccessStatus = $true
    }
}
