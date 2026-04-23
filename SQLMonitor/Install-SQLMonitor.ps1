[CmdletBinding()]
Param (
    [Parameter(Mandatory=$true)]
    [String]$SqlInstanceToBaseline,

    [Parameter(Mandatory=$false)]
    [String]$DbaDatabase = 'DBA',

    [Parameter(Mandatory=$false)]
    [String]$SqlInstanceAsDataDestination,

    [Parameter(Mandatory=$false)]
    [String]$SqlInstanceForTsqlJobs,

    [Parameter(Mandatory=$false)]
    [String]$SqlInstanceForPowershellJobs,

    [Parameter(Mandatory=$true)]
    [String]$InventoryServer,

    [Parameter(Mandatory=$false)]
    [String]$InventoryDatabase = 'DBA',

    [Parameter(Mandatory=$false)]
    [String]$HostName,

    [Parameter(Mandatory=$false)]
    [String]$SQLMonitorPath,

    [Parameter(Mandatory=$true)]
    [String]$DbaToolsFolderPath,

    [Parameter(Mandatory=$false)]
    [String]$FirstResponderKitZipFile,

    [Parameter(Mandatory=$false)]
    [String]$DarlingDataZipFile,

    [Parameter(Mandatory=$false)]
    [String]$OlaHallengrenSolutionZipFile,

    [Parameter(Mandatory=$false)]
    [String]$RemoteSQLMonitorPath = 'C:\SQLMonitor',

    [Parameter(Mandatory=$false)]
    [String[]]$DbaGroupMailId,

    [Parameter(Mandatory=$false)]
    [ValidateSet("1__sp_WhoIsActive", "2__AllDatabaseObjects", "3__XEventSession",
                "4__FirstResponderKitObjects", "5__DarlingDataObjects", "6__OlaHallengrenSolutionObjects",
                "7__sp_WhatIsRunning", "8__usp_GetAllServerInfo", "9__CopyDbaToolsModule2Host",
                "10__CopyPerfmonFolder2Host", "11__SetupPerfmonDataCollector", "12__CreateCredentialProxy",
                "13__CreateJobCollectDiskSpace", "14__CreateJobCollectPerfmonData", "15__CreateJobCollectWaitStats",
                "16__CreateJobCollectXEvents", "17__CreateJobCollectFileIOStats", "18__CreateJobPartitionsMaintenance",
                "19__CreateJobPurgeTables", "20__CreateJobRemoveXEventFiles", "21__CreateJobRunLogSaver",
                "22__CreateJobRunTempDbSaver", "23__CreateJobRunWhoIsActive", "24__CreateJobRunBlitz",
                "24__CreateJobRunBlitzIndex", "26__CreateJobRunBlitzIndexWeekly", "27__CreateJobCollectMemoryClerks",
                "28__CreateJobCollectPrivilegedInfo", "29__CreateJobCollectAgHealthState", "30__CreateJobCheckSQLAgentJobs",
                "31__CreateJobCaptureAlertMessages", "32__CreateSQLAgentAlerts", "33__CreateJobUpdateSqlServerVersions",
                "34__CreateJobCheckInstanceAvailability", "35__CreateJobGetAllServerStableInfo", "36__CreateJobGetAllServerVolatileInfo",
                "37__CreateJobGetAllServerCollectionLatencyInfo", "38__CreateJobGetAllServerSqlAgentJobs", "39__CreateJobGetAllServerDiskSpace",
                "40__CreateJobGetAllServerLogSpaceConsumers", "41__CreateJobGetAllServerTempdbSpaceUsage", "42__CreateJobGetAllServerAgHealthState",
                "43__CreateJobGetAllServerServices", "44__CreateJobGetAllServerBackups", "45__CreateJobGetAllServerAlertHistory",
                "46__CreateJobGetAllServerDashboardMail", "47__CreateJobStopStuckSQLMonitorJobs", "48__CreateJobCollectLoginExpirationInfo",
                "49__CreateJobPopulateInventoryTables", "50__CreateJobSendLoginExpiryEmails", "51__CreateJobComputeAllServerVolatileInfoHistoryHourly",
                "52__WhoIsActivePartition", "53__BlitzIndexPartition", "54__BlitzPartition",
                "55__EnablePageCompression", "56__GrafanaLogin", "57__LinkedServerOnInventory",
                "58__LinkedServerForDataDestinationInstance", "59__AlterViewsForDataDestinationInstance")]
    [String]$StartAtStep = "1__sp_WhoIsActive",

    [Parameter(Mandatory=$false)]
    [ValidateSet("1__sp_WhoIsActive", "2__AllDatabaseObjects", "3__XEventSession",
                "4__FirstResponderKitObjects", "5__DarlingDataObjects", "6__OlaHallengrenSolutionObjects",
                "7__sp_WhatIsRunning", "8__usp_GetAllServerInfo", "9__CopyDbaToolsModule2Host",
                "10__CopyPerfmonFolder2Host", "11__SetupPerfmonDataCollector", "12__CreateCredentialProxy",
                "13__CreateJobCollectDiskSpace", "14__CreateJobCollectPerfmonData", "15__CreateJobCollectWaitStats",
                "16__CreateJobCollectXEvents", "17__CreateJobCollectFileIOStats", "18__CreateJobPartitionsMaintenance",
                "19__CreateJobPurgeTables", "20__CreateJobRemoveXEventFiles", "21__CreateJobRunLogSaver",
                "22__CreateJobRunTempDbSaver", "23__CreateJobRunWhoIsActive", "24__CreateJobRunBlitz",
                "24__CreateJobRunBlitzIndex", "26__CreateJobRunBlitzIndexWeekly", "27__CreateJobCollectMemoryClerks",
                "28__CreateJobCollectPrivilegedInfo", "29__CreateJobCollectAgHealthState", "30__CreateJobCheckSQLAgentJobs",
                "31__CreateJobCaptureAlertMessages", "32__CreateSQLAgentAlerts", "33__CreateJobUpdateSqlServerVersions",
                "34__CreateJobCheckInstanceAvailability", "35__CreateJobGetAllServerStableInfo", "36__CreateJobGetAllServerVolatileInfo",
                "37__CreateJobGetAllServerCollectionLatencyInfo", "38__CreateJobGetAllServerSqlAgentJobs", "39__CreateJobGetAllServerDiskSpace",
                "40__CreateJobGetAllServerLogSpaceConsumers", "41__CreateJobGetAllServerTempdbSpaceUsage", "42__CreateJobGetAllServerAgHealthState",
                "43__CreateJobGetAllServerServices", "44__CreateJobGetAllServerBackups", "45__CreateJobGetAllServerAlertHistory",
                "46__CreateJobGetAllServerDashboardMail", "47__CreateJobStopStuckSQLMonitorJobs", "48__CreateJobCollectLoginExpirationInfo",
                "49__CreateJobPopulateInventoryTables", "50__CreateJobSendLoginExpiryEmails", "51__CreateJobComputeAllServerVolatileInfoHistoryHourly",
                "52__WhoIsActivePartition", "53__BlitzIndexPartition", "54__BlitzPartition",
                "55__EnablePageCompression", "56__GrafanaLogin", "57__LinkedServerOnInventory",
                "58__LinkedServerForDataDestinationInstance", "59__AlterViewsForDataDestinationInstance")]
    [String[]]$SkipSteps,

    [Parameter(Mandatory=$false)]
    [ValidateSet("1__sp_WhoIsActive", "2__AllDatabaseObjects", "3__XEventSession",
                "4__FirstResponderKitObjects", "5__DarlingDataObjects", "6__OlaHallengrenSolutionObjects",
                "7__sp_WhatIsRunning", "8__usp_GetAllServerInfo", "9__CopyDbaToolsModule2Host",
                "10__CopyPerfmonFolder2Host", "11__SetupPerfmonDataCollector", "12__CreateCredentialProxy",
                "13__CreateJobCollectDiskSpace", "14__CreateJobCollectPerfmonData", "15__CreateJobCollectWaitStats",
                "16__CreateJobCollectXEvents", "17__CreateJobCollectFileIOStats", "18__CreateJobPartitionsMaintenance",
                "19__CreateJobPurgeTables", "20__CreateJobRemoveXEventFiles", "21__CreateJobRunLogSaver",
                "22__CreateJobRunTempDbSaver", "23__CreateJobRunWhoIsActive", "24__CreateJobRunBlitz",
                "24__CreateJobRunBlitzIndex", "26__CreateJobRunBlitzIndexWeekly", "27__CreateJobCollectMemoryClerks",
                "28__CreateJobCollectPrivilegedInfo", "29__CreateJobCollectAgHealthState", "30__CreateJobCheckSQLAgentJobs",
                "31__CreateJobCaptureAlertMessages", "32__CreateSQLAgentAlerts", "33__CreateJobUpdateSqlServerVersions",
                "34__CreateJobCheckInstanceAvailability", "35__CreateJobGetAllServerStableInfo", "36__CreateJobGetAllServerVolatileInfo",
                "37__CreateJobGetAllServerCollectionLatencyInfo", "38__CreateJobGetAllServerSqlAgentJobs", "39__CreateJobGetAllServerDiskSpace",
                "40__CreateJobGetAllServerLogSpaceConsumers", "41__CreateJobGetAllServerTempdbSpaceUsage", "42__CreateJobGetAllServerAgHealthState",
                "43__CreateJobGetAllServerServices", "44__CreateJobGetAllServerBackups", "45__CreateJobGetAllServerAlertHistory",
                "46__CreateJobGetAllServerDashboardMail", "47__CreateJobStopStuckSQLMonitorJobs", "48__CreateJobCollectLoginExpirationInfo",
                "49__CreateJobPopulateInventoryTables", "50__CreateJobSendLoginExpiryEmails", "51__CreateJobComputeAllServerVolatileInfoHistoryHourly",
                "52__WhoIsActivePartition", "53__BlitzIndexPartition", "54__BlitzPartition",
                "55__EnablePageCompression", "56__GrafanaLogin", "57__LinkedServerOnInventory",
                "58__LinkedServerForDataDestinationInstance", "59__AlterViewsForDataDestinationInstance")]
    [String]$StopAtStep,

    [Parameter(Mandatory=$false)]
    [ValidateSet("1__sp_WhoIsActive", "2__AllDatabaseObjects", "3__XEventSession",
                "4__FirstResponderKitObjects", "5__DarlingDataObjects", "6__OlaHallengrenSolutionObjects",
                "7__sp_WhatIsRunning", "8__usp_GetAllServerInfo", "9__CopyDbaToolsModule2Host",
                "10__CopyPerfmonFolder2Host", "11__SetupPerfmonDataCollector", "12__CreateCredentialProxy",
                "13__CreateJobCollectDiskSpace", "14__CreateJobCollectPerfmonData", "15__CreateJobCollectWaitStats",
                "16__CreateJobCollectXEvents", "17__CreateJobCollectFileIOStats", "18__CreateJobPartitionsMaintenance",
                "19__CreateJobPurgeTables", "20__CreateJobRemoveXEventFiles", "21__CreateJobRunLogSaver",
                "22__CreateJobRunTempDbSaver", "23__CreateJobRunWhoIsActive", "24__CreateJobRunBlitz",
                "24__CreateJobRunBlitzIndex", "26__CreateJobRunBlitzIndexWeekly", "27__CreateJobCollectMemoryClerks",
                "28__CreateJobCollectPrivilegedInfo", "29__CreateJobCollectAgHealthState", "30__CreateJobCheckSQLAgentJobs",
                "31__CreateJobCaptureAlertMessages", "32__CreateSQLAgentAlerts", "33__CreateJobUpdateSqlServerVersions",
                "34__CreateJobCheckInstanceAvailability", "35__CreateJobGetAllServerStableInfo", "36__CreateJobGetAllServerVolatileInfo",
                "37__CreateJobGetAllServerCollectionLatencyInfo", "38__CreateJobGetAllServerSqlAgentJobs", "39__CreateJobGetAllServerDiskSpace",
                "40__CreateJobGetAllServerLogSpaceConsumers", "41__CreateJobGetAllServerTempdbSpaceUsage", "42__CreateJobGetAllServerAgHealthState",
                "43__CreateJobGetAllServerServices", "44__CreateJobGetAllServerBackups", "45__CreateJobGetAllServerAlertHistory",
                "46__CreateJobGetAllServerDashboardMail", "47__CreateJobStopStuckSQLMonitorJobs", "48__CreateJobCollectLoginExpirationInfo",
                "49__CreateJobPopulateInventoryTables", "50__CreateJobSendLoginExpiryEmails", "51__CreateJobComputeAllServerVolatileInfoHistoryHourly",
                "52__WhoIsActivePartition", "53__BlitzIndexPartition", "54__BlitzPartition",
                "55__EnablePageCompression", "56__GrafanaLogin", "57__LinkedServerOnInventory",
                "58__LinkedServerForDataDestinationInstance", "59__AlterViewsForDataDestinationInstance")]
    [String[]]$OnlySteps,

    [Parameter(Mandatory=$false)]
    [PSCredential]$SqlCredential,

	[Parameter(Mandatory=$true)]
	[securestring]$GrafanaLoginPassword,

	[Parameter(Mandatory=$false)]
	[switch]$RotateGrafanaLoginPassword,

    [Parameter(Mandatory=$false)]
    [PSCredential]$WindowsCredential,

    [Parameter(Mandatory=$false)]
    [int]$RetentionDays,

    [Parameter(Mandatory=$false)]
    [bool]$DropCreatePowerShellJobs = $false,

    [Parameter(Mandatory=$false)]
    [bool]$DropCreateWhoIsActiveTable = $false,

    [Parameter(Mandatory=$false)]
    [bool]$DropCreateLinkedServerOnInventory = $false,

    [Parameter(Mandatory=$false)]
    [bool]$DropCreateLinkedServerForDataDestinationServer = $false,

    [Parameter(Mandatory=$false)]
    [bool]$SkipPowerShellJobs = $false,

    [Parameter(Mandatory=$false)]
    [bool]$SkipMultiServerviewsUpgrade = $true,

    [Parameter(Mandatory=$false)]
    [bool]$SkipTsqlJobs = $false,

    [Parameter(Mandatory=$false)]
    [bool]$SkipRDPSessionSteps = $false,

    [Parameter(Mandatory=$false)]
    [bool]$SkipInventorySteps = $false,

    [Parameter(Mandatory=$false)]
    [bool]$SkipMultiMailJobSteps = $true,

    [Parameter(Mandatory=$false)]
    [bool]$SkipWindowsAdminAccessTest = $false,

    [Parameter(Mandatory=$false)]
    [bool]$SkipMailProfileCheck = $false,

    [Parameter(Mandatory=$false)]
    [bool]$EnableEmailAlerts = $true,

    [Parameter(Mandatory=$false)]
    [bool]$SkipCollationCheck = $false,

    [Parameter(Mandatory=$false)]
    [bool]$SkipPageCompression = $false,

    [Parameter(Mandatory=$false)]
    [bool]$SkipDriveCheck = $false,

    [Parameter(Mandatory=$false)]
    [bool]$SkipPingCheck = $false,

    [Parameter(Mandatory=$false)]
    [bool]$HasCustomizedTsqlJobs = $false,

    [Parameter(Mandatory=$false)]
    [bool]$ForceTSQLStepType4TsqlJobs = $true,

    [Parameter(Mandatory=$false)]
    [bool]$HasCustomizedPowerShellJobs = $false,

    [Parameter(Mandatory=$false)]
    [bool]$OverrideCustomizedTsqlJobs = $false,

    [Parameter(Mandatory=$false)]
    [bool]$OverrideCustomizedPowerShellJobs = $false,

    [Parameter(Mandatory=$false)]
    [bool]$ForceSetupOfTaskSchedulerJobs = $false,

    [Parameter(Mandatory=$false)]
    [bool]$ConfirmValidationOfMultiInstance = $false,

    [Parameter(Mandatory=$false)]
    [bool]$ConfirmSetupOfTaskSchedulerJobs = $false,

    [Parameter(Mandatory=$false)]
    [bool]$UpdateSQLAgentJobsThreshold = $true,

    [Parameter(Mandatory=$false)]
    [String]$XEventDirectory,

    [Parameter(Mandatory=$false)]
    [int]$JobsExecutionWaitTimeoutMinutes = 5,

    [Parameter(Mandatory=$false)]
    [bool]$MemoryOptimizedObjectsUsage = $true,

    [Parameter(Mandatory=$false)]
    [bool]$InstanceScopeFeaturesOnly = $false,

    [Parameter(Mandatory=$false)]
    [bool]$IsManagedInstance = $false,

    [Parameter(Mandatory=$false)]
    [bool]$DryRun = $false,

    [Parameter(Mandatory=$false)]
    [String]$PreQuery,

    [Parameter(Mandatory=$false)]
    [String]$PostQuery,

    [Parameter(Mandatory=$false)]
    [bool]$ReturnInlineErrorMessage = $false,

    [Parameter(Mandatory=$false)]
    [String]$GrafanaDashboardPortal = 'https://sqlmonitor.ajaydwivedi.com:3000/'
)

# Make sure to remove spaces
$SqlInstanceToBaseline = $SqlInstanceToBaseline.Trim()

$startTime = Get-Date
$ErrorActionPreference = "Stop"
$sqlmonitorVersion = '2026-02-27'
$sqlmonitorVersionDate = '2026-02-27'
$releaseDiscussionURL = "https://ajaydwivedi.com/sqlmonitor/common-errors"
$clientName = "Wrapper-InstallSQLMonitor.ps1"

<#
    v2026-Jan-31
        -> Issue#3 - Infra to Track Server and Database Configuration Changes
        -> Issue#54 - Capture Working Set & % Processor Time for all processes in Perfmon
        
    v2025-Jul-04
        -> Issue#44 - Bug fix for Foreign key issue while baselining server as the server does not exist in dbo.vw_all_server_info

    v2025-Jul-04
        -> Issue#43 - Add constraint in dbo.login_email_mapping to validate sql_instance_ip
        -> Add tsql code at end to execute on Inventory

    v2025-May-29
        -> Issue#41 - Bug/Fix DateDiff Overflow in WhoIsActive dashboard

    v2025-Feb-16
        -> Issue#33 - Setup Inventory on SQLExpress
        -> Issue#34 - Added more procs for Install-DbaDarlingData
        -> Issue#24 - Add support for Managed Instances (PAAS)

    v2024-Nov-28
        -> Issue#12 - Optimize table dbo.all_server_volatile_info_history
        -> Issue#6 - Preserve data of History Tables in Inventory Upgrade
        -> Issue#4 - Get Max Server Memory in dbo.all_server_stable_info
        -> Issue#57 - Put steps of Get-AllServerInfo & Get-AllServerCollectedData into separate jobs
        -> Issue#10 - Alert Engine using Python + SQLServer

    v2024-Sep-30
        -> Issue#51 - Integrate inventory objects
        -> Issue#10 - Work on SQLMonitor Alerting
        -> Issue#50 - Add switch to create tsql script jobs as TSQL Step type
        -> Issue#44 - Copy grafana login from Inventory to keep same SID
        -> Issue$43 - Add avg_disk_wait_ms
        -> Issue#38 - Add Infra to Track AG State Change

    v2024-Mar-31
        -> Issue#30 - Add flag for choice of MemoryOptimized Objects
        -> Issue#29 - Add additional verification step for Instance-Availability apart from job [(dba) Check-InstanceAvailability]
        -> Issue#21 - Add Parameters to Skip Particular Wait Type in usp_waits_per_core_per_minute
        -> Issue#19 - Control Immediate Removal of Perfmon File in Job [(dba) Collect-PerfmonData]
        -> Issue#22 - Added Trust for Certificate & Encryption for SQL Connections
        -> Issue#24 - Daily Mailer for All Server Health Dashboard
        -> Updated few debugging queries
        -> Updated dbo.usp_get_credential as part of bug fix
        -> Updated BlitzIndex Analysis Dashboard variables and ID

    v1.6.0 - 2023-Dec-30
        -> Issue#13 - Capture sp_Blitz & Create Dashboard
        -> Issue#12 - Compatibility with new dbatools version

    v1.5.0.6 - Intermediate Release - 2023-Aug-30
        -> Issue#7 - Dashboard for SQL Agent Jobs
        -> Issue#272 - Implement TempDbSaver
        -> Issue#273 - Implement LogSaver
        -> Central Dashboard for Disk Space
        -> Issue#280 - Add host_distribution & processor_name in dbo.all_server_table_info
        -> Issue#278 - Capture AG Latency Over Time

    v1.5.0.4 - Intermediate Release - 2023-Jul-19
        -> Issue#270 - Add feature to Ignore Existing Jobs if Customized
        -> Issue#269 - Add last execution duration in table dbo.sql_agent_job_stats
        -> Issue#268 - Add tables sql_agent_job_stats & memory_clerks in Collection Latency Dashboard
        -> Issue#267 - Error in M/r Grants Pending panel on Monitoring - Live - Distributed dashboard
        -> Issue#265 - Disabled Jobs Not Appearing in Dashboard Panel SQLAgent Job Activity Monitor
    v1.5.0 - 2023-June-30
        -> Issue#255 - Support of SQLExpress for Inventory Setup
    v1.4.0 - 2023-Mar-31
        -> https://github.com/imajaydwivedi/SQLMonitor/releases/tag/v1.4.0
#>

$verbose = $false;
if ($PSBoundParameters.ContainsKey('Verbose')) { # Command line specifies -Verbose[:$false]
    $verbose = $PSBoundParameters.Get_Item('Verbose')
}

$debug = $false;
if ($PSBoundParameters.ContainsKey('Debug')) { # Command line specifies -Debug[:$false]
    $debug = $PSBoundParameters.Get_Item('Debug')
}


# Declare other important variables/Parameters
[String]$MailProfileFileName = "DDL-[DatabaseMail_Using_GMail].sql"
[String]$WhoIsActiveFileName = "DDL-[sp_WhoIsActive_v12_00(Modified)].sql"
[String]$GrafanaLoginFileName = "DCL-[grafana-login].sql"
[String]$UpdateSQLAgentJobsThresholdFileName = "QRY-UPDATE-[SQLAgentJobsThreshold].sql"
[String]$AllDatabaseObjectsFileName = "SCH-Create-All-Objects.sql"
[String]$XEventSessionFileName = "SCH-Create-XEvents.sql"
[String]$XEventSessionRingBufferFileName = "SCH-Create-XEvents-RingBuffer.sql"
[String]$WhatIsRunningFileName = "SCH-sp_WhatIsRunning.sql"
[String]$UspGetAllServerInfoFileName = "SCH-usp_GetAllServerInfo.sql"
[String]$UspCheckInstanceAvailabilityFileName = "SCH-usp_check_instance_availability.sql"
[String]$UspGetAllServerDashboardMailFileName = "SCH-usp_GetAllServerDashboardMail.sql"
[String]$UspGetAllServerCollectedDataFileName = "SCH-usp_GetAllServerCollectedData.sql"
[String]$UspWrapperGetAllServerInfoFileName = "SCH-usp_wrapper_GetAllServerInfo.sql"
[String]$UspWrapperGetAllServerCollectedDataFileName = "SCH-usp_wrapper_GetAllServerCollectedData.sql"
[String]$UspWrapperCollectPrivilegedInfoFileName = "SCH-usp_wrapper_CollectPrivilegedInfo.sql"
[String]$UspCaptureAlertMessagesFileName = "SCH-usp_capture_alert_messages.sql"
[String]$UspCreateAgentAlertsFileName = "SCH-usp_create_agent_alerts.sql"
[String]$UspCollectWaitStatsFileName = "SCH-usp_collect_wait_stats.sql"
[String]$UspCollectFileIOStatsFileName = "SCH-usp_collect_file_io_stats.sql"
[String]$UspCollectXeventMetricsFileName = "SCH-usp_collect_xevent_metrics.sql"
[String]$UspCollectXeventMetricsRingBufferFileName = "SCH-usp_collect_xevent_metrics_ringbuffer.sql"
[String]$UspCollectPrivilegedInfoFileName = "SCH-usp_collect_privileged_info.sql"
[String]$UspComputeAllServerVolatileInfoHistoryHourlyFileName = "SCH-usp_compute_all_server_volatile_info_history_hourly.sql"
[String]$UspPartitionMaintenanceFileName = "SCH-usp_partition_maintenance.sql"
[String]$UspPurgeTablesFileName = "SCH-usp_purge_tables.sql"
[String]$UspRunWhoIsActiveFileName = "SCH-usp_run_WhoIsActive.sql"
[String]$UspActiveRequestsCountFileName = "SCH-usp_active_requests_count.sql"
[String]$UspAvgDiskWaitMsFileName = "SCH-usp_avg_disk_wait_ms.sql"
[String]$UspAvgDiskLatencyMsFileName = "SCH-usp_avg_disk_latency_ms.sql"
[String]$UspEnablePageCompressionFileName = "SCH-usp_enable_page_compression.sql"
[String]$UspCollectAllServerLoginExpirationInfoFileName = "SCH-usp_collect_all_server_login_expiration_info.sql"
[String]$UspCollectMemoryClerksFileName = "SCH-usp_collect_memory_clerks.sql"
[String]$UspCollectAgHealthStateFileName = "SCH-usp_collect_ag_health_state.sql"
[String]$UspCollectDiskSpaceFileName = "SCH-usp_collect_disk_space.sql"
[String]$UspCollectPerformanceMetricsFileName = "SCH-usp_collect_performance_metrics.sql"
[String]$UspCheckSQLAgentJobsFileName = "SCH-usp_check_sql_agent_jobs.sql"
[String]$UspLogSaverFileName = "SCH-usp_LogSaver.sql"
[String]$UspPopulateSmaSqlInstanceFileName = "SCH-usp_populate_sma_sql_instance.sql"
[String]$UspSendLoginExpiryEmailsFileName = "SCH-usp_send_login_expiry_emails.sql"
[String]$UspTempDbSaverFileName = "SCH-usp_TempDbSaver.sql"
[String]$UspWaitsPerCorePerMinuteFileName = "SCH-usp_waits_per_core_per_minute.sql"
[String]$UspWrapperPopulateSmaSqlInstanceFileName = "SCH-usp_wrapper_populate_sma_sql_instance.sql"
[String]$WhoIsActivePartitionFileName = "SCH-WhoIsActive-Partitioning.sql"
[String]$BlitzIndexPartitionFileName = "SCH-BlitzIndex-Partitioning.sql"
[String]$BlitzIndexMode0PartitionFileName = "SCH-BlitzIndex_Mode0-Partitioning.sql"
[String]$BlitzIndexMode1PartitionFileName = "SCH-BlitzIndex_Mode1-Partitioning.sql"
[String]$BlitzIndexMode4PartitionFileName = "SCH-BlitzIndex_Mode4-Partitioning.sql"
[String]$BlitzPartitionFileName = "SCH-Blitz-Partitioning.sql"

# Jobs present on All Servers
[String]$CaptureAlertMessagesJobFileName = "SCH-Job-[(dba) Capture-AlertMessages].sql"
[String]$CheckSQLAgentJobsJobFileName = "SCH-Job-[(dba) Check-SQLAgentJobs].sql"
[String]$CollectMemoryClerksJobFileName = "SCH-Job-[(dba) Collect-MemoryClerks].sql"
[String]$CollectAgHealthStateJobFileName = "SCH-Job-[(dba) Collect-AgHealthState].sql"
[String]$CollectDiskSpaceJobFileName = "SCH-Job-[(dba) Collect-DiskSpace].sql"
[String]$CollectDiskSpaceManagedInstanceJobFileName = "SCH-Job-[(dba) Collect-DiskSpace-ManagedInstance].sql"
[String]$CollectOSProcessesJobFileName = "SCH-Job-[(dba) Collect-OSProcesses].sql"
[String]$CollectPerfmonDataJobFileName = "SCH-Job-[(dba) Collect-PerfmonData].sql"
[String]$CollectPerfmonDataManagedInstanceJobFileName = "SCH-Job-[(dba) Collect-PerfmonData-ManagedInstance].sql"
[String]$CollectWaitStatsJobFileName = "SCH-Job-[(dba) Collect-WaitStats].sql"
[String]$CollectFileIOStatsJobFileName = "SCH-Job-[(dba) Collect-FileIOStats].sql"
[String]$CollectXEventsJobFileName = "SCH-Job-[(dba) Collect-XEvents].sql"
[String]$CollectPrivilegedInfoJobFileName = "SCH-Job-[(dba) Collect-PrivilegedInfo].sql"
[String]$PartitionsMaintenanceJobFileName = "SCH-Job-[(dba) Partitions-Maintenance].sql"
[String]$PurgeTablesJobFileName = "SCH-Job-[(dba) Purge-Tables].sql"
[String]$RemoveXEventFilesJobFileName = "SCH-Job-[(dba) Remove-XEventFiles].sql"
[String]$RunBlitzIndexJobFileName = "SCH-Job-[(dba) Run-BlitzIndex].sql"
[String]$RunBlitzIndexWeeklyJobFileName = "SCH-Job-[(dba) Run-BlitzIndex - Weekly].sql"
[String]$RunBlitzJobFileName = "SCH-Job-[(dba) Run-Blitz].sql"
[String]$RunLogSaverJobFileName = "SCH-Job-[(dba) Run-LogSaver].sql"
[String]$RunTempDbSaverJobFileName = "SCH-Job-[(dba) Run-TempDbSaver].sql"
[String]$RunWhoIsActiveJobFileName = "SCH-Job-[(dba) Run-WhoIsActive].sql"

# Jobs present on Inventory Only
[String]$CheckInstanceAvailabilityJobFileName = "SCH-Job-[(dba) Check-InstanceAvailability].sql"
[String]$CollectLoginExpirationInfoJobFileName = "SCH-Job-[(dba) Collect Login Expiration Info].sql"
[String]$GetAllServerStableInfoJobFileName = "SCH-Job-[(dba) Get-AllServerStableInfo].sql"
[String]$GetAllServerVolatileInfoJobFileName = "SCH-Job-[(dba) Get-AllServerVolatileInfo].sql"
[String]$GetAllServerCollectionLatencyInfoJobFileName = "SCH-Job-[(dba) Get-AllServerCollectionLatencyInfo].sql"
[String]$GetAllServerSqlAgentJobsJobFileName = "SCH-Job-[(dba) Get-AllServerSqlAgentJobs].sql"
[String]$GetAllServerDiskSpaceJobFileName = "SCH-Job-[(dba) Get-AllServerDiskSpace].sql"
[String]$GetAllServerLogSpaceConsumersJobFileName = "SCH-Job-[(dba) Get-AllServerLogSpaceConsumers].sql"
[String]$GetAllServerTempdbSpaceUsageJobFileName = "SCH-Job-[(dba) Get-AllServerTempdbSpaceUsage].sql"
[String]$GetAllServerAgHealthStateJobFileName = "SCH-Job-[(dba) Get-AllServerAgHealthState].sql"
[String]$GetAllServerServicesJobFileName = "SCH-Job-[(dba) Get-AllServerServices].sql"
[String]$GetAllServerBackupsJobFileName = "SCH-Job-[(dba) Get-AllServerBackups].sql"
[String]$GetAllServerAlertHistoryJobFileName = "SCH-Job-[(dba) Get-AllServerAlertHistory].sql"
[String]$GetAllServerDashboardMailJobFileName = "SCH-Job-[(dba) Get-AllServerDashboardMail].sql"
[String]$InventorySpecificObjectsFileName = "SCH-Create-Inventory-Specific-Objects.sql"
[String]$LinkedServerOnInventoryFileName = "SCH-Linked-Servers-Sample.sql"
[String]$PopulateInventoryTablesJobFileName = "SCH-Job-[(dba) Populate Inventory Tables].sql"
[String]$SendLoginExpiryEmailsJobFileName = "SCH-Job-[(dba) Send Login Expiry EMails].sql"
[String]$ComputeAllServerVolatileInfoHistoryHourlyJobFileName = "SCH-Job-[(dba) Compute-AllServerVolatileInfoHistoryHourly].sql"
[String]$StopStuckSQLMonitorJobsJobFileName = "SCH-Job-[(dba) Stop-StuckSQLMonitorJobs].sql"
[String]$TestWindowsAdminAccessJobFileName = "SCH-Job-[(dba) Test-WindowsAdminAccess].sql"
[String]$UpdateSqlServerVersionsJobFileName = "SCH-Job-[(dba) Update-SqlServerVersions].sql"

# Reset default if Task Scheduler jobs to be created.
if($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) {
    $ForceTSQLStepType4TsqlJobs = $false
    $ConfirmSetupOfTaskSchedulerJobs = $true
    $ForceSetupOfTaskSchedulerJobs = $true
}

# Check if PortNo is specified for $SqlInstanceToBaseline
#"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract port no from `$SqlInstanceToBaseline ($SqlInstanceToBaseline) if provided."
$Port4SqlInstanceToBaseline = $null
$SqlInstanceToBaselineWithOutPort = $SqlInstanceToBaseline
if($SqlInstanceToBaseline -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
    $Port4SqlInstanceToBaseline = $Matches['PortNo']
    $SqlInstanceToBaselineWithOutPort = $Matches['SqlInstance']
}

# Check if PortNo is specified for $InventoryServer
#"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract port no from `$InventoryServer ($InventoryServer) if provided."
$Port4InventoryServer = $null
$InventoryServerWithOutPort = $InventoryServer
if($InventoryServer -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
    $Port4InventoryServer = $Matches['PortNo']
    $InventoryServerWithOutPort = $Matches['SqlInstance']
}

# Port for $InventoryServer will always be present is it connects on non-default port
if($SqlInstanceToBaselineWithOutPort -eq $InventoryServerWithOutPort) {
    if($SqlInstanceToBaseline -ne $InventoryServer) {
        if($verbose) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Logically InventoryServer & SqlInstanceToBaseline seems same, but still differ in port."
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Correcting this different port situation."
        }
        $SqlInstanceToBaseline = $InventoryServer
        $Port4SqlInstanceToBaseline = $Port4InventoryServer
    }
}


"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'START:', "Working on server [$SqlInstanceToBaseline]." | Write-Host -ForegroundColor Yellow
if(-not [String]::IsNullOrEmpty($Port4SqlInstanceToBaseline)) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "SQL Port for [$SqlInstanceToBaseline] => $Port4SqlInstanceToBaseline." | Write-Host -ForegroundColor Yellow
}
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'START:', "Deploying SQLMonitor v$sqlmonitorVersion released on $sqlmonitorVersionDate.." | Write-Host -ForegroundColor Yellow
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'START:', "For issues with this version, kindly visit $releaseDiscussionURL" | Write-Host -ForegroundColor Yellow
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'START:', "For general support, join #sqlmonitor channel on 'sqlcommunity.slack.com <https://ajaydwivedi.com/go/slack>' workspace.`n" | Write-Host -ForegroundColor Yellow
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'START:', "For paid support, kindly reach out to 'Ajay Dwivedi <ajay.dwivedi2007@gmail.com>'`n" | Write-Host -ForegroundColor Yellow


# All Steps
$AllSteps = @(  "1__sp_WhoIsActive", "2__AllDatabaseObjects", "3__XEventSession",
                "4__FirstResponderKitObjects", "5__DarlingDataObjects", "6__OlaHallengrenSolutionObjects",
                "7__sp_WhatIsRunning", "8__usp_GetAllServerInfo", "9__CopyDbaToolsModule2Host",
                "10__CopyPerfmonFolder2Host", "11__SetupPerfmonDataCollector", "12__CreateCredentialProxy",
                "13__CreateJobCollectDiskSpace", "14__CreateJobCollectPerfmonData", "15__CreateJobCollectWaitStats",
                "16__CreateJobCollectXEvents", "17__CreateJobCollectFileIOStats", "18__CreateJobPartitionsMaintenance",
                "19__CreateJobPurgeTables", "20__CreateJobRemoveXEventFiles", "21__CreateJobRunLogSaver",
                "22__CreateJobRunTempDbSaver", "23__CreateJobRunWhoIsActive", "24__CreateJobRunBlitz",
                "24__CreateJobRunBlitzIndex", "26__CreateJobRunBlitzIndexWeekly", "27__CreateJobCollectMemoryClerks",
                "28__CreateJobCollectPrivilegedInfo", "29__CreateJobCollectAgHealthState", "30__CreateJobCheckSQLAgentJobs",
                "31__CreateJobCaptureAlertMessages", "32__CreateSQLAgentAlerts", "33__CreateJobUpdateSqlServerVersions",
                "34__CreateJobCheckInstanceAvailability", "35__CreateJobGetAllServerStableInfo", "36__CreateJobGetAllServerVolatileInfo",
                "37__CreateJobGetAllServerCollectionLatencyInfo", "38__CreateJobGetAllServerSqlAgentJobs", "39__CreateJobGetAllServerDiskSpace",
                "40__CreateJobGetAllServerLogSpaceConsumers", "41__CreateJobGetAllServerTempdbSpaceUsage", "42__CreateJobGetAllServerAgHealthState",
                "43__CreateJobGetAllServerServices", "44__CreateJobGetAllServerBackups", "45__CreateJobGetAllServerAlertHistory",
                "46__CreateJobGetAllServerDashboardMail", "47__CreateJobStopStuckSQLMonitorJobs", "48__CreateJobCollectLoginExpirationInfo",
                "49__CreateJobPopulateInventoryTables", "50__CreateJobSendLoginExpiryEmails", "51__CreateJobComputeAllServerVolatileInfoHistoryHourly",
                "52__WhoIsActivePartition", "53__BlitzIndexPartition", "54__BlitzPartition",
                "55__EnablePageCompression", "56__GrafanaLogin", "57__LinkedServerOnInventory",
                "58__LinkedServerForDataDestinationInstance", "59__AlterViewsForDataDestinationInstance")

# TSQL Jobs
$TsqlJobSteps = @(
                "15__CreateJobCollectWaitStats", "16__CreateJobCollectXEvents", "17__CreateJobCollectFileIOStats",
                "18__CreateJobPartitionsMaintenance", "19__CreateJobPurgeTables", "21__CreateJobRunLogSaver",
                "22__CreateJobRunTempDbSaver", "23__CreateJobRunWhoIsActive", "20__CreateJobRemoveXEventFiles",
                "24__CreateJobRunBlitz", "24__CreateJobRunBlitzIndex", "26__CreateJobRunBlitzIndexWeekly", 
                "27__CreateJobCollectMemoryClerks", "28__CreateJobCollectPrivilegedInfo", "29__CreateJobCollectAgHealthState", 
                "30__CreateJobCheckSQLAgentJobs", "31__CreateJobCaptureAlertMessages", "35__CreateJobGetAllServerStableInfo",
                "36__CreateJobGetAllServerVolatileInfo", "37__CreateJobGetAllServerCollectionLatencyInfo", "38__CreateJobGetAllServerSqlAgentJobs",
                "39__CreateJobGetAllServerDiskSpace", "40__CreateJobGetAllServerLogSpaceConsumers", "41__CreateJobGetAllServerTempdbSpaceUsage",
                "42__CreateJobGetAllServerAgHealthState", "43__CreateJobGetAllServerServices", "44__CreateJobGetAllServerBackups",
                "46__CreateJobGetAllServerDashboardMail", "48__CreateJobCollectLoginExpirationInfo", "49__CreateJobPopulateInventoryTables",
                "50__CreateJobSendLoginExpiryEmails", "51__CreateJobComputeAllServerVolatileInfoHistoryHourly")

# PowerShell Jobs
$PowerShellJobSteps = @(
                "13__CreateJobCollectDiskSpace", "14__CreateJobCollectPerfmonData",
                "33__CreateJobUpdateSqlServerVersions", "34__CreateJobCheckInstanceAvailability", "47__CreateJobStopStuckSQLMonitorJobs",
                "49__CreateJobPopulateInventoryTables")

# RDPSessionSteps
$RDPSessionSteps = @("9__CopyDbaToolsModule2Host", "10__CopyPerfmonFolder2Host", "11__SetupPerfmonDataCollector")

# Jobs Only created on Inventory Server
$InventorySteps = @("8__usp_GetAllServerInfo", "33__CreateJobUpdateSqlServerVersions", "34__CreateJobCheckInstanceAvailability", 
                "35__CreateJobGetAllServerStableInfo", "36__CreateJobGetAllServerVolatileInfo", "37__CreateJobGetAllServerCollectionLatencyInfo",
                "38__CreateJobGetAllServerSqlAgentJobs", "39__CreateJobGetAllServerDiskSpace", "40__CreateJobGetAllServerLogSpaceConsumers",
                "41__CreateJobGetAllServerTempdbSpaceUsage", "42__CreateJobGetAllServerAgHealthState", "43__CreateJobGetAllServerServices",
                "44__CreateJobGetAllServerBackups", "45__CreateJobGetAllServerAlertHistory", "46__CreateJobGetAllServerDashboardMail", "47__CreateJobStopStuckSQLMonitorJobs",
                "48__CreateJobCollectLoginExpirationInfo", "49__CreateJobPopulateInventoryTables", "50__CreateJobSendLoginExpiryEmails",
                "51__CreateJobComputeAllServerVolatileInfoHistoryHourly")

# Jobs that send mail notification to others apart from DBA team
$MultiMailJobSteps = @("46__CreateJobGetAllServerDashboardMail", "50__CreateJobSendLoginExpiryEmails")

# If Managed Instance
if($IsManagedInstance) {
    $InstanceScopeFeaturesOnly = $true
}

if($InstanceScopeFeaturesOnly) {
    $SkipPowerShellJobs = $true
    $SkipRDPSessionSteps = $true
    $ForceTSQLStepType4TsqlJobs = $true
    $SkipWindowsAdminAccessTest = $true

    $PowerShellJobSteps = $($PowerShellJobSteps | % {if($_ -notin @('13__CreateJobCollectDiskSpace','14__CreateJobCollectPerfmonData')){$_}});
    $TsqlJobSteps = $TsqlJobSteps + $( @('13__CreateJobCollectDiskSpace','14__CreateJobCollectPerfmonData') | % {if($_ -notin $TsqlJobSteps){$_}} );
}

# Validate to ensure either of Skip Or Only Steps are provided
if($OnlySteps.Count -gt 0 -and $SkipSteps.Count -gt 0) {
    if ($ReturnInlineErrorMessage) {
        "Parameters {OnlySteps} & {SkipSteps} are mutually exclusive.`n`tOnly one of these should be provided." | Write-Error
    }
    else {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Parameters {OnlySteps} & {SkipSteps} are mutually exclusive.`n`tOnly one of these should be provided." | Write-Host -ForegroundColor Red
        Write-Error "Stop here. Fix above issue."
    }
}

# Print warning if OnlySteps are provided
if($OnlySteps.Count -gt 0) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter {OnlySteps} has been provided.`n`tThis parameter is mutually exclusive with other parameters.`n`tSo overrides all other parameters." | Write-Host -ForegroundColor Yellow
    Write-Warning "ATTENTION Required on above message."
}


# Add $PowerShellJobSteps to Skip Jobs
if($SkipPowerShellJobs) {
    $SkipSteps = $SkipSteps + $($PowerShellJobSteps | % {if($_ -notin $SkipSteps){$_}});
}

# Add $RDPSessionSteps to Skip Jobs
if($SkipRDPSessionSteps) {
    $SkipSteps = $SkipSteps + $($RDPSessionSteps | % {if($_ -notin $SkipSteps){$_}});
}

# Add $InventorySteps to Skip Jobs
if($SkipInventorySteps) {
    $SkipSteps = $SkipSteps + $($InventorySteps | % {if($_ -notin $SkipSteps){$_}});
}

# Print Job Step names
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$PowerShellJobSteps => `n`n`t`t`t`t$($PowerShellJobSteps -join "`n`t`t`t`t")`n"
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$RDPSessionSteps => `n`n`t`t`t`t$($RDPSessionSteps -join "`n`t`t`t`t")`n"
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$InventorySteps => `n`n`t`t`t`t$($InventorySteps -join "`n`t`t`t`t")`n"
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$MultiMailJobSteps => `n`n`t`t`t`t$($MultiMailJobSteps -join "`n`t`t`t`t")`n"

# Add $TsqlJobSteps to Skip Jobs
if($SkipTsqlJobs) {
    $SkipSteps = $SkipSteps + $($TsqlJobSteps | % {if($_ -notin $SkipSteps){$_}});
}

# Skip Compression
if($SkipPageCompression -and ('55__EnablePageCompression' -notin $SkipSteps)) {
    $SkipSteps += @('55__EnablePageCompression')
}

# For backward compatability
$SkipAllJobs = $false
if($SkipTsqlJobs -and $SkipPowerShellJobs) {
    $SkipAllJobs = $true
}

"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Clearing old PSSessions.."
Get-PSSession | Remove-PSSession

if($SqlInstanceToBaseline -eq '.' -or $SqlInstanceToBaseline -eq 'localhost') {
    if ($ReturnInlineErrorMessage) {
        "'localhost' or '.' are not validate SQLInstance names." | Write-Error
    }
    else {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "'localhost' or '.' are not validate SQLInstance names." | Write-Host -ForegroundColor Red
        Write-Error "Stop here. Fix above issue."
    }
}

# Evaluate path of SQLMonitor folder
if( (-not [String]::IsNullOrEmpty($PSScriptRoot)) -or ((-not [String]::IsNullOrEmpty($SQLMonitorPath)) -and $(Test-Path $SQLMonitorPath)) ) {
    if([String]::IsNullOrEmpty($SQLMonitorPath)) {
        $SQLMonitorPath = $(Split-Path $PSScriptRoot -Parent)
    }
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$SQLMonitorPath = '$SQLMonitorPath'"
}
else {
    if ($ReturnInlineErrorMessage) {
		"Kindly provide 'SQLMonitorPath' parameter value" | Write-Error
	}
	else {
		"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly provide 'SQLMonitorPath' parameter value" | Write-Host -ForegroundColor Red
        Write-Error "Stop here. Fix above issue."
    }
}

# Logs folder
if($RemoteSQLMonitorPath.EndsWith('\')) {
    $logsPath = $RemoteSQLMonitorPath+'Logs';
} else {
    $logsPath = $RemoteSQLMonitorPath+'\'+'Logs';
}

# Set windows credential if valid AD credential is provided as SqlCredential
if( [String]::IsNullOrEmpty($WindowsCredential) -and (-not [String]::IsNullOrEmpty($SqlCredential)) -and $SqlCredential.UserName -like "*\*" ) {
    $WindowsCredential = $SqlCredential
}

# Remove end trailer of '\'
if($RemoteSQLMonitorPath.EndsWith('\')) {
    $RemoteSQLMonitorPath = $RemoteSQLMonitorPath.TrimEnd('\')
}

"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$SqlInstanceToBaseline = [$SqlInstanceToBaseline]"
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$SqlCredential => "
$SqlCredential | ft -AutoSize
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$WindowsCredential => "
$WindowsCredential | ft -AutoSize

# Construct File Path Variables
$ddlPath = Join-Path $SQLMonitorPath "DDLs"
$psScriptPath = Join-Path $SQLMonitorPath "SQLMonitor"
$isUpgradeScenario = $false
$SkipPowerShellJobs4SQLCluster = $false
$SkipPowerShellJobs4Host = $false

$mailProfileFilePath = "$ddlPath\$MailProfileFileName"
$WhoIsActiveFilePath = "$ddlPath\$WhoIsActiveFileName"
$AllDatabaseObjectsFilePath = "$ddlPath\$AllDatabaseObjectsFileName"
$InventorySpecificObjectsFilePath = "$ddlPath\$InventorySpecificObjectsFileName"
$XEventSessionFilePath = "$ddlPath\$XEventSessionFileName"
$XEventSessionRingBufferFilePath = "$ddlPath\$XEventSessionRingBufferFileName"
$WhatIsRunningFilePath = "$ddlPath\$WhatIsRunningFileName"
$GetAllServerInfoFilePath = "$ddlPath\$UspGetAllServerInfoFileName"
$UspCheckInstanceAvailabilityFilePath = "$ddlPath\$UspCheckInstanceAvailabilityFileName"
$GetAllServerCollectedDataFilePath = "$ddlPath\$UspGetAllServerCollectedDataFileName"
$UspWrapperGetAllServerInfoFilePath = "$ddlPath\$UspWrapperGetAllServerInfoFileName"
$UspWrapperGetAllServerCollectedDataFilePath = "$ddlPath\$UspWrapperGetAllServerCollectedDataFileName"
$UspGetAllServerDashboardMailFilePath = "$ddlPath\$UspGetAllServerDashboardMailFileName"
$UspWrapperCollectPrivilegedInfoFilePath = "$ddlPath\$UspWrapperCollectPrivilegedInfoFileName"
$UspCaptureAlertMessagesFilePath = "$ddlPath\$UspCaptureAlertMessagesFileName"
$UspCreateAgentAlertsFilePath = "$ddlPath\$UspCreateAgentAlertsFileName"
$UspCollectWaitStatsFilePath = "$ddlPath\$UspCollectWaitStatsFileName"
$UspCollectFileIOStatsFilePath = "$ddlPath\$UspCollectFileIOStatsFileName"
$UspCollectXeventMetricsFilePath = "$ddlPath\$UspCollectXeventMetricsFileName"
$UspCollectXeventMetricsRingBufferFilePath = "$ddlPath\$UspCollectXeventMetricsRingBufferFileName"
$UspPartitionMaintenanceFilePath = "$ddlPath\$UspPartitionMaintenanceFileName"
$UspPurgeTablesFilePath = "$ddlPath\$UspPurgeTablesFileName"
$UspRunWhoIsActiveFilePath = "$ddlPath\$UspRunWhoIsActiveFileName"
$UspActiveRequestsCountFilePath = "$ddlPath\$UspActiveRequestsCountFileName"
$UspWaitsPerCorePerMinuteFilePath = "$ddlPath\$UspWaitsPerCorePerMinuteFileName"
$UspAvgDiskWaitMsFilePath = "$ddlPath\$UspAvgDiskWaitMsFileName"
$UspAvgDiskLatencyMsFilePath = "$ddlPath\$UspAvgDiskLatencyMsFileName"
$UspEnablePageCompressionFilePath = "$ddlPath\$UspEnablePageCompressionFileName"
$UspCollectAllServerLoginExpirationInfoFilePath = "$ddlPath\$UspCollectAllServerLoginExpirationInfoFileName"
$UspCollectMemoryClerksFilePath = "$ddlPath\$UspCollectMemoryClerksFileName"
$UspCollectAgHealthStateFilePath = "$ddlPath\$UspCollectAgHealthStateFileName"
$UspCollectDiskSpaceFilePath = "$ddlPath\$UspCollectDiskSpaceFileName"
$UspCollectPerformanceMetricsFilePath = "$ddlPath\$UspCollectPerformanceMetricsFileName"
$UspCheckSQLAgentJobsFilePath = "$ddlPath\$UspCheckSQLAgentJobsFileName"
$UspLogSaverFilePath = "$ddlPath\$UspLogSaverFileName"
$UspPopulateSmaSqlInstanceFilePath = "$ddlPath\$UspPopulateSmaSqlInstanceFileName"
$UspSendLoginExpiryEmailsFilePath = "$ddlPath\$UspSendLoginExpiryEmailsFileName"
$UspComputeAllServerVolatileInfoHistoryHourlyFilePath = "$ddlPath\$UspComputeAllServerVolatileInfoHistoryHourlyFileName"
$UspTempDbSaverFilePath = "$ddlPath\$UspTempDbSaverFileName"
$UspWrapperPopulateSmaSqlInstanceFilePath = "$ddlPath\$UspWrapperPopulateSmaSqlInstanceFileName"
$WhoIsActivePartitionFilePath = "$ddlPath\$WhoIsActivePartitionFileName"
$BlitzIndexPartitionFilePath = "$ddlPath\$BlitzIndexPartitionFileName"
$BlitzPartitionFilePath = "$ddlPath\$BlitzPartitionFileName"
$BlitzIndexMode0PartitionFilePath = "$ddlPath\$BlitzIndexMode0PartitionFileName"
$BlitzIndexMode1PartitionFilePath = "$ddlPath\$BlitzIndexMode1PartitionFileName"
$BlitzIndexMode4PartitionFilePath = "$ddlPath\$BlitzIndexMode4PartitionFileName"
$GrafanaLoginFilePath = "$ddlPath\$GrafanaLoginFileName"
$UspCollectPrivilegedInfoFilePath = "$ddlPath\$UspCollectPrivilegedInfoFileName"

# Jobs for All Servers
$CaptureAlertMessagesJobFilePath = "$ddlPath\$CaptureAlertMessagesJobFileName"
$CollectDiskSpaceJobFilePath = "$ddlPath\$CollectDiskSpaceJobFileName"
$CollectDiskSpaceManagedInstanceJobFilePath = "$ddlPath\$CollectDiskSpaceManagedInstanceJobFileName"
$CollectOSProcessesJobFilePath = "$ddlPath\$CollectOSProcessesJobFileName"
$CollectPerfmonDataJobFilePath = "$ddlPath\$CollectPerfmonDataJobFileName"
$CollectPerfmonDataManagedInstanceJobFilePath = "$ddlPath\$CollectPerfmonDataManagedInstanceJobFileName"
$CollectWaitStatsJobFilePath = "$ddlPath\$CollectWaitStatsJobFileName"
$CollectFileIOStatsJobFilePath = "$ddlPath\$CollectFileIOStatsJobFileName"
$CollectXEventsJobFilePath = "$ddlPath\$CollectXEventsJobFileName"
$CollectPrivilegedInfoJobFilePath = "$ddlPath\$CollectPrivilegedInfoJobFileName"

$CollectPrivilegedInfoJobFilePath = "$ddlPath\$CollectPrivilegedInfoJobFileName"
$CollectPrivilegedInfoJobFilePath = "$ddlPath\$CollectPrivilegedInfoJobFileName"

$PartitionsMaintenanceJobFilePath = "$ddlPath\$PartitionsMaintenanceJobFileName"
$PurgeTablesJobFilePath = "$ddlPath\$PurgeTablesJobFileName"
$RemoveXEventFilesJobFilePath = "$ddlPath\$RemoveXEventFilesJobFileName"
$RunLogSaverJobFilePath = "$ddlPath\$RunLogSaverJobFileName"
$RunTempDbSaverJobFilePath = "$ddlPath\$RunTempDbSaverJobFileName"
$RunWhoIsActiveJobFilePath = "$ddlPath\$RunWhoIsActiveJobFileName"
$RunBlitzIndexJobFilePath = "$ddlPath\$RunBlitzIndexJobFileName"
$RunBlitzIndexWeeklyJobFilePath = "$ddlPath\$RunBlitzIndexWeeklyJobFileName"
$RunBlitzJobFilePath = "$ddlPath\$RunBlitzJobFileName"
$CollectMemoryClerksJobFilePath = "$ddlPath\$CollectMemoryClerksJobFileName"
$CollectAgHealthStateJobFilePath = "$ddlPath\$CollectAgHealthStateJobFileName"
$CheckSQLAgentJobsJobFilePath = "$ddlPath\$CheckSQLAgentJobsJobFileName"
$LinkedServerOnInventoryFilePath = "$ddlPath\$LinkedServerOnInventoryFileName"
$TestWindowsAdminAccessJobFilePath = "$ddlPath\$TestWindowsAdminAccessJobFileName"

# Jobs for Inventory Server
$GetAllServerDashboardMailJobFilePath = "$ddlPath\$GetAllServerDashboardMailJobFileName"
#$GetAllServerInfoJobFilePath = "$ddlPath\$GetAllServerInfoJobFileName"
#$GetAllServerCollectedDataJobFilePath = "$ddlPath\$GetAllServerCollectedDataJobFileName"

$GetAllServerStableInfoJobFilePath = "$ddlPath\$GetAllServerStableInfoJobFileName"
$GetAllServerVolatileInfoJobFilePath = "$ddlPath\$GetAllServerVolatileInfoJobFileName"
$GetAllServerCollectionLatencyInfoJobFilePath = "$ddlPath\$GetAllServerCollectionLatencyInfoJobFileName"

$GetAllServerSqlAgentJobsJobFilePath = "$ddlPath\$GetAllServerSqlAgentJobsJobFileName"
$GetAllServerDiskSpaceJobFilePath = "$ddlPath\$GetAllServerDiskSpaceJobFileName"
$GetAllServerLogSpaceConsumersJobFilePath = "$ddlPath\$GetAllServerLogSpaceConsumersJobFileName"
$GetAllServerTempdbSpaceUsageJobFilePath = "$ddlPath\$GetAllServerTempdbSpaceUsageJobFileName"
$GetAllServerAgHealthStateJobFilePath = "$ddlPath\$GetAllServerAgHealthStateJobFileName"
$GetAllServerServicesJobFilePath = "$ddlPath\$GetAllServerServicesJobFileName"
$GetAllServerBackupsJobFilePath = "$ddlPath\$GetAllServerBackupsJobFileName"
$GetAllServerAlertHistoryJobFilePath = "$ddlPath\$GetAllServerAlertHistoryJobFileName"

$CheckInstanceAvailabilityJobFilePath = "$ddlPath\$CheckInstanceAvailabilityJobFileName"
$UpdateSQLAgentJobsThresholdFilePath = "$ddlPath\$UpdateSQLAgentJobsThresholdFileName"
$CollectLoginExpirationInfoJobFilePath = "$ddlPath\$CollectLoginExpirationInfoJobFileName"
$PopulateInventoryTablesJobFilePath = "$ddlPath\$PopulateInventoryTablesJobFileName"
$SendLoginExpiryEmailsJobFilePath = "$ddlPath\$SendLoginExpiryEmailsJobFileName"
$ComputeAllServerVolatileInfoHistoryHourlyJobFilePath = "$ddlPath\$ComputeAllServerVolatileInfoHistoryHourlyJobFileName"
$StopStuckSQLMonitorJobsJobFilePath = "$ddlPath\$StopStuckSQLMonitorJobsJobFileName"
$UpdateSqlServerVersionsJobFilePath = "$ddlPath\$UpdateSqlServerVersionsJobFileName"


"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$ddlPath = '$ddlPath'"
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$psScriptPath = '$psScriptPath'"

"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Import dbatools module.."
Import-Module dbatools -Verbose:$false -Debug:$false

# Compute steps to execute
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Compute Steps to execute.."
$StartAtStepNumber = 1
$StopAtStepNumber = $AllSteps.Count+1

if(-not [String]::IsNullOrEmpty($StartAtStep)) {
    [int]$StartAtStepNumber = $StartAtStep -replace "__\w+", ''
}
if(-not [String]::IsNullOrEmpty($StopAtStep)) {
    [int]$StopAtStepNumber = $StopAtStep -replace "__\w+", ''
}


$Steps2Execute = @()
$Steps2ExecuteRaw = @()
if(-not [String]::IsNullOrEmpty($SkipSteps)) {
    $Steps2ExecuteRaw += Compare-Object -ReferenceObject $AllSteps -DifferenceObject $SkipSteps | Where-Object {$_.SideIndicator -eq '<='} | Select-Object -ExpandProperty InputObject -Unique
}
else {
    $Steps2ExecuteRaw += $AllSteps
}

$Steps2Execute += $Steps2ExecuteRaw | ForEach-Object { 
                            $currentStepNumber = [int]$($_ -replace "__\w+", '');
                            $passThrough = $true;
                            if( -not ($currentStepNumber -ge $StartAtStepNumber -and $currentStepNumber -le $StopAtStepNumber) ) {
                                $passThrough = $false
                            }
                            if( $passThrough -and ($SkipAllJobs -and $_ -like '*__CreateJob*') ) {
                                $passThrough = $false
                            }
                            if($passThrough) {$_}
                        }

if($OnlySteps.Count -gt 0) {
    # Override Steps to Execute by OnlySteps parameter value
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Override `$Steps2Execute with value from `$OnlySteps.."
    $Steps2Execute = $OnlySteps
}

"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$StartAtStep -> $StartAtStep.."
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$StopAtStep -> $StopAtStep.."
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Total steps to execute -> $($Steps2Execute.Count)."


# Fail-fast: GrafanaLoginPassword required on inventory when step 56 runs
if ($null -eq $GrafanaLoginPassword -and
    '56__GrafanaLogin' -in $Steps2Execute -and
    $SqlInstanceToBaselineWithOutPort -eq $InventoryServerWithOutPort) {
    throw "GrafanaLoginPassword is required for inventory-server installs that include step 56__GrafanaLogin. Supply -GrafanaLoginPassword."
}

# Setup SQL Connection for Inventory $conInventoryServer
try {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "[Connect-DbaInstance] Create connection for InventoryServer '$InventoryServer'.."
    $conInventoryServer = Connect-DbaInstance -SqlInstance $InventoryServer -Database master -ClientName $clientName `
                                        -SqlCredential $SqlCredential -TrustServerCertificate -EncryptConnection -ErrorAction Stop -Verbose:$false -Debug:$false
}
catch {
    $errMessage = "Connect-DbaInstance => $($_.Exception.Message)"
    
    if ($ReturnInlineErrorMessage) 
    {
        if([String]::IsNullOrEmpty($SqlCredential)) {
            $errMessage = "SQL Connection to `$InventoryServer ($InventoryServer) failed.`nKindly provide SqlCredentials.`n$errMessage.."
        } else {
            $errMessage = "SQL Connection to `$InventoryServer ($InventoryServer) failed.`nProvided SqlCredentials seems to be NOT working.`n$errMessage.."
        }
        
        $errMessage | Write-Error
    }
    else
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "SQL Connection to `$InventoryServer ($InventoryServer) failed." | Write-Host -ForegroundColor Red
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "$errMessage." | Write-Host -ForegroundColor Red
        if([String]::IsNullOrEmpty($SqlCredential)) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly provide SqlCredentials." | Write-Host -ForegroundColor Red
        } else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Provided SqlCredentials seems to be NOT working." | Write-Host -ForegroundColor Red
        }

        Write-Error "Stop here. Fix above issue."
    }
}


# Construct Instance Details Query
$sqlInstanceDetails = @"
if OBJECT_ID('dbo.instance_details') is not null
begin
	;with cte_instance_details as (
		select * from dbo.instance_details id
		where sql_instance = '$SqlInstanceToBaselineWithOutPort'
		$(if([String]::IsNullOrEmpty($HostName)) { '--' })and [host_name] = '$HostName'
	)
	select t.*, id.*
	from cte_instance_details id
	full outer join (values ('Inventory',object_id('dbo.instance_details'))) t(source, id_object_id)
		on 1=1;
end
else
	select *
	from (values ('Inventory',object_id('dbo.instance_details'))) t(source, id_object_id);
"@

if($verbose) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$sqlInstanceDetails query => `n$sqlInstanceDetails`n" | Write-Host -ForegroundColor Cyan
}


# From Inventory, get dbo.instance_details data
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Fetching [dbo].[instance_details] data from Inventory ([$InventoryServer].[$InventoryDatabase]).."
$instanceDetailsFromInventory = @()
try {
    $instanceDetailsFromInventory += $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlInstanceDetails -Verbose:$false -Debug:$false
}
catch {
    $errMessage = $_.Exception.Message

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Could not fetch details from [$InventoryServer].[$InventoryDatabase].[dbo].[instance_details] info."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly check the error => `n$errMessage" | Write-Host -ForegroundColor Red
    Write-Error "Stop here. Fix above issue."
}

# Perform all validations based on data fetched from Inventory
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Performing all kinds of validations based on data of `$instanceDetailsFromInventory."
if($verbose) {
    $instanceDetailsFromInventory | Select-Object source, sql_instance, @{l='port';e={$_.sql_instance_port}}, host_name, `
            database, sqlmonitor_version, created_date_utc, is_enabled, is_available, @{l='is_link_working';e={$_.is_linked_server_working}} | Format-Table -AutoSize
    $instanceDetailsFromInventory | Select-Object source, @{l='tsql_jobs_server';e={$_.collector_tsql_jobs_server}}, `
            @{l='powershell_jobs_server';e={$_.collector_powershell_jobs_server}}, @{l='data_destination';e={$_.data_destination_sql_instance}}, `
            more_info | Format-Table -AutoSize
}

# Local Parameters
$countHostRecords = 0
$isInventoryPresent = $false
$hasMultipleHostRecords = $false
$hasActiveHostRecords = $false
$hasMultipleActiveHostRecords = $false
$activeHostRecordsFromInventory = @()
$countActiveHostRecordsFromInventory = 0
$isAvailable = $false
$isLinkedServerWorking = $false
$sqlInstanceToBaselineWithPortFromInventory = $SqlInstanceToBaseline
$sqlInstancePortFromInventory = $Port4SqlInstanceToBaseline
[String]$sqlInstancePortFromInventory = $null
[String]$collectorTsqlJobsServerFromInventory = $null
[String]$collectorPSJobsServerFromInventory = $null
[String]$dataDestinationServerFromInventory = $null
$hostNameFromInventory = $HostName
$moreInfoFromInventory = $null
[String]$dbaGroupMailIdFromInventory = $null


if(-not [String]::IsNullOrEmpty($instanceDetailsFromInventory[0].id_object_id)) {
    $isInventoryPresent = $true
}
if(-not [String]::IsNullOrEmpty($instanceDetailsFromInventory[0].sql_instance)) {
    $isUpgradeScenario = $true
    $countHostRecords = $instanceDetailsFromInventory.Count
}
if($countHostRecords -gt 1) {
    $hasMultipleHostRecords = $true
}
if($isUpgradeScenario) {
    $activeHostRecordsFromInventory += $instanceDetailsFromInventory | Where-Object {$_.is_enabled -eq $true}
    $countActiveHostRecordsFromInventory = $activeHostRecordsFromInventory.Count
}
if($countActiveHostRecordsFromInventory -gt 0) {
    $hasActiveHostRecords = $true
    
    if($countActiveHostRecordsFromInventory -gt 1) {
        $hasMultipleActiveHostRecords = $true
    }
}

# Ensure 1st server to baseline is Inventory Server
if(($isInventoryPresent -eq $false) -and ($SqlInstanceToBaseline -ne $InventoryServer)) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly baseline your Inventory server [$InventoryServer] first." | Write-Host -ForegroundColor Red
    Start-Sleep -Seconds 1
    Write-Error "Stop here. Fix above issue."
}

# Print if Upgrade or Fresh setup
if($isUpgradeScenario) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Server [$SqlInstanceToBaseline] found in Inventory [$InventoryServer].[$InventoryDatabase]."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "So assuming upgrade scenario for [$SqlInstanceToBaseline]."
}
else {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Server [$SqlInstanceToBaseline] not found in Inventory [$InventoryServer].[$InventoryDatabase]."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "So assuming fresh baseline with SQLMonitor."
}

# If more than one active hosts present, throw error
if($countActiveHostRecordsFromInventory -gt 1) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "More than one active Hosts found for [$SqlInstanceToBaselineWithOutPort] in dbo.instance_details." | Write-Host -ForegroundColor Red
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly provide HostName parameter value." | Write-Host -ForegroundColor Red
    Start-Sleep -Seconds 1
    Write-Error "Stop here. Fix above issue."
}

# If no active, but more than one inactive hosts present, throw error
if( ($countActiveHostRecordsFromInventory -eq 0) -and ($countHostRecords -gt 1) ) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "More than one inactive Hosts found for [$SqlInstanceToBaselineWithOutPort] in dbo.instance_details." | Write-Host -ForegroundColor Red
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly provide HostName parameter value." | Write-Host -ForegroundColor Red
    Start-Sleep -Seconds 1
    Write-Error "Stop here. Fix above issue."
}

# If only one active host found, then pull all details from active host record
if($hasActiveHostRecords -and $countActiveHostRecordsFromInventory -eq 1) 
{
    $isAvailable = $activeHostRecordsFromInventory[0].is_available
    $isLinkedServerWorking = $activeHostRecordsFromInventory[0].is_linked_server_working
    $sqlInstancePortFromInventory = $activeHostRecordsFromInventory[0].sql_instance_port
    $collectorTsqlJobsServerFromInventory = $activeHostRecordsFromInventory[0].collector_tsql_jobs_server
    $collectorPSJobsServerFromInventory = $activeHostRecordsFromInventory[0].collector_powershell_jobs_server
    $dataDestinationServerFromInventory = $activeHostRecordsFromInventory[0].data_destination_sql_instance
    $hostNameFromInventory = $activeHostRecordsFromInventory[0].host_name
    $moreInfoFromInventory = $activeHostRecordsFromInventory[0].more_info
    $dbaGroupMailIdFromInventory = $activeHostRecordsFromInventory[0].dba_group_mail_id
    $dbaDatabaseFromInventory = $activeHostRecordsFromInventory[0].database

    # If $isAvailable, then Port information is correct in dbo.instance_details
    if ($isAvailable) {
        if (-not [String]::IsNullOrEmpty($sqlInstancePortFromInventory)) {
            $SqlInstanceToBaselineWithPortFromInventory = "$SqlInstanceToBaselineWithOutPort,$sqlInstancePortFromInventory"
        }
    }

    if([String]::IsNullOrEmpty($SqlInstanceAsDataDestination)) {
        $SqlInstanceAsDataDestination = $dataDestinationServerFromInventory
    }
    if([String]::IsNullOrEmpty($SqlInstanceForTsqlJobs)) {
        $SqlInstanceForTsqlJobs = $collectorTsqlJobsServerFromInventory
    }
    if([String]::IsNullOrEmpty($SqlInstanceForPowershellJobs)) {
        $SqlInstanceForPowershellJobs = $collectorPSJobsServerFromInventory
    }
    if($DbaDatabase -ne $dbaDatabaseFromInventory) {
        $DbaDatabase = $dbaDatabaseFromInventory
    }
}

# If no active host found, then pull all details from inactive host record
if($hasActiveHostRecords -eq $false -and $countHostRecords -eq 1) 
{
    $isAvailable = $instanceDetailsFromInventory[0].is_available
    $isLinkedServerWorking = $instanceDetailsFromInventory[0].is_linked_server_working
    $sqlInstancePortFromInventory = $instanceDetailsFromInventory[0].sql_instance_port
    $collectorTsqlJobsServerFromInventory = $instanceDetailsFromInventory[0].collector_tsql_jobs_server
    $collectorPSJobsServerFromInventory = $instanceDetailsFromInventory[0].collector_powershell_jobs_server
    $dataDestinationServerFromInventory = $instanceDetailsFromInventory[0].data_destination_sql_instance
    $dbaGroupMailIdFromInventory = $instanceDetailsFromInventory[0].dba_group_mail_id
    $dbaDatabaseFromInventory = $instanceDetailsFromInventory[0].database

    if([String]::IsNullOrEmpty($SqlInstanceAsDataDestination)) {
        $SqlInstanceAsDataDestination = $dataDestinationServerFromInventory
    }
    if([String]::IsNullOrEmpty($SqlInstanceForTsqlJobs)) {
        $SqlInstanceForTsqlJobs = $collectorTsqlJobsServerFromInventory
    }
    if([String]::IsNullOrEmpty($SqlInstanceForPowershellJobs)) {
        $SqlInstanceForPowershellJobs = $collectorPSJobsServerFromInventory
    }
    if($DbaDatabase -ne $dbaDatabaseFromInventory) {
        $DbaDatabase = $dbaDatabaseFromInventory
    }
}


# Attempt 01: Setup SQL Connection for SqlInstanceToBaseline
[String]$errMessageAttempt1 = $null
$conSqlInstanceToBaseline = $null
try {
    if ($SqlInstanceToBaselineWithOutPort -ne $InventoryServerWithOutPort) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "[Connect-DbaInstance] First attempt to create connection for `$SqlInstanceToBaseline ($SqlInstanceToBaseline).."
        $conSqlInstanceToBaseline = Connect-DbaInstance -SqlInstance $SqlInstanceToBaseline -Database master -ClientName $clientName `
                                                        -SqlCredential $SqlCredential -TrustServerCertificate -EncryptConnection -ErrorAction Stop -Verbose:$false -Debug:$false
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$conSqlInstanceToBaseline connection created successfully in Attempt1."
        if($verbose) {
            $conSqlInstanceToBaseline
        }
    }
    else {
        $conSqlInstanceToBaseline = $conInventoryServer
    }
}
catch {
    $errMessageAttempt1 = $_.Exception.Message
    
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "[Connect-DbaInstance] Failed to create connection to `$SqlInstanceToBaseline ($SqlInstanceToBaseline) in first attempt." | Write-Host -ForegroundColor Red
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "$errMessageAttempt1" | Write-Host -ForegroundColor Red
}

# Attempt 02: Setup SQL Connection for SqlInstanceToBaseline
[String]$errMessageAttempt2 = $null
#if([String]::IsNullOrEmpty($conSqlInstanceToBaseline) -and $isAvailable) 
if([String]::IsNullOrEmpty($conSqlInstanceToBaseline)) 
{
    try {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "[Connect-DbaInstance] Second attempt to create connection for `$SqlInstanceToBaseline ($SqlInstanceToBaselineWithPortFromInventory).."
        $conSqlInstanceToBaseline = Connect-DbaInstance -SqlInstance $SqlInstanceToBaselineWithPortFromInventory -Database master -ClientName $clientName `
                                                        -SqlCredential $SqlCredential -TrustServerCertificate -EncryptConnection -ErrorAction Stop -Verbose:$false -Debug:$false
        
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$conSqlInstanceToBaseline connection created successfully."
        $SqlInstanceToBaseline = $SqlInstanceToBaselineWithPortFromInventory

        if($SqlInstanceToBaseline -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
            $Port4SqlInstanceToBaseline = $Matches['PortNo']
            $SqlInstanceToBaselineWithOutPort = $Matches['SqlInstance']
        }

        if($verbose) {
            $conSqlInstanceToBaseline
        }
    }
    catch {
        $errMessageAttempt2 = $_.Exception.Message
    
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "[Connect-DbaInstance] Failed to create connection to `$SqlInstanceToBaseline ($SqlInstanceToBaselineWithPortFromInventory) in second attempt." | Write-Host -ForegroundColor Red
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "$errMessageAttempt2" | Write-Host -ForegroundColor Red
    }
}

# If both connection attempt failed for $SqlInstanceToBaseline, then suggest SqlCredentials
if([String]::IsNullOrEmpty($conSqlInstanceToBaseline)) {
    if([String]::IsNullOrEmpty($SqlCredential)) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly try using 'SqlCredential' parameter." | Write-Host -ForegroundColor Red
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "This credential should have sysadmin on InventoryServer & SqlInstanceToBaseline." | Write-Host -ForegroundColor Red
        Start-Sleep -Seconds 1
        Write-Error "Stop here. Fix above issue."
    }
    else {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly investigate further based on above error messages to resolve SQL Connectivity." | Write-Host -ForegroundColor Red
        Start-Sleep -Seconds 1
        Write-Error "Stop here. Fix above issue."
    }
}

# Extract port no for DataDestination server
$Port4SqlInstanceAsDataDestination = $null
if($SqlInstanceAsDataDestination -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
    $Port4SqlInstanceAsDataDestination = $Matches['PortNo']
    $SqlInstanceAsDataDestinationWithOutPort = $Matches['SqlInstance']
}
if($SqlInstanceToBaselineWithOutPort -eq $SqlInstanceAsDataDestinationWithOutPort) {
    $SqlInstanceAsDataDestination = $SqlInstanceToBaseline
    $Port4SqlInstanceAsDataDestination = $Port4SqlInstanceToBaseline
}

# Extract port no for DataDestination server
$Port4SqlInstanceForTsqlJobs = $null
if($SqlInstanceForTsqlJobs -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
    $Port4SqlInstanceForTsqlJobs = $Matches['PortNo']
    $SqlInstanceForTsqlJobsWithOutPort = $Matches['SqlInstance']
}
if($SqlInstanceToBaselineWithOutPort -eq $SqlInstanceForTsqlJobsWithOutPort) {
    $SqlInstanceForTsqlJobs = $SqlInstanceToBaseline
    $Port4SqlInstanceForTsqlJobs = $Port4SqlInstanceToBaseline
}

# Extract port no for DataDestination server
$Port4SqlInstanceForPowershellJobs = $null
if($SqlInstanceForPowershellJobs -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
    $Port4SqlInstanceForPowershellJobs = $Matches['PortNo']
    $SqlInstanceForPowershellJobsWithOutPort = $Matches['SqlInstance']
}
if($SqlInstanceToBaselineWithOutPort -eq $SqlInstanceForPowershellJobsWithOutPort) {
    $SqlInstanceForPowershellJobs = $SqlInstanceToBaseline
    $Port4SqlInstanceForPowershellJobs = $Port4SqlInstanceToBaseline
}

# Check if PortNo is specified for $dataDestinationServerFromInventory
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract port no from `$dataDestinationServerFromInventory ($dataDestinationServerFromInventory) if provided."
$Port4DataDestinationServerFromInventory = $null
$DataDestinationServerFromInventoryWithoutPort = $dataDestinationServerFromInventory
if($dataDestinationServerFromInventory -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
    $Port4DataDestinationServerFromInventory = $Matches['PortNo']
    $DataDestinationServerFromInventoryWithoutPort = $Matches['SqlInstance']
}

# Check if PortNo is specified for $collectorPSJobsServerFromInventory
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract port no from `$collectorPSJobsServerFromInventory ($collectorPSJobsServerFromInventory) if provided."
$Port4collectorPSJobsServerFromInventory = $null
$collectorPSJobsServerFromInventoryWithoutPort = $collectorPSJobsServerFromInventory
if($collectorPSJobsServerFromInventory -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
    $Port4collectorPSJobsServerFromInventory = $Matches['PortNo']
    $collectorPSJobsServerFromInventoryWithoutPort = $Matches['SqlInstance']
}

# Check if PortNo is specified for $collectorTsqlJobsServerFromInventory
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract port no from `$collectorTsqlJobsServerFromInventory ($collectorTsqlJobsServerFromInventory) if provided."
$Port4collectorTsqlJobsServerFromInventory = $null
$collectorTsqlJobsServerFromInventoryWithoutPort = $collectorTsqlJobsServerFromInventory
if($collectorTsqlJobsServerFromInventory -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
    $Port4collectorTsqlJobsServerFromInventory = $Matches['PortNo']
    $collectorTsqlJobsServerFromInventoryWithoutPort = $Matches['SqlInstance']
}

# Construct Instance Details Query for BaselineServer
$sqlInstanceDetails = @"
if OBJECT_ID('dbo.instance_details') is not null
begin
	;with cte_instance_details as (
		select * from dbo.instance_details id
		where sql_instance = '$SqlInstanceToBaselineWithOutPort'
		$(if([String]::IsNullOrEmpty($HostName)) { '--' })and [host_name] = '$HostName'
	)
	select t.*, id.*
	from cte_instance_details id
	full outer join (values ('BaselineServer',object_id('dbo.instance_details'))) t(source, id_object_id)
		on 1=1;
end
else
	select *
	from (values ('BaselineServer',object_id('dbo.instance_details'))) t(source, id_object_id);
"@

if($verbose) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$sqlInstanceDetails query => `n$sqlInstanceDetails`n" | Write-Host -ForegroundColor Cyan
}

# From SqlInstanceToBaseline, get dbo.instance_details data
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Fetching [dbo].[instance_details] data from BaselineServer ([$SqlInstanceToBaseline].[$DbaDatabase]).."
$instanceDetailsFromBaselineServer = @()
try {
    $instanceDetailsFromBaselineServer += $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlInstanceDetails -Verbose:$false -Debug:$false
}
catch {
    $errMessage = $_.Exception.Message

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Could not fetch details from [$SqlInstanceToBaseline].[$DbaDatabase].[dbo].[instance_details] info."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly check the error => `n$errMessage" | Write-Host -ForegroundColor Red
    Write-Error "Stop here. Fix above issue."
}

# Perform all validations based on data fetched from Inventory
if($verbose) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Validate data of dbo.instance_details from BaselineServer ([$SqlInstanceToBaseline].[$DbaDatabase])."
    $instanceDetailsFromBaselineServer | Select-Object source, sql_instance, @{l='port';e={$_.sql_instance_port}}, host_name, `
            database, sqlmonitor_version, created_date_utc, is_enabled, is_available, @{l='is_link_working';e={$_.is_linked_server_working}} | Format-Table -AutoSize
    $instanceDetailsFromBaselineServer | Select-Object source, @{l='tsql_jobs_server';e={$_.collector_tsql_jobs_server}}, `
            @{l='powershell_jobs_server';e={$_.collector_powershell_jobs_server}}, @{l='data_destination';e={$_.data_destination_sql_instance}}, `
            more_info | Format-Table -AutoSize
}


# Get Server Info
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Fetching basic server info for `$SqlInstanceToBaseline => [$SqlInstanceToBaseline].."
$sqlServerInfo = @"
DECLARE @Domain NVARCHAR(255);
begin try
	EXEC master.dbo.xp_regread 'HKEY_LOCAL_MACHINE', 'SYSTEM\CurrentControlSet\services\Tcpip\Parameters', N'Domain',@Domain OUTPUT;
end try
begin catch
	print 'some erorr accessing registry'
end catch

;with server_services as (
	select *
	from sys.dm_server_services 
	where servicename like 'SQL Server (%)'
	or servicename like 'SQL Server Agent (%)'
)
select	[domain] = default_domain(),
		[domain_reg] = @Domain,
		--[ip] = CONNECTIONPROPERTY('local_net_address'),
		[at_server_name] = @@SERVERNAME,
		[MachineName] = serverproperty('MachineName'),
		[ServerName] = serverproperty('ServerName'),
		[host_name] = COALESCE(SERVERPROPERTY('ComputerNamePhysicalNetBIOS'),SERVERPROPERTY('ServerName')),
		[ProductVersion] = SERVERPROPERTY('ProductVersion'),
		[service_name_str] = servicename,
		[service_name] = case	when @@servicename = 'MSSQLSERVER' and servicename like 'SQL Server (%)' then 'MSSQLSERVER'
								when @@servicename = 'MSSQLSERVER' and servicename like 'SQL Server Agent (%)' then 'SQLSERVERAGENT'
								when @@servicename <> 'MSSQLSERVER' and servicename like 'SQL Server (%)' then 'MSSQL$'+@@servicename
								when @@servicename <> 'MSSQLSERVER' and servicename like 'SQL Server Agent (%)' then 'SQLAgent'+@@servicename
								else 'MSSQL$'+@@servicename end,
        service_account,
		[Edition] = SERVERPROPERTY('Edition'),
		[EngineEdition] = SERVERPROPERTY('EngineEdition'),
        [is_clustered] = case when exists (select 1 from sys.dm_os_cluster_nodes) then 1 else 0 end
from (values ('Basic-Details')) d (RunningQuery)
full outer join	server_services ss
	on 1=1;
"@
try {
    $resultServerInfo = @()
    $resultServerInfo += $conSqlInstanceToBaseline | Invoke-DbaQuery -Query $sqlServerInfo -EnableException -Verbose:$false -Debug:$false

    if ($resultServerInfo.Count -eq 1) {
        $IsManagedInstance = $true
        $InstanceScopeFeaturesOnly = $true

        $dbServiceInfo = $resultServerInfo
        $agentServiceInfo = $resultServerInfo
    }
    else {
        $dbServiceInfo = $resultServerInfo | Where-Object {$_.service_name_str -like "SQL Server (*)"}
        $agentServiceInfo = $resultServerInfo | Where-Object {$_.service_name_str -like "SQL Server Agent (*)"}
    }

    $resultServerInfo | Select domain, domain_reg, at_server_name, MachineName, ServerName, host_name, ProductVersion | Format-Table -AutoSize
    $resultServerInfo | Select service_account, Edition, is_clustered, service_name_str, service_name | Format-Table -AutoSize
}
catch {
    $errMessage = $_
    
    if ($ReturnInlineErrorMessage) 
    {
        if([String]::IsNullOrEmpty($SqlCredential)) {
            $errMessage = "SQL Connection to [$SqlInstanceToBaseline] failed.`nKindly provide SqlCredentials.`n$($errMessage.Exception.Message).."
        } else {
            $errMessage = "SQL Connection to [$SqlInstanceToBaseline] failed.`nProvided SqlCredentials seems to be NOT working.`n$($errMessage.Exception.Message).."
        }

        $errMessage | Write-Error
    }
    else
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "SQL Connection to [$SqlInstanceToBaseline] failed."
        if([String]::IsNullOrEmpty($SqlCredential)) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly provide SqlCredentials." | Write-Host -ForegroundColor Red
        } else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Provided SqlCredentials seems to be NOT working." | Write-Host -ForegroundColor Red
        }
        Write-Error "Stop here. Fix above issue."
    }
}

# Check if Azure SQL MI
if($dbServiceInfo.Edition -in @('SQL Azure') -or $dbServiceInfo.Edition -eq 8) 
{
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Azure SQL Managed Instance detected." | Write-Host -ForegroundColor Yellow

    $IsManagedInstance = $true
    $InstanceScopeFeaturesOnly = $true
}


# Extract Version Info & Partition Info
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Major & Minor Version of SQL Server.."

[bool]$IsNonPartitioned = $true
if($dbServiceInfo.ProductVersion -match "(?'MajorVersion'\d+)\.\d+\.(?'MinorVersion'\d+)\.\d+")
{
    [int]$MajorVersion = $Matches['MajorVersion']
    [int]$MinorVersion = $Matches['MinorVersion']
    [bool]$IsCompressionSupported = $false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Figure out Edition, `$IsNonPartitioned & `$IsCompressionSupported.."
    if($dbServiceInfo.Edition -like 'Enterprise*' -or $dbServiceInfo.Edition -like 'Developer*') {
        $IsNonPartitioned = $false
    }
    elseif($dbServiceInfo.Edition -like 'Standard*' -or $dbServiceInfo.Edition -like 'Express*')
    {
        if($MajorVersion -gt 13) {
            $IsNonPartitioned = $false
        }
        elseif ($MajorVersion -eq 13 -and $MinorVersion -ge 4001) {
            $IsNonPartitioned = $false
        }
    }

    if($MajorVersion -ge 13) {
        $IsCompressionSupported = $true
    }
    elseif ($dbServiceInfo.Edition -like 'Enterprise*' -or $dbServiceInfo.Edition -like 'Developer*') {
        $IsCompressionSupported = $true
    }
}

# Extract domain & isClustered property
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Figure out `$domain of server.."
[bool]$isClustered = $dbServiceInfo.is_clustered
[string]$domain = $dbServiceInfo.domain_reg
if([String]::IsNullOrEmpty($domain)) {
    if($dbServiceInfo.domain -ne 'WORKGROUP' -and (-not [String]::IsNullOrEmpty($dbServiceInfo.domain))) {
        $domain = $dbServiceInfo.domain+'.com'
    }
    else {
        if([String]::IsNullOrEmpty($dbServiceInfo.domain)) {
            $domain = 'WORKGROUP'
        }
        else {
            $domain = $dbServiceInfo.domain
        }
    }
}


# Initialize $instanceDetails from $instanceDetailsFromInventory or $instanceDetailsFromBaselineServer
$instanceDetails = @()
if(-not [String]::IsNullOrEmpty($instanceDetailsFromInventory)) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Initialize `$instanceDetails using `$instanceDetailsFromInventory.."    
    $instanceDetails += $instanceDetailsFromInventory
} else {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Initialize `$instanceDetails using `$instanceDetailsFromBaselineServer.."    
    $instanceDetails += $instanceDetailsFromBaselineServer
}

# If instance details found, then use same to initiate empty parameters
if ( $countHostRecords -gt 0 ) 
{
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Instance details found in [dbo].[instance_details]."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Using available info from this table."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Incase details of [dbo].[instance_details] are outdated, `n`t`t`t`tconsider updating same 1st on Inventory & Local Instance both."
    $instanceDetails | Format-Table -AutoSize

    # If no DBA Mail provided, then fetch from dbo.instance_details
    if($DbaGroupMailId.Count -eq 0) {
        $DbaGroupMailId += $($dbaGroupMailIdFromInventory -split ';')
    }

    # If SQL Port is not provided, but should be present
    if ([String]::IsNullOrEmpty($Port4SqlInstanceToBaseline)) {
        if (-not [String]::IsNullOrEmpty($sqlInstancePortFromInventory)) {
            $Port4SqlInstanceToBaseline = $sqlInstancePortFromInventory
            $SqlInstanceToBaseline = $sqlInstanceToBaselineWithPortFromInventory
        }
    }

    if( ($RemoteSQLMonitorPath -ne $instanceDetails[0].sqlmonitor_script_path) -and $RemoteSQLMonitorPath -ne 'C:\SQLMonitor' ) {
        if ($ReturnInlineErrorMessage) {
            "RemoteSQLMonitorPath parameter value does not match with dbo.instance_details. `nConsider updating details of dbo.instance_details on Inventory & Local Instance both." | Write-Error
        }
        else {            
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "RemoteSQLMonitorPath parameter value does not match with dbo.instance_details." | Write-Host -ForegroundColor Red
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Consider updating details of dbo.instance_details on Inventory & Local Instance both."
        
            "STOP here, and fix above issue." | Write-Error
        }
    }else {
        if( ($RemoteSQLMonitorPath -ne $instanceDetails[0].sqlmonitor_script_path) -and $RemoteSQLMonitorPath -eq 'C:\SQLMonitor' ) {
            $RemoteSQLMonitorPath = $instanceDetails[0].sqlmonitor_script_path
        }
    }

    if ([String]::IsNullOrEmpty($SqlInstanceAsDataDestination)) {
        $SqlInstanceAsDataDestination = $instanceDetails[0].data_destination_sql_instance
    }
    else {
        if( $SqlInstanceAsDataDestination -ne $instanceDetails[0].data_destination_sql_instance ) {
            if ($ReturnInlineErrorMessage) {
                "SqlInstanceAsDataDestination parameter value does not match with dbo.instance_details. `nConsider updating details of dbo.instance_details on Inventory & Local Instance both." | Write-Error
            }
            else {            
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "SqlInstanceAsDataDestination parameter value does not match with dbo.instance_details." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Consider updating details of dbo.instance_details on Inventory & Local Instance both."
        
                "STOP here, and fix above issue." | Write-Error
            }
        }
    }

    if ([String]::IsNullOrEmpty($SqlInstanceForPowershellJobs)) {
        $SqlInstanceForPowershellJobs = $instanceDetails[0].collector_powershell_jobs_server
    }
    else {
        if( $SqlInstanceForPowershellJobs -ne $instanceDetails[0].collector_powershell_jobs_server ) {
            if ($ReturnInlineErrorMessage) {
                "SqlInstanceForPowershellJobs parameter value does not match with dbo.instance_details. `nConsider updating details of dbo.instance_details on Inventory & Local Instance both." | Write-Error
            }
            else {            
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "SqlInstanceForPowershellJobs parameter value does not match with dbo.instance_details." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Consider updating details of dbo.instance_details on Inventory & Local Instance both."
        
                "STOP here, and fix above issue." | Write-Error
            }
        }
    }

    if ($DbaGroupMailId.Count -eq 0) {
        $DbaGroupMailId += 'dba_team@gmail.com'
    }
    
    if ([String]::IsNullOrEmpty($SqlInstanceForTsqlJobs)) {
        $SqlInstanceForTsqlJobs = $instanceDetails[0].collector_tsql_jobs_server
    }
    else {
        if( $SqlInstanceForTsqlJobs -ne $instanceDetails[0].collector_tsql_jobs_server ) {
            if ($ReturnInlineErrorMessage) {
                "SqlInstanceForTsqlJobs parameter value does not match with dbo.instance_details. `nConsider updating details of dbo.instance_details on Inventory & Local Instance both." | Write-Error
            }
            else {            
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "SqlInstanceForTsqlJobs parameter value does not match with dbo.instance_details." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Consider updating details of dbo.instance_details on Inventory & Local Instance both."
        
                "STOP here, and fix above issue." | Write-Error
            }
        }
    }

    # Get additional parameter settings
    if (-not [String]::IsNullOrEmpty($instanceDetails[0].more_info)) {
        $moreInfoJSON = $instanceDetails[0].more_info
        $moreInfo = $moreInfoJSON | ConvertFrom-Json

        if( ($moreInfo.InstanceScopeFeaturesOnly -eq $true) -and ($InstanceScopeFeaturesOnly -ne $moreInfo.InstanceScopeFeaturesOnly) ) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "InstanceScopeFeaturesOnly parameter value does not match with dbo.instance_details.more_info." | Write-Host -ForegroundColor Yellow
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "So preferring Inventory Based parameter values over local parameter values passed."
            $InstanceScopeFeaturesOnly = $true

            $SkipPowerShellJobs = $true
            $SkipRDPSessionSteps = $true
            $ForceTSQLStepType4TsqlJobs = $true
            $SkipWindowsAdminAccessTest = $true

            $PowerShellJobSteps = $($PowerShellJobSteps | % {if($_ -notin @('13__CreateJobCollectDiskSpace','14__CreateJobCollectPerfmonData')){$_}});
            $TsqlJobSteps = $TsqlJobSteps + $( @('13__CreateJobCollectDiskSpace','14__CreateJobCollectPerfmonData') | % {if($_ -notin $TsqlJobSteps){$_}} );
        }

        if( ($moreInfo.ForceSetupOfTaskSchedulerJobs -eq $true) -and ($ForceSetupOfTaskSchedulerJobs -ne $moreInfo.ForceSetupOfTaskSchedulerJobs) ) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "ForceSetupOfTaskSchedulerJobs parameter value does not match with dbo.instance_details.more_info." | Write-Host -ForegroundColor Yellow
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "So preferring Inventory Based parameter values over local parameter values passed."
            $ForceSetupOfTaskSchedulerJobs = $true
        }

        if( ($moreInfo.HasCustomizedTsqlJobs -eq $true) -and ($HasCustomizedTsqlJobs -ne $moreInfo.HasCustomizedTsqlJobs) ) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "HasCustomizedTsqlJobs parameter value does not match with dbo.instance_details.more_info." #| Write-Host -ForegroundColor Yellow
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "So preferring Inventory Based parameter values over local parameter values passed."
            $HasCustomizedTsqlJobs = $true            
        }

        if( ($moreInfo.HasCustomizedPowerShellJobs -eq $true) -and ($HasCustomizedPowerShellJobs -ne $moreInfo.HasCustomizedPowerShellJobs) ) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "HasCustomizedPowerShellJobs parameter value does not match with dbo.instance_details.more_info." #| Write-Host -ForegroundColor Yellow
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "So preferring Inventory Based parameter values over local parameter values passed."
            $HasCustomizedPowerShellJobs = $true
        }

        if($OverrideCustomizedTsqlJobs) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "`$HasCustomizedTsqlJobs parameter value has been reset to false based on `$OverrideCustomizedTsqlJobs parameter." | Write-Host -ForegroundColor Yellow
            $HasCustomizedTsqlJobs = $false
        }
        if($OverrideCustomizedPowerShellJobs) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "`$HasCustomizedPowerShellJobs parameter value has been reset to false based on `$OverrideCustomizedPowerShellJobs parameter." | Write-Host -ForegroundColor Yellow
            $HasCustomizedPowerShellJobs = $false
        }
    }

    if( ($DbaDatabase -ne $instanceDetails[0].database) -and $DbaDatabase -ne 'DBA' ) {
        if ($ReturnInlineErrorMessage) {
            "DbaDatabase parameter value does not match with dbo.instance_details.`nConsider updating details of dbo.instance_details on Inventory & Local Instance both." | Write-Error
        }
        else {            
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "DbaDatabase parameter value does not match with dbo.instance_details." | Write-Host -ForegroundColor Red
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Consider updating details of dbo.instance_details on Inventory & Local Instance both."
        
            "STOP here, and fix above issue." | Write-Error
        }
    }else {
        if( ($DbaDatabase -ne $instanceDetails[0].database) -and $DbaDatabase -eq 'DBA' ) {
            $DbaDatabase = $instanceDetails[0].database
        }
    }

    if(-not $ConfirmValidationOfMultiInstance) {
        $ConfirmValidationOfMultiInstance = $true
    }

    $isUpgradeScenario = $true
}

# If fresh install, then set SkipMultiServerviewsUpgrade to False
if(-not $isUpgradeScenario) {
    $SkipMultiServerviewsUpgrade = $false

    $SkipMultiMailJobSteps = $false
}

# Remove jobs that have multiple emails for Upgrade Scenario unless Skip if False
if($SkipMultiMailJobSteps) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Skipping jobs (`$MultiMailJobSteps) that have mail notification for other folks apart from DBA.." | Write-Host -ForegroundColor Yellow
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "To avoid this skipping, kindly set SkipMultiMailJobSteps to false.." | Write-Host -ForegroundColor Yellow
    $Steps2Execute = $Steps2Execute | ForEach-Object {if($_ -notin $MultiMailJobSteps){$_}}
}

if($DbaGroupMailId -eq 'dba_team@gmail.com') {
    if ($ReturnInlineErrorMessage) {
		"Kindly provide a valid value for DbaGroupMailId parameter." | Write-Error
	}
	else {            
		"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly provide a valid value for DbaGroupMailId parameter." | Write-Host -ForegroundColor Red
        Write-Error "Stop here. Fix above issue."
    }
}


# Extract HostName of SqlInstance if NULL in parameter value
if([String]::IsNullOrEmpty($HostName)) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract HostName of SqlInstance.."
    $HostName = $dbServiceInfo.host_name;
}
else {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Validate HostName.."
    # If Sql Cluster, then host can be different
    # If not sql cluster, then host should be same
    if(-not $isClustered) {
        if($HostName -ne $dbServiceInfo.host_name) {
            if ($ReturnInlineErrorMessage) {
		        "Provided HostName does not match with SQLInstance host name." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Provided HostName does not match with SQLInstance host name." | Write-Host -ForegroundColor Red
                "STOP and check above error message" | Write-Error
            }
        }
    }
}

# Setup PSSession on HostName to setup Perfmon Data Collector. $ssn4PerfmonSetup
if( (($SkipRDPSessionSteps -eq $false) -or $ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) ) #-and ($HostName -ne $env:COMPUTERNAME)
{
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Create PSSession for host [$HostName].."
    $ssnHostName = $HostName

    # Try reaching server using HostName provided/detected, if fails, then use FQDN
    if (-not (Test-Connection -ComputerName $ssnHostName -Quiet -Count 1)) {
        if($domain -ne 'WORKGROUP.com' -and $domain -ne 'WORKGROUP') {
            $ssnHostName = "$HostName.$domain"
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Host [$HostName] not pingable. So trying FQDN form [$ssnHostName].."
        }
    }

    # Try reaching using FQDN, if fails & not a clustered instance, then use SqlInstanceToBaseline itself
    if ( (-not (Test-Connection -ComputerName $ssnHostName -Quiet -Count 1)) -and (-not $isClustered) ) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Host [$ssnHostName] not pingable. Since its not clustered instance, So trying `$SqlInstanceToBaseline parameter value itself.."
        $ssnHostName = $SqlInstanceToBaseline
    }

    # If not reachable after all attempts, raise error
    if ( -not (Test-Connection -ComputerName $ssnHostName -Quiet -Count 1) ) {        
        if($SkipPingCheck) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Host [$ssnHostName] not pingable." | Write-Host -ForegroundColor Cyan
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Skip ping validation to Host [$ssnHostName].." | Write-Host -ForegroundColor Cyan
        }else {
            if ($ReturnInlineErrorMessage) {
		        "Host [$ssnHostName] not pingable.`nKindly provide HostName either in FQDN or ipv4 format." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Host [$ssnHostName] not pingable." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly provide HostName either in FQDN or ipv4 format." | Write-Host -ForegroundColor Red
                "STOP and check above error message" | Write-Error      
            }  
        }
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$ssnHostName => '$ssnHostName'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Domain of SqlInstance being baselined => [$domain]"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Domain of current host => [$($env:USERDOMAIN)]"

    $ssn4PerfmonSetup = $null
    $errVariables = @()

    # First Attempt without Any credentials
    try {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Trying for PSSession on [$ssnHostName] normally.."
            $ssn4PerfmonSetup = New-PSSession -ComputerName $ssnHostName 
        }
    catch { $errVariables += $_ }

    # Second Attempt for Trusted Cross Domains
    if( [String]::IsNullOrEmpty($ssn4PerfmonSetup) ) {
        try { 
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Trying for PSSession on [$ssnHostName] assuming cross domain.."
            $ssn4PerfmonSetup = New-PSSession -ComputerName $ssnHostName -Authentication Negotiate 
        }
        catch { $errVariables += $_ }
    }

    # 3rd Attempt with Credentials
    if( [String]::IsNullOrEmpty($ssn4PerfmonSetup) -and (-not [String]::IsNullOrEmpty($WindowsCredential)) ) {
        try {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Attemping PSSession for [$ssnHostName] using provided WindowsCredentials.."
            $ssn4PerfmonSetup = New-PSSession -ComputerName $ssnHostName -Credential $WindowsCredential    
        }
        catch { $errVariables += $_ }

        if( [String]::IsNullOrEmpty($ssn4PerfmonSetup) ) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Attemping PSSession for [$ssnHostName] using provided WindowsCredentials with Negotiate attribute.."
            $ssn4PerfmonSetup = New-PSSession -ComputerName $ssnHostName -Credential $WindowsCredential -Authentication Negotiate
        }
    }

    if ( [String]::IsNullOrEmpty($ssn4PerfmonSetup) ) {
        if ($ReturnInlineErrorMessage) {
            "Provide WindowsCredential for accessing server [$ssnHostName] of domain '$domain'." | Write-Error
        }
        else {            
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Provide WindowsCredential for accessing server [$ssnHostName] of domain '$domain'." | Write-Host -ForegroundColor Red
            "STOP here, and fix above issue." | Write-Error
        }
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$ssn4PerfmonSetup PSSession for [$HostName].."
    $ssn4PerfmonSetup
    "`n"
}

Write-Verbose "Validate if IP Address has been provided for HostName"
# Validate if IPv4 is provided instead of DNS name for HostName
$pattern = "^([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])(\.([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])){3}$"
if($HostName  -match $pattern) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "IP address has been provided for `$HostName parameter."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Fetching DNS name for [$HostName].."
    $HostName = Invoke-Command -Session $ssn4PerfmonSetup -ScriptBlock { $env:COMPUTERNAME }
}

# Validate if FQDN is provided instead of single part HostName
$pattern = "(?=^.{4,253}$)(^((?!-)[a-zA-Z0-9-]{1,63}(?<!-)\.)+[a-zA-Z]{2,63}$)"
if($HostName  -match $pattern) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "FQDN has been provided for `$HostName parameter."
    
    if($SkipRDPSessionSteps -or $SkipPowerShellJobs -or $IsManagedInstance -or $InstanceScopeFeaturesOnly) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Skip DNS name fetch for [$HostName].."
    }
    else {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Fetching DNS name for [$HostName].."
        $HostName = Invoke-Command -Session $ssn4PerfmonSetup -ScriptBlock { $env:COMPUTERNAME }
    }
}

# Check No of SQL Services on HostName
Write-Verbose "Check No of SQL Services on HostName"
if( ($SkipPowerShellJobs -eq $false) -or ('20__CreateJobRemoveXEventFiles' -in $Steps2Execute) )
{
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Check for number of SQLServices on [$HostName].."

    $sqlServicesOnHost = @()

    # Localhost system
    if( $HostName -eq $env:COMPUTERNAME ) {
        $sqlServicesOnHost += Get-Service MSSQL* | Where-Object {$_.DisplayName -like 'SQL Server (*)' -and $_.StartType -ne 'Disabled'}
    }
    
    # Remote host
    if($HostName -ne $env:COMPUTERNAME)
    {
        # if pssession is null
        if([String]::IsNullOrEmpty($ssn4PerfmonSetup)) 
        {
            # If Destination instance is not provided, throw error
            if([String]::IsNullOrEmpty($SqlInstanceAsDataDestination) -or (-not $ConfirmValidationOfMultiInstance)) {
                if ($ReturnInlineErrorMessage) {
                    "Kindly provide values for parameter SqlInstanceAsDataDestination & ConfirmValidationOfMultiInstance as `$ssn4PerfmonSetup is null." | Write-Error
                }
                else {            
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly provide values for parameter SqlInstanceAsDataDestination & ConfirmValidationOfMultiInstance as `$ssn4PerfmonSetup is null." | Write-Host -ForegroundColor Red
                    "STOP here, and fix above issue." | Write-Error
                }
            }
        }
        
        # if pssession is not null
        if(-not [String]::IsNullOrEmpty($ssn4PerfmonSetup)) {
        $sqlServicesOnHost += Invoke-Command -Session $ssn4PerfmonSetup -ScriptBlock { 
                                    Get-Service MSSQL* | Where-Object {$_.DisplayName -like 'SQL Server (*)' -and $_.StartType -ne 'Disabled'} 
                            }
        }
    }

    # If more than one sql services found, then ensure appropriate parameters are provided
    if($sqlServicesOnHost.Count -gt 1) 
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "[$($sqlServicesOnHost.Count)] database engine Services found on [$HostName].."

        # If Destination instance is not provided, throw error
        if([String]::IsNullOrEmpty($SqlInstanceAsDataDestination) -or (-not $ConfirmValidationOfMultiInstance)) 
        {
            if([String]::IsNullOrEmpty($SqlInstanceAsDataDestination)) {
                $errMessage = "Kindly provide value for parameter SqlInstanceAsDataDestination as host has multiple database engine services, `n`t and Perfmon data can be saved on only on one SQLInstance."
                #"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', $errMessage | Write-Host -ForegroundColor Red
            }
            if(-not $ConfirmValidationOfMultiInstance) {
                $errMessage = "Kindly set ConfirmValidationOfMultiInstance parameter to true as host has multiple database engine services, `n`t and Perfmon data can be saved on only on one SQLInstance."
                #"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly set ConfirmValidationOfMultiInstance parameter to true as host has multiple database engine services, `n`t and Perfmon data can be saved on only on one SQLInstance." | Write-Host -ForegroundColor Red
            }

            if ($ReturnInlineErrorMessage) {
                $errMessage | Write-Error
            }
            else {            
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', $errMessage | Write-Host -ForegroundColor Red
        
                "STOP here, and fix above issue." | Write-Error
            }
        }
    }
}

# Set $SqlInstanceAsDataDestination same as $SqlInstanceToBaseline if NULL
Write-Verbose 'Set `$SqlInstanceAsDataDestination same as `$SqlInstanceToBaseline if NULL'
$Port4SqlInstanceAsDataDestination = $null
if([String]::IsNullOrEmpty($SqlInstanceAsDataDestination)) {
    $SqlInstanceAsDataDestination = $SqlInstanceToBaseline
    $SqlInstanceAsDataDestinationWithOutPort = $SqlInstanceToBaselineWithOutPort
} else {
    $SqlInstanceAsDataDestinationWithOutPort = $SqlInstanceAsDataDestination

    # Check if PortNo is specified
    if($SqlInstanceAsDataDestination -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
        $Port4SqlInstanceAsDataDestination = $Matches['PortNo']
        $SqlInstanceAsDataDestinationWithOutPort = $Matches['SqlInstance']
    }
}

# Set $SqlInstanceForTsqlJobs same as $SqlInstanceToBaseline if NULL
Write-Verbose 'Set `$SqlInstanceForTsqlJobs same as `$SqlInstanceToBaseline if NULL'
$Port4SqlInstanceForTsqlJobs = $null
if([String]::IsNullOrEmpty($SqlInstanceForTsqlJobs)) {
    $SqlInstanceForTsqlJobs = $SqlInstanceToBaseline
    $SqlInstanceForTsqlJobsWithOutPort = $SqlInstanceToBaselineWithOutPort
} else {
    $SqlInstanceForTsqlJobsWithOutPort = $SqlInstanceForTsqlJobs

    # Check if PortNo is specified
    if($SqlInstanceForTsqlJobs -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
        $Port4SqlInstanceForTsqlJobs = $Matches['PortNo']
        $SqlInstanceForTsqlJobsWithOutPort = $Matches['SqlInstance']
    }
}

# Set $SqlInstanceForPowershellJobs same as $SqlInstanceToBaseline if NULL
Write-Verbose 'Set `$SqlInstanceForPowershellJobs same as `$SqlInstanceToBaseline if NULL'
$Port4SqlInstanceForPowershellJobs
if([String]::IsNullOrEmpty($SqlInstanceForPowershellJobs)) {
    $SqlInstanceForPowershellJobs = $SqlInstanceToBaseline
    $SqlInstanceForPowershellJobsWithOutPort = $SqlInstanceToBaselineWithOutPort
} else {
    $SqlInstanceForPowershellJobsWithOutPort = $SqlInstanceForPowershellJobs

    # Check if PortNo is specified
    if($SqlInstanceForPowershellJobs -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
        $Port4SqlInstanceForPowershellJobs = $Matches['PortNo']
        $SqlInstanceForPowershellJobsWithOutPort = $Matches['SqlInstance']
    }
}


# Check if only port no is wrong/missing
if($SqlInstanceAsDataDestinationWithOutPort -eq $SqlInstanceToBaselineWithOutPort) 
{
    if($SqlInstanceAsDataDestination -ne $SqlInstanceToBaseline) 
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Port is wrong for DataDestination. Updating same.."
        $SqlInstanceAsDataDestination = $SqlInstanceToBaseline
        $Port4SqlInstanceAsDataDestination = $Port4SqlInstanceToBaseline

        $sqlUpdateDataDestination = @"
update instance_details 
set data_destination_sql_instance = '$SqlInstanceAsDataDestination'
where sql_instance = '$SqlInstanceToBaselineWithOutPort'
and host_name = '$HostName'
"@
        $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlUpdateDataDestination -EnableException
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlUpdateDataDestination -EnableException
    }
}

if($SqlInstanceForTsqlJobsWithOutPort -eq $SqlInstanceToBaselineWithOutPort) 
{
    if($SqlInstanceForTsqlJobs -ne $SqlInstanceToBaseline) 
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Port is wrong for TsqlJobs server. Updating same.."
        $SqlInstanceForTsqlJobs = $SqlInstanceToBaseline
        $Port4SqlInstanceForTsqlJobs = $Port4SqlInstanceToBaseline

        $sqlUpdateTsqlJobsServer = @"
update dbo.instance_details 
set collector_tsql_jobs_server = '$SqlInstanceForTsqlJobs'
where sql_instance = '$SqlInstanceToBaselineWithOutPort'
and host_name = '$HostName'
"@
        $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlUpdateTsqlJobsServer -EnableException
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlUpdateTsqlJobsServer -EnableException
    }
}

if($SqlInstanceForPowershellJobsWithOutPort -eq $SqlInstanceToBaselineWithOutPort) 
{
    if($SqlInstanceForPowershellJobs -ne $SqlInstanceToBaseline) 
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Port is wrong for PowerShell server. Updating same.."
        $SqlInstanceForPowershellJobs = $SqlInstanceToBaseline
        $Port4SqlInstanceForPowershellJobs = $Port4SqlInstanceToBaseline

        $sqlUpdatePowerShellJobsServer = @"
update dbo.instance_details 
set collector_powershell_jobs_server = '$SqlInstanceForPowershellJobs'
where sql_instance = '$SqlInstanceToBaselineWithOutPort'
and host_name = '$HostName'
"@
        $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlUpdatePowerShellJobsServer -EnableException
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlUpdatePowerShellJobsServer -EnableException
    }
}


"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$DbaDatabase = [$DbaDatabase]" | Write-Host -ForegroundColor Yellow
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$isUpgradeScenario = [$isUpgradeScenario]" | Write-Host -ForegroundColor Yellow
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$InstanceScopeFeaturesOnly = [$InstanceScopeFeaturesOnly]" | Write-Host -ForegroundColor Yellow
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$SqlInstanceAsDataDestination = [$SqlInstanceAsDataDestination]"
if( (-not [String]::IsNullOrEmpty($Port4SqlInstanceAsDataDestination)) -and ($SqlInstanceToBaseline -ne $SqlInstanceAsDataDestination) ) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "SQL Port for [$SqlInstanceAsDataDestination] => $Port4SqlInstanceAsDataDestination." | Write-Host -ForegroundColor Yellow
}
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$SqlInstanceForTsqlJobs = [$SqlInstanceForTsqlJobs]"
if( (-not [String]::IsNullOrEmpty($Port4SqlInstanceForTsqlJobs)) -and ($SqlInstanceToBaseline -ne $SqlInstanceForTsqlJobs) ) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "SQL Port for [$SqlInstanceForTsqlJobs] => $Port4SqlInstanceForTsqlJobs." | Write-Host -ForegroundColor Yellow
}
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$SqlInstanceForPowershellJobs = [$SqlInstanceForPowershellJobs]"
if( (-not [String]::IsNullOrEmpty($Port4SqlInstanceForPowershellJobs)) -and ($SqlInstanceToBaseline -ne $SqlInstanceForPowershellJobs) ) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "SQL Port for [$SqlInstanceForPowershellJobs] => $Port4SqlInstanceForPowershellJobs." | Write-Host -ForegroundColor Yellow
}

# Create SQL Connections
if([String]::IsNullOrEmpty($SqlInstanceAsDataDestination) -or ($SqlInstanceAsDataDestinationWithOutPort -eq $SqlInstanceToBaselineWithOutPort)) {
    $conSqlInstanceAsDataDestination = $conSqlInstanceToBaseline
} else {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "[Connect-DbaInstance] Create connection for `$SqlInstanceAsDataDestination = '$SqlInstanceAsDataDestination'.."
    $conSqlInstanceAsDataDestination = Connect-DbaInstance -SqlInstance $SqlInstanceAsDataDestination -Database master -ClientName $clientName `
                                                -SqlCredential $SqlCredential -TrustServerCertificate -EncryptConnection -Verbose:$false -Debug:$false
}

if([String]::IsNullOrEmpty($SqlInstanceForTsqlJobs) -or ($SqlInstanceForTsqlJobsWithOutPort -eq $SqlInstanceToBaselineWithOutPort)) {
    $conSqlInstanceForTsqlJobs = $conSqlInstanceToBaseline
} else {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "[Connect-DbaInstance] Create connection for `$SqlInstanceForTsqlJobs = '$SqlInstanceForTsqlJobs'.."
    $conSqlInstanceForTsqlJobs = Connect-DbaInstance -SqlInstance $SqlInstanceForTsqlJobs -Database master -ClientName $clientName `
                                                -SqlCredential $SqlCredential -TrustServerCertificate -EncryptConnection -Verbose:$false -Debug:$false
}

if([String]::IsNullOrEmpty($SqlInstanceForPowershellJobs) -or ($SqlInstanceForPowershellJobsWithOutPort -eq $SqlInstanceToBaselineWithOutPort)) {
    $conSqlInstanceForPowershellJobs = $conSqlInstanceToBaseline
} else {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "[Connect-DbaInstance] Create connection for `$SqlInstanceForPowershellJobs = '$SqlInstanceForPowershellJobs'.."
    $conSqlInstanceForPowershellJobs = Connect-DbaInstance -SqlInstance $SqlInstanceForPowershellJobs -Database master -ClientName $clientName `
                                                -SqlCredential $SqlCredential -TrustServerCertificate -EncryptConnection -Verbose:$false -Debug:$false
}


# If destination is provided, then validate if perfmon is not already getting collected
if ( ($sqlServicesOnHost.Count -gt 1) -and (-not [String]::IsNullOrEmpty($SqlInstanceAsDataDestination)) -and ($ConfirmValidationOfMultiInstance) ) 
{
    Write-Verbose 'If destination is provided, then validate if perfmon is not already getting collected'
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Validate if Perfmon data is not being collected already on [$SqlInstanceAsDataDestination] for same host.."
    $sqlPerfmonRecord = @"
if OBJECT_ID('dbo.performance_counters') is not null
begin
	select top 1 'dbo.performance_counters' as QueryData, getutcdate() as current_time_utc, collection_time_utc, pc.host_name
	from dbo.performance_counters pc with (nolock)
	where pc.collection_time_utc >= DATEADD(minute,-20,GETUTCDATE()) and host_name = '$HostName'
	order by pc.collection_time_utc desc
end
"@
    $resultPerfmonRecord = @()
    try {
        $resultPerfmonRecord += $conSqlInstanceAsDataDestination | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlPerfmonRecord -EnableException -Verbose:$false -Debug:$false
    }
    catch {
        $errMessage = $_.Exception.Message
        if ($ReturnInlineErrorMessage) {
            $errMessage | Write-Error
        }
        else {            
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "$errMessage" | Write-Host -ForegroundColor Red
            "STOP here, and fix above issue." | Write-Error
        }
    }

    if($resultPerfmonRecord.Count -eq 0) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "No Perfmon data record found for last 20 minutes for host [$HostName] on [$SqlInstanceAsDataDestination]."
    }
    else {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Perfmon data records of latest 20 minutes for host [$HostName] are present on [$SqlInstanceAsDataDestination]."
    }
}


# If Express edition, then ensure another server is mentioned for Creating jobs
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Checking if [$SqlInstanceToBaseline] is Express Edition.."
$isExpressEdition = $false
if($dbServiceInfo.Edition -like 'Express*') {
    $isExpressEdition = $true
    if($ConfirmSetupOfTaskSchedulerJobs -eq $false) 
    {
        if( ($SqlInstanceForTsqlJobs -eq $SqlInstanceToBaseline) -or ($SqlInstanceForPowershellJobs -eq $SqlInstanceToBaseline) ) 
        {
            if ($ReturnInlineErrorMessage) 
            {
		        @"
Curent instance is Express edition.
Option 01: Kindly provide a different SQLInstance for parameters SqlInstanceForTsqlJobs & SqlInstanceForPowershellJobs.
Option 02: Using parameter ConfirmSetupOfTaskSchedulerJobs, kindly confirm for windows Task Scheduler based jobs.
"@ | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Curent instance is Express edition." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Option 01: Kindly provide a different SQLInstance for parameters SqlInstanceForTsqlJobs & SqlInstanceForPowershellJobs." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Option 02: Using parameter ConfirmSetupOfTaskSchedulerJobs, kindly confirm for windows Task Scheduler based jobs." | Write-Host -ForegroundColor Red
                "STOP and check above error message" | Write-Error
            }
        }
    }
    else {
        if([String]::IsNullOrEmpty($WindowsCredential)) {
            if ($ReturnInlineErrorMessage) {
		        "Curent instance is Express edition.`nSo, parameter WindowsCredential is must." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Curent instance is Express edition." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "So, parameter WindowsCredential is must." | Write-Host -ForegroundColor Red
                "STOP and check above error message" | Write-Error
            }
        }
    }
}


# Validate database collation
if(-not $SkipCollationCheck) 
{
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Validating Collation of databases.."
    $sqlDbCollation = @"
select name as [db_name], collation_name from sys.databases 
where collation_name not in ('SQL_Latin1_General_CP1_CI_AS') 
and name in ('master','msdb','tempdb','$DbaDatabase')
"@
    $dbCollationResult = @()
    $dbCollationResult += $conSqlInstanceToBaseline | Invoke-DbaQuery -Query $sqlDbCollation -EnableException -Verbose:$false -Debug:$false
    if($dbCollationResult.Count -ne 0) {
        if ($ReturnInlineErrorMessage) 
        {
		    "Collation of below databases is not [SQL_Latin1_General_CP1_CI_AS].`nKindly rectify this collation problem, or Using SkipCollationCheck parameter.`n`n$($dbCollationResult | Format-Table -AutoSize)" | Write-Error
	    }
	    else {            
		    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Collation of below databases is not [SQL_Latin1_General_CP1_CI_AS]." | Write-Host -ForegroundColor Red
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly rectify this collation problem, or Using SkipCollationCheck parameter." | Write-Host -ForegroundColor Red
            $dbCollationResult | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Red
            Start-Sleep -Seconds 1
            Write-Error "Stop here. Fix above issue."
        }
    }
}


# Get HostName for $SqlInstanceForPowershellJobs
if($SqlInstanceToBaseline -ne $SqlInstanceForPowershellJobs) 
{
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Fetching basic info for `$SqlInstanceForPowershellJobs => [$SqlInstanceForPowershellJobs].."
    try {
        $jobServerServicesInfo = $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Query $sqlServerInfo -EnableException -Verbose:$false -Debug:$false
        $jobServerDbServiceInfo = $jobServerServicesInfo | Where-Object {$_.service_name_str -like "SQL Server (*)"}
        $jobServerAgentServiceInfo = $jobServerServicesInfo | Where-Object {$_.service_name_str -like "SQL Server Agent (*)"}
        $jobServerServicesInfo | Format-Table -AutoSize
    }
    catch {
        $errMessage = $_
    
        if ($ReturnInlineErrorMessage) 
        {
            if([String]::IsNullOrEmpty($SqlCredential)) {
                $errMessage = "SQL Connection to [$SqlInstanceToBaseline] failed.`nKindly provide SqlCredentials.`n$($errMessage.Exception.Message).."
            } else {
                $errMessage = "SQL Connection to [$SqlInstanceToBaseline] failed.`nProvided SqlCredentials seems to be NOT working.`n$($errMessage.Exception.Message).."
            }

            $errMessage | Write-Error
        }
        else
        {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "SQL Connection to [$SqlInstanceToBaseline] failed."
            if([String]::IsNullOrEmpty($SqlCredential)) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly provide SqlCredentials." | Write-Host -ForegroundColor Red
            } else {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Provided SqlCredentials seems to be NOT working." | Write-Host -ForegroundColor Red
            }
            Write-Error "Stop here. Fix above issue."
        }
    }
}
else {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$SqlInstanceToBaseline ~ `$SqlInstanceForPowershellJobs.."
    $jobServerServicesInfo = $resultServerInfo
    $jobServerDbServiceInfo = $dbServiceInfo
    $jobServerAgentServiceInfo = $agentServiceInfo
}

# Setup PSSession on $SqlInstanceForPowershellJobs
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Validating if PSSession is needed on `$SqlInstanceForPowershellJobs.."
if( (-not $SkipRDPSessionSteps) -and ($HostName -ne $jobServerDbServiceInfo.host_name) )
{
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Create PSSession for host [$($jobServerDbServiceInfo.host_name)].."
    $ssnHostName = $jobServerDbServiceInfo.host_name #+'.'+$jobServerDbServiceInfo.domain_reg

    # Try reaching server using HostName provided/detected, if fails, then use FQDN
    if (-not (Test-Connection -ComputerName $ssnHostName -Quiet -Count 1)) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Host [$ssnHostName] not pingable. So trying FQDN form.."
        $ssnHostName = $ssnHostName+'.'+$jobServerDbServiceInfo.domain_reg
    }

    # Try reaching using FQDN, if fails & not a clustered instance, then use SqlInstanceToBaseline itself
    if (-not (Test-Connection -ComputerName $ssnHostName -Quiet -Count 1)) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Host [$ssnHostName] not pingable. So trying `$SqlInstanceForPowershellJobs parameter value itself.."
        $ssnHostName = $SqlInstanceForPowershellJobs
    }

    # Try reaching using FQDN, if fails & not a clustered instance, then use SqlInstanceToBaseline itself
    if ( -not (Test-Connection -ComputerName $ssnHostName -Quiet -Count 1) ) {
        if ($ReturnInlineErrorMessage) {
		    "Host [$ssnHostName] not pingable.`nKindly ensure pssession is working for `$SqlInstanceForPowershellJobs [$SqlInstanceForPowershellJobs]." | Write-Error
	    }
	    else {            
		    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Host [$ssnHostName] not pingable." | Write-Host -ForegroundColor Red
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly ensure pssession is working for `$SqlInstanceForPowershellJobs [$SqlInstanceForPowershellJobs]." | Write-Host -ForegroundColor Red
            "STOP and check above error message" | Write-Error
        }
    }

    $ssnJobServer = $null
    $errVariables = @()

    # First Attempt without Any credentials
    try {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Trying for PSSession on [$ssnHostName] normally.."
            $ssnJobServer = New-PSSession -ComputerName $ssnHostName 
        }
    catch { $errVariables += $_ }

    # Second Attempt for Trusted Cross Domains
    if( [String]::IsNullOrEmpty($ssnJobServer) ) {
        try { 
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Trying for PSSession on [$ssnHostName] assuming cross domain.."
            $ssnJobServer = New-PSSession -ComputerName $ssnHostName -Authentication Negotiate 
        }
        catch { $errVariables += $_ }
    }

    # 3rd Attempt with Credentials
    if( [String]::IsNullOrEmpty($ssnJobServer) -and (-not [String]::IsNullOrEmpty($WindowsCredential)) ) {
        try {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Trying PSSession for [$ssnHostName] using provided WindowsCredentials.."
            $ssnJobServer = New-PSSession -ComputerName $ssnHostName -Credential $WindowsCredential    
        }
        catch { $errVariables += $_ }

        if( [String]::IsNullOrEmpty($ssn) ) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Trying PSSession for [$ssnHostName] using provided WindowsCredentials with Negotiate attribute.."
            $ssnJobServer = New-PSSession -ComputerName $ssnHostName -Credential $WindowsCredential -Authentication Negotiate
        }
    }

    if ( [String]::IsNullOrEmpty($ssnJobServer) ) {
        if ($ReturnInlineErrorMessage) {
            "Provide WindowsCredential for accessing server [$ssnHostName] of domain '$($sqlServerInfo.domain)'." | Write-Error
        }
        else {            
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Provide WindowsCredential for accessing server [$ssnHostName] of domain '$($sqlServerInfo.domain)'." | Write-Host -ForegroundColor Red
            "STOP here, and fix above issue." | Write-Error
        }
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "PSSession for [$($jobServerDbServiceInfo.host_name)].."
    $ssnJobServer
}
else {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$ssnJobServer is same as `$ssn4PerfmonSetup."
    $ssnJobServer = $ssn4PerfmonSetup
}


# Service Account and Access Validation
$requireProxy = $false
if( ($SkipPowerShellJobs -or $SkipAllJobs) -and ($SkipWindowsAdminAccessTest -eq $false) -and ('20__CreateJobRemoveXEventFiles' -notin $Steps2Execute) ) { $SkipWindowsAdminAccessTest = $true }
if($SkipWindowsAdminAccessTest -eq $false)
{
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Validate for WindowsCredential if SQL Service Accounts are non-priviledged.."

    # If Express edition, and Task scheduler jobs are required
    if( (-not [String]::IsNullOrEmpty($WindowsCredential)) -and ($isExpressEdition) -and ($ConfirmSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$WindowsCredential = '$($WindowsCredential.UserName)'"
        $adminAccessTestScript = {
                $AdminName = Get-LocalUser | Where-Object {$_.Name -like 'Admin*' -and $_.Enabled} | Select-Object -ExpandProperty Name -First 1
                Get-ChildItem "C:\Users\$AdminName" -ErrorAction Stop | Out-Null
            }
        Invoke-Command -Session $ssn4PerfmonSetup -ScriptBlock $adminAccessTestScript -ErrorAction Stop
    }
    else # If not express edition
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$TestWindowsAdminAccessJobFilePath = '$TestWindowsAdminAccessJobFilePath'"
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating & executing job [(dba) Test-WindowsAdminAccess] on [$SqlInstanceForPowershellJobs].."
        $sqlTestWindowsAdminAccessFilePath = [System.IO.File]::ReadAllText($TestWindowsAdminAccessJobFilePath)
        $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query $sqlTestWindowsAdminAccessFilePath -EnableException -Verbose:$false -Debug:$false

        $testWindowsAdminAccessJobHistory = @()
        $loopStartTime = Get-Date
        $sleepDurationSeconds = 5
        $loopTotalDurationThresholdSeconds = 300    
    
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Fetching execution history for job [(dba) Test-WindowsAdminAccess] on [$SqlInstanceForPowershellJobs].."
        while ($testWindowsAdminAccessJobHistory.Count -eq 0 -and $( (New-TimeSpan $loopStartTime $(Get-Date)).TotalSeconds -le $loopTotalDurationThresholdSeconds ) )
        {
            $testWindowsAdminAccessJobHistory += Get-DbaAgentJobHistory -SqlInstance $conSqlInstanceForPowershellJobs -Job '(dba) Test-WindowsAdminAccess' `
                                                        -ExcludeJobSteps -EnableException

            if($testWindowsAdminAccessJobHistory.Count -eq 0) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Wait for $sleepDurationSeconds seconds as the job might be running.."
                Start-Sleep -Seconds $sleepDurationSeconds
            }
        }

        if($testWindowsAdminAccessJobHistory.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Status of job [(dba) Test-WindowsAdminAccess] on [$SqlInstanceForPowershellJobs] could not be fetched on time. Kindly validate." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Status of job [(dba) Test-WindowsAdminAccess] on [$SqlInstanceForPowershellJobs] could not be fetched on time. Kindly validate." | Write-Host -ForegroundColor Red
                "STOP and check above error message" | Write-Error
            }
        }
        else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "[(dba) Test-WindowsAdminAccess] Job history => '$($testWindowsAdminAccessJobHistory.Message)'."
            $testWindowsAdminAccessJobHistory | Format-Table -AutoSize
        }

        $hasWindowsAdminAccess = $false
        $sqlServerAgentInfo = if($SqlInstanceForPowershellJobs -ne $SqlInstanceToBaseline) {$jobServerAgentServiceInfo} else {$agentServiceInfo}

        if($testWindowsAdminAccessJobHistory.Status -ne 'Succeeded') {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "SQL Agent service account [$($sqlServerAgentInfo.service_account)] DO NOT have admin access at windows."
        } else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "SQL Agent service account [$($sqlServerAgentInfo.service_account)] has admin access at windows."
            $hasWindowsAdminAccess = $true
        }

        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Remove test job [(dba) Test-WindowsAdminAccess].."
        $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query "EXEC msdb.dbo.sp_delete_job @job_name=N'(dba) Test-WindowsAdminAccess'" -EnableException -Verbose:$false -Debug:$false


        $requireProxy = $(-not $hasWindowsAdminAccess)
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$hasWindowsAdminAccess = $hasWindowsAdminAccess"
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$requireProxy = $requireProxy"

        if($requireProxy -and [String]::IsNullOrEmpty($WindowsCredential)) {
            if ($ReturnInlineErrorMessage) {
		        "Kindly provide WindowsCredential to create SQL Agent Job Proxy." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly provide WindowsCredential to create SQL Agent Job Proxy." | Write-Host -ForegroundColor Red
                "STOP and check above error message" | Write-Error
            }
        }
    }
}
else {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Since SkipWindowsAdminAccessTest is set to TRUE, assuming `$requireProxy to $requireProxy."
}


if (-not $EnableEmailAlerts) {
    $SkipMailProfileCheck = $true
}

# Validate mail profile
if(-not $SkipMailProfileCheck)
{
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Checking for default global mail profile.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$MailProfileFilePath = '$MailProfileFilePath'"
    $sqlMailProfile = @"
SELECT p.name as profile_name, p.description as profile_description, a.name as mail_account, 
		a.email_address, a.display_name, a.replyto_address, s.servername, s.port, s.servername,
		pp.is_default
FROM msdb.dbo.sysmail_profile p 
JOIN msdb.dbo.sysmail_principalprofile pp ON pp.profile_id = p.profile_id AND pp.is_default = 1
JOIN msdb.dbo.sysmail_profileaccount pa ON p.profile_id = pa.profile_id 
JOIN msdb.dbo.sysmail_account a ON pa.account_id = a.account_id 
JOIN msdb.dbo.sysmail_server s ON a.account_id = s.account_id
WHERE pp.is_default = 1
"@
    $mailProfile = @()
    $mailProfile += $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlMailProfile -EnableException -Verbose:$false -Debug:$false
    if($mailProfile.Count -lt 1) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Kindly create default global mail profile." | Write-Host -ForegroundColor Red
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Kindly utilize '$mailProfileFilePath." | Write-Host -ForegroundColor Red
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Opening the file '$mailProfileFilePath' in notepad.." | Write-Host -ForegroundColor Red
        notepad "$mailProfileFilePath"

        $mailProfile += Get-DbaDbMailProfile -SqlInstance $conSqlInstanceToBaseline
        if($mailProfile.Count -ne 0) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Below mail profile(s) exists.`nOne of them can be set to default global profile." | Write-Host -ForegroundColor Red
            $mailProfile | Format-Table -AutoSize
        }

        if ($ReturnInlineErrorMessage) {
		    "Kindly create default global mail profile.." | Write-Error
	    }
	    else {            
		    Write-Error "Stop here. Fix above issue."
        }
    }
}


# Fetch DBA Database File Path
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Fetch [$DbaDatabase] path.."
if($SkipDriveCheck) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Skip validation of C:\ drive check.." | Write-Host -ForegroundColor Cyan
}
$sqlDbaDatabasePath = @"
select top 1 physical_name FROM sys.master_files 
where database_id = DB_ID('$DbaDatabase') and type_desc = 'ROWS' 
$(if($SkipDriveCheck){'--'})and physical_name not like 'C:\%' 
order by file_id;
"@
$resultDbaDatabasePath = @()
$resultDbaDatabasePath += $conSqlInstanceToBaseline | Invoke-DbaQuery -Database master -Query $sqlDbaDatabasePath -EnableException -Verbose:$false -Debug:$false
if($resultDbaDatabasePath.Count -eq 0) {
    if ($ReturnInlineErrorMessage) {
		"Seems either [$DbaDatabase] does not exists, or the data/log files are present in C:\ drive. `n`t Kindly rectify this issue." | Write-Error
	}
	else {            
		"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Seems either [$DbaDatabase] does not exists, or the data/log files are present in C:\ drive. `n`t Kindly rectify this issue." | Write-Host -ForegroundColor Red
        Write-Error "Stop here. Fix above issue."
    }
}
else {
    $dbaDatabasePath = $resultDbaDatabasePath[0].physical_name
}
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$dbaDatabasePath => '$dbaDatabasePath'.."


# Execute PreQuery
if(-not [String]::IsNullOrEmpty($PreQuery)) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Executing PreQuery on [$SqlInstanceToBaseline].." | Write-Host -ForegroundColor Cyan
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $PreQuery -EnableException -Verbose:$false -Debug:$false
}


# If Express edition, and Task scheduler jobs are required
if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
    -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
{
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Ensuring Logs folder on remote server path '$logsPath'.."

    $createFolderScript = {
            Param ($logsPath)
            New-Item $logsPath -ItemType Directory -Force | Out-Null
        }
    Invoke-Command -Session $ssn4PerfmonSetup -ScriptBlock $createFolderScript -ArgumentList @($logsPath) -ErrorAction Stop
}


# Get dbo.instance_details info by HostName
    # This helps to figure out if OS Jobs already exists for same host with other sql instance
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Fetch entries for same host from [$InventoryServer].[$InventoryDatabase].[dbo].[instance_details].."
$instanceHostDetails = @()
if(-not [String]::IsNullOrEmpty($HostName)) {
    $sqlInstanceHostDetails = @"
    if OBJECT_ID('dbo.instance_details') is not null
    begin
	    select	sql_instance_with_port = coalesce(sql_instance+','+sql_instance_port,sql_instance), 
		        collector_powershell_jobs_server, host_name, data_destination_sql_instance
        from dbo.instance_details 
        where is_enabled = 1 and is_alias = 0 
        and [host_name] = '$HostName'
    end
    else
	    print 'Seems fresh install as dbo.instance_details does not exist'
"@
}

if($verbose) {
    $sqlInstanceHostDetails | Write-Host -ForegroundColor Cyan
}

try {
    $instanceHostDetails += $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlInstanceHostDetails -Verbose:$false -Debug:$false
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Host related info fetched."

    if($verbose) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$instanceHostDetails =>`n"
        $instanceHostDetails | Format-Table -AutoSize
    }

    if($instanceHostDetails.Count -gt 0)
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Same host entries already present."    

        $hostEntryForCurrentCombination = @()
        $hostEntryForCurrentCombination += $instanceHostDetails | `
                Where-Object {$_.sql_instance_with_port -eq $SqlInstanceToBaseline -and $_.data_destination_sql_instance -eq $SqlInstanceAsDataDestination}
        
        $hostEntryOwnedBySqlInstance = @()
        $hostEntryOwnedBySqlInstance += $instanceHostDetails | `
                Where-Object {$_.sql_instance_with_port -eq $SqlInstanceToBaseline -and $_.sql_instance_with_port -eq $_.data_destination_sql_instance}

        $hostEntryOwnedByOtherServer = @()
        $hostEntryOwnedByOtherServer += $instanceHostDetails | `
                Where-Object {$_.sql_instance_with_port -ne $SqlInstanceToBaseline -and $_.sql_instance_with_port -eq $_.data_destination_sql_instance} | `
                Select-Object data_destination_sql_instance -Unique

        if($verbose) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$hostEntryForCurrentCombination =>`n"
            $hostEntryForCurrentCombination | Format-Table -AutoSize

            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$hostEntryOwnedBySqlInstance =>`n"
            $hostEntryOwnedBySqlInstance | Format-Table -AutoSize

            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$hostEntryOwnedByOtherServer =>`n"
            $hostEntryOwnedByOtherServer | Format-Table -AutoSize
        }

        if($hostEntryOwnedByOtherServer.Count -gt 0) {
            $SkipPowerShellJobs4Host = $true
            $hostOwnerServer = $hostEntryOwnedByOtherServer | Select-Object -ExpandProperty data_destination_sql_instance -First 1
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "OS Jobs for host [$HostName] are owned by [$hostOwnerServer] Server. So skipping `$PowerShellJobSteps." | Write-Host -ForegroundColor Yellow            
        }
    }

    #if($instanceHostDetails.Count -gt 0 -and $isClustered -eq $true) {
        #$SkipPowerShellJobs4SQLCluster = $true
    #}
}
catch {
    $errMessage = $_

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Could not fetch host related details from [$InventoryServer].[$InventoryDatabase].[dbo].[instance_details] info."
}


# Check if PortNo is specified for TSQL Jobs Server
$Port4SqlInstanceForTsqlJobs = $null
$SqlInstanceForTsqlJobsWithOutPort = $SqlInstanceForTsqlJobs
if($SqlInstanceForTsqlJobs -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
    $Port4SqlInstanceForTsqlJobs = $Matches['PortNo']
    $SqlInstanceForTsqlJobsWithOutPort = $Matches['SqlInstance']
}

# Check if PortNo is specified for PowerShell Jobs Server
$Port4SqlInstanceForPowershellJobs = $null
$SqlInstanceForPowershellJobsWithOutPort = $SqlInstanceForPowershellJobs
if($SqlInstanceForPowershellJobs -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
    $Port4SqlInstanceForPowershellJobs = $Matches['PortNo']
    $SqlInstanceForPowershellJobsWithOutPort = $Matches['SqlInstance']
}

# Check if PortNo is specified for Data Destination Server
$Port4SqlInstanceAsDataDestination = $null
$SqlInstanceAsDataDestinationWithOutPort = $SqlInstanceAsDataDestination
if($SqlInstanceAsDataDestination -match "(?'SqlInstance'.+),(?'PortNo'\d+)") {
    $Port4SqlInstanceAsDataDestination = $Matches['PortNo']
    $SqlInstanceAsDataDestinationWithOutPort = $Matches['SqlInstance']
}

if($SqlInstanceToBaselineWithOutPort -eq $SqlInstanceAsDataDestinationWithOutPort) {
    if($SqlInstanceToBaseline -ne $SqlInstanceAsDataDestination) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "SqlInstanceAsDataDestination is not matching in dbo.instance_details." | Write-Host -ForegroundColor Red
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "`$SqlInstanceToBaseline ($SqlInstanceToBaseline) <> `$SqlInstanceAsDataDestination ($SqlInstanceAsDataDestination)." | Write-Host -ForegroundColor Red
        Write-Error "Stop here. Fix above issue."
    }
}

if($SqlInstanceToBaselineWithOutPort -eq $SqlInstanceForTsqlJobsWithOutPort) {
    if($SqlInstanceToBaseline -ne $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "SqlInstanceForTsqlJobs is not matching in dbo.instance_details." | Write-Host -ForegroundColor Red
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "`$SqlInstanceToBaseline ($SqlInstanceToBaseline) <> `$SqlInstanceForTsqlJobs ($SqlInstanceForTsqlJobs)." | Write-Host -ForegroundColor Red
        Write-Error "Stop here. Fix above issue."
    }
}

if($SqlInstanceToBaselineWithOutPort -eq $SqlInstanceForPowershellJobsWithOutPort) {
    if($SqlInstanceToBaseline -ne $SqlInstanceForPowershellJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "SqlInstanceForPowershellJobs is not matching in dbo.instance_details." | Write-Host -ForegroundColor Red
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "`$SqlInstanceToBaseline ($SqlInstanceToBaseline) <> `$SqlInstanceForPowershellJobs ($SqlInstanceForPowershellJobs)." | Write-Host -ForegroundColor Red
        Write-Error "Stop here. Fix above issue."
    }
}

# When baselining Managed Instance (PAAS)
if($InstanceScopeFeaturesOnly) 
{
    $ForceTSQLStepType4TsqlJobs = $true
    $SkipPowerShellJobs = $true
    $SkipRDPSessionSteps = $true
    $SkipWindowsAdminAccessTest = $true

    $CollectDiskSpaceJobFilePath = $CollectDiskSpaceManagedInstanceJobFilePath
    $CollectPerfmonDataJobFilePath = $CollectPerfmonDataManagedInstanceJobFilePath

    $PowerShellJobSteps = $($PowerShellJobSteps | % {if($_ -notin @('13__CreateJobCollectDiskSpace','14__CreateJobCollectPerfmonData')){$_}});
    $TsqlJobSteps = $TsqlJobSteps + $( @('13__CreateJobCollectDiskSpace','14__CreateJobCollectPerfmonData') | % {if($_ -notin $TsqlJobSteps){$_}} );

    # re-define Steps2Execute
    $Steps2Execute = $Steps2Execute | ForEach-Object {if($_ -notin $PowerShellJobSteps){$_}}
    $Steps2Execute = $Steps2Execute | ForEach-Object {if($_ -notin $RDPSessionSteps){$_}}
}


# 1__sp_WhoIsActive
$stepName = '1__sp_WhoIsActive'
if($stepName -in $Steps2Execute) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$WhoIsActiveFilePath = '$WhoIsActiveFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating sp_WhoIsActive in [master] database.."
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database master -File $(Get-Item -LiteralPath $WhoIsActiveFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Checking if sp_WhoIsActive is present in [$DbaDatabase] also.."
    $sqlCheckWhoIsActiveExistence = "select [is_existing] = case when OBJECT_ID('dbo.sp_WhoIsActive') is null then 0 else 1 end;"
    $existsWhoIsActive = $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlCheckWhoIsActiveExistence -EnableException -Verbose:$false -Debug:$false | Select-Object -ExpandProperty is_existing;
    if($existsWhoIsActive -eq 1) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update sp_WhoIsActive definition in [$DbaDatabase] also.."
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $WhoIsActiveFilePath) -EnableException -Verbose:$false -Debug:$false
    }
}


# 2__AllDatabaseObjects
$stepName = '2__AllDatabaseObjects'
if($stepName -in $Steps2Execute)
{
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating All Objects in [$DbaDatabase] database.."

    # Retrieve actual content & dump in temporary file
    $tempAllDatabaseObjectsFileName = "$($AllDatabaseObjectsFileName -replace '.sql','')-RuntimeUsedFile.sql"
    $tempAllDatabaseObjectsFilePath = Join-Path $ddlPath $tempAllDatabaseObjectsFileName

    $AllDatabaseObjectsFileContent = [System.IO.File]::ReadAllText($AllDatabaseObjectsFilePath)

    # MultiServerViews ~ [vw_performance_counters],[vw_disk_space]
    if($SkipMultiServerviewsUpgrade -eq $true) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "MultiServerViews are being skipped for upgrade."
        $AllDatabaseObjectsFileContent = $AllDatabaseObjectsFileContent.Replace("declare @recreate_multi_server_views bit = 1;", "declare @recreate_multi_server_views bit = 0;")
    }
    else {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "MultiServerViews are considered for upgrade."
    }

    # Modify content if SQL Server does not support Partitioning
    if($IsNonPartitioned) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Partitioning is not supported on current server."

        $AllDatabaseObjectsFileContent = $AllDatabaseObjectsFileContent.Replace('declare @is_partitioned bit = 1;', 'declare @is_partitioned bit = 0;')
        $AllDatabaseObjectsFileContent = $AllDatabaseObjectsFileContent.Replace(' on ps_dba', ' --on ps_dba')
    }

    # Modify content if upgrade scenario
    if($isUpgradeScenario) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "For upgrade scenario, don't touch partitioning."

        $AllDatabaseObjectsFileContent = $AllDatabaseObjectsFileContent.Replace('declare @is_upgrade bit = 0;', 'declare @is_upgrade bit = 1;')
    }

    # Modify AllDatabaseObjectsFileContent if MemoryOptimized Objects are NOT to be used
    if($MemoryOptimizedObjectsUsage -eq $false) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$MemoryOptimizedObjectsUsage is false. So use disk-based objects."

        $AllDatabaseObjectsFileContent = $AllDatabaseObjectsFileContent.Replace('@MemoryOptimizedObjectUsage bit = 1', '@MemoryOptimizedObjectUsage bit = 0')
    }

    $AllDatabaseObjectsFileContent | Out-File -FilePath $tempAllDatabaseObjectsFilePath
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Runtime All Server Objects file code is generated."

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$AllDatabaseObjectsFilePath = '$tempAllDatabaseObjectsFilePath'"
    try {
        if($verbose -or $debug) {
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $tempAllDatabaseObjectsFilePath) -EnableException -MessagesToOutput -Verbose:$false -Debug:$false | Write-Verbose
        }
        else {
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $tempAllDatabaseObjectsFilePath) -EnableException -Verbose:$false -Debug:$false
        }
    }
    catch {
        $errMessage = $_

        if ($ReturnInlineErrorMessage) {
		    "Below error occurred while trying to execute script '$tempAllDatabaseObjectsFilePath'.`n$($errMessage.Exception.Message)" | Write-Error
	    }
	    else {            
		    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Below error occurred while trying to execute script '$tempAllDatabaseObjectsFilePath'." | Write-Host -ForegroundColor Red
            $($errMessage.Exception.Message -Split [Environment]::NewLine) | % {"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "$_"} | Write-Host -ForegroundColor Red

            Write-Error "Stop here. Fix above issue."
        }
    }

    # Cleanup temporary file path
    if($true) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Remove temp file '$tempAllDatabaseObjectsFilePath'.."
        Remove-Item -Path $tempAllDatabaseObjectsFilePath | Out-Null
    }

    # Update InventoryServer Objects
    if($InventoryServer -eq $SqlInstanceToBaseline)
    {
        if($SkipInventorySteps) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Skip Update of Objects on Inventory Server as `$SkipInventorySteps is true.."
        }
        else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update objects on Inventory Server.."
            $InventorySpecificObjectsFileText = [System.IO.File]::ReadAllText($InventorySpecificObjectsFilePath)

            # Modify AllDatabaseObjectsFileContent if MemoryOptimized Objects are NOT to be used
            if($MemoryOptimizedObjectsUsage -eq $false) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$MemoryOptimizedObjectsUsage is false. So use disk-based objects for Inventory objects."
                $InventorySpecificObjectsFileText = $InventorySpecificObjectsFileText.Replace('@MemoryOptimizedObjectUsage bit = 1', '@MemoryOptimizedObjectUsage bit = 0')
            }

            $dbaDatabaseParentPath = Split-Path $dbaDatabasePath -Parent
            $memoryOptimizedFilePath = if($dbaDatabaseParentPath -notmatch '\\$') { "$dbaDatabaseParentPath\MemoryOptimized.ndf" } else { "$($dabaDatabaseParentPath)MemoryOptimized.ndf" }
            #$InventorySpecificObjectsFileText = $InventorySpecificObjectsFileText.Replace('E:\Data\MemoryOptimized.ndf', "$(Join-Path $dbaDatabaseParentPath 'MemoryOptimized.ndf')")
            $InventorySpecificObjectsFileText = $InventorySpecificObjectsFileText.Replace('E:\Data\MemoryOptimized.ndf', $memoryOptimizedFilePath)

            $InventorySpecificObjectsFileText = $InventorySpecificObjectsFileText.Replace('dba_team@gmail.com', $($DbaGroupMailId -join ';'))
            $InventorySpecificObjectsFileText = $InventorySpecificObjectsFileText.Replace('dba.manager@gmail.com', $($DbaGroupMailId -join ';'))
            $InventorySpecificObjectsFileText = $InventorySpecificObjectsFileText.Replace('sre.vp@gmail.com', $($DbaGroupMailId -join ';'))
            $InventorySpecificObjectsFileText = $InventorySpecificObjectsFileText.Replace('cto@gmail.com', $($DbaGroupMailId -join ';'))
            $InventorySpecificObjectsFileText = $InventorySpecificObjectsFileText.Replace('noc@gmail.com', $($DbaGroupMailId -join ';'))
            $InventorySpecificObjectsFileText = $InventorySpecificObjectsFileText.Replace('http://localhost:3000/d/', $GrafanaDashboardPortal)
        

            if($verbose -or $debug) {
                $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $InventorySpecificObjectsFileText -EnableException -MessagesToOutput -Verbose:$false -Debug:$false | Write-Verbose
            }
            else {
                $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $InventorySpecificObjectsFileText -EnableException -Verbose:$false -Debug:$false
            }
        }
    }

    if ($PSBoundParameters.ContainsKey('EnableEmailAlerts')) {
        $emailValue  = if ($EnableEmailAlerts) { '1' } else { '0' }
        $emailReason = "Set by Install-SQLMonitor.ps1 on $(Get-Date -Format 'u')"
        $sqlUpsert   = @"
UPDATE dbo.sma_params
SET    param_value = '$emailValue',
       remarks     = '$emailReason'
WHERE  param_key   = 'email_delivery_enabled';
"@
        $conInventoryServer |
            Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlUpsert `
                            -EnableException -Verbose:$false -Debug:$false
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCaptureAlertMessagesFilePath = '$UspCaptureAlertMessagesFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCaptureAlertMessagesFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCreateAgentAlertsFilePath = '$UspCreateAgentAlertsFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCreateAgentAlertsFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCollectWaitStatsFilePath = '$UspCollectWaitStatsFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCollectWaitStatsFilePath) -EnableException -Verbose:$false -Debug:$false

    if($IsManagedInstance -or $InstanceScopeFeaturesOnly) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCollectXeventMetricsRingBufferFilePath = '$UspCollectXeventMetricsRingBufferFilePath'"
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCollectXeventMetricsRingBufferFilePath) -EnableException -Verbose:$false -Debug:$false
    }
    else {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCollectXeventMetricsFilePath = '$UspCollectXeventMetricsFilePath'"
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCollectXeventMetricsFilePath) -EnableException -Verbose:$false -Debug:$false
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCollectPrivilegedInfoFilePath = '$UspCollectPrivilegedInfoFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCollectPrivilegedInfoFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspLogSaverFilePath = '$UspLogSaverFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspLogSaverFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspTempDbSaverFilePath = '$UspTempDbSaverFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspTempDbSaverFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspPartitionMaintenanceFilePath = '$UspPartitionMaintenanceFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspPartitionMaintenanceFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspPurgeTablesFilePath = '$UspPurgeTablesFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspPurgeTablesFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspActiveRequestsCountFilePath = '$UspActiveRequestsCountFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspActiveRequestsCountFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspWaitsPerCorePerMinuteFilePath = '$UspWaitsPerCorePerMinuteFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspWaitsPerCorePerMinuteFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspAvgDiskWaitMsFilePath = '$UspAvgDiskWaitMsFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspAvgDiskWaitMsFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspAvgDiskLatencyMsFilePath = '$UspAvgDiskLatencyMsFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspAvgDiskLatencyMsFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspEnablePageCompressionFilePath = '$UspEnablePageCompressionFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspEnablePageCompressionFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspRunWhoIsActiveFilePath = '$UspRunWhoIsActiveFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspRunWhoIsActiveFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCollectFileIOStatsFilePath = '$UspCollectFileIOStatsFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCollectFileIOStatsFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCollectMemoryClerksFilePath = '$UspCollectMemoryClerksFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCollectMemoryClerksFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCollectAgHealthStateFilePath = '$UspCollectAgHealthStateFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCollectAgHealthStateFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCollectDiskSpaceFilePath = '$UspCollectDiskSpaceFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCollectDiskSpaceFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCollectPerformanceMetricsFilePath = '$UspCollectPerformanceMetricsFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCollectPerformanceMetricsFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCheckSQLAgentJobsFilePath = '$UspCheckSQLAgentJobsFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspCheckSQLAgentJobsFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspWrapperCollectPrivilegedInfoFilePath = '$UspWrapperCollectPrivilegedInfoFilePath'"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UspWrapperCollectPrivilegedInfoFilePath) -EnableException -Verbose:$false -Debug:$false

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Adding entry into SqlInstanceToBaseline [$SqlInstanceToBaseline].[$DbaDatabase].[dbo].[instance_hosts].."
    $sqlAddInstanceHost = @"
        if not exists (select * from dbo.instance_hosts where host_name = '$HostName')
        begin
	        insert dbo.instance_hosts ([host_name])
	        select [host_name] = '$HostName';
            
            select 'dbo.instance_hosts' as RunningQuery, * from dbo.instance_hosts where [host_name] = '$HostName';
        end
"@
    # Populate $SqlInstanceToBaseline
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlAddInstanceHost -EnableException -Verbose:$false -Debug:$false | Format-Table -AutoSize

    # Populate $SqlInstanceAsDataDestination
    if( ($SqlInstanceAsDataDestination -ne $SqlInstanceToBaseline) -and ($InventoryServer -ne $SqlInstanceAsDataDestination) ) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Adding entry into SqlInstanceAsDataDestination [$SqlInstanceAsDataDestination].[$DbaDatabase].[dbo].[instance_hosts].."
        $conSqlInstanceAsDataDestination | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlAddInstanceHost -EnableException -Verbose:$false -Debug:$false | Format-Table -AutoSize
    }

    # Populate $InventoryServer
    if($InventoryServer -ne $SqlInstanceToBaseline) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Adding entry into InventoryServer [$InventoryServer].[$InventoryDatabase].[dbo].[instance_hosts].."
        $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlAddInstanceHost -EnableException -Verbose:$false -Debug:$false | Format-Table -AutoSize
    }


    # Compute [more_info] as JSON
    if( $ForceSetupOfTaskSchedulerJobs -or $HasCustomizedTsqlJobs -or $HasCustomizedPowerShellJobs -or $InstanceScopeFeaturesOnly -or $IsManagedInstance )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Compute [more_info] column value.."

        $moreInfo = New-Object psobject -Property @{
            ForceSetupOfTaskSchedulerJobs = $ForceSetupOfTaskSchedulerJobs
            HasCustomizedTsqlJobs = $HasCustomizedTsqlJobs
            HasCustomizedPowerShellJobs = $HasCustomizedPowerShellJobs
            InstanceScopeFeaturesOnly = $InstanceScopeFeaturesOnly
            IsManagedInstance = $IsManagedInstance
        }
        $moreInfoJSON = $moreInfo | ConvertTo-Json
    }


    $sqlAddInstanceHostMapping = @"
    if not exists (select * from dbo.instance_details where sql_instance = '$SqlInstanceToBaselineWithOutPort' and [host_name] = '$HostName')
    begin
	    insert dbo.instance_details 
            (   [sql_instance], [sql_instance_port], [host_name], [database], [collector_tsql_jobs_server], 
                [collector_powershell_jobs_server], [data_destination_sql_instance],
                [dba_group_mail_id], [sqlmonitor_script_path]
            )
	    select	[sql_instance] = '$SqlInstanceToBaselineWithOutPort',
                [sql_instance_port] = $(if([String]::IsNullOrEmpty($Port4SqlInstanceToBaseline)){'null'}else{"'$Port4SqlInstanceToBaseline'"}),
			    [host_name] = '$Hostname',
                [database] = '$DbaDatabase',
			    [collector_tsql_jobs_server] = '$SqlInstanceForTsqlJobs',
                [collector_powershell_jobs_server] = '$SqlInstanceForPowershellJobs',
                [data_destination_sql_instance] = '$SqlInstanceAsDataDestination',
                [dba_group_mail_id] = '$($DbaGroupMailId -join ';')',
			    [sqlmonitor_script_path] = '$RemoteSQLMonitorPath'

        select 'dbo.instance_details' as RunningQuery, * from dbo.instance_details where [sql_instance] = '$SqlInstanceToBaselineWithOutPort';
    end
"@

    $sqlAddInstanceHostMappingInInventory = @"
    if not exists (select * from dbo.instance_details where sql_instance = '$SqlInstanceToBaselineWithOutPort' and [host_name] = '$HostName')
    begin
	    insert dbo.instance_details 
            (   [sql_instance], [sql_instance_port], [host_name], [database], [collector_tsql_jobs_server], 
                [collector_powershell_jobs_server], [data_destination_sql_instance],
                [dba_group_mail_id], [sqlmonitor_script_path], [more_info]
            )
	    select	[sql_instance] = '$SqlInstanceToBaselineWithOutPort',
                [sql_instance_port] = $(if([String]::IsNullOrEmpty($Port4SqlInstanceToBaseline)){'null'}else{"'$Port4SqlInstanceToBaseline'"}),
			    [host_name] = '$Hostname',
                [database] = '$DbaDatabase',
			    [collector_tsql_jobs_server] = '$SqlInstanceForTsqlJobs',
                [collector_powershell_jobs_server] = '$SqlInstanceForPowershellJobs',
                [data_destination_sql_instance] = '$SqlInstanceAsDataDestination',
                [dba_group_mail_id] = '$($DbaGroupMailId -join ';')',
			    [sqlmonitor_script_path] = '$RemoteSQLMonitorPath',
                [more_info] = case when '$moreInfoJSON' = '' then null else '$moreInfoJSON' end

        select 'dbo.instance_details' as RunningQuery, * from dbo.instance_details where [sql_instance] = '$SqlInstanceToBaselineWithOutPort';
    end
"@

    if($verbose) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$sqlAddInstanceHostMappingInInventory => `n"
        $sqlAddInstanceHostMappingInInventory | Write-Host -ForegroundColor Cyan
        Write-Host ""
    }
    
    # Populate $SqlInstanceToBaseline
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Adding entry into SqlInstanceToBaseline [$SqlInstanceToBaseline].[$DbaDatabase].[dbo].[instance_details].."
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlAddInstanceHostMapping -EnableException -Verbose:$false -Debug:$false | Format-Table -AutoSize

    # Populate $SqlInstanceAsDataDestination
    if( ($SqlInstanceAsDataDestination -ne $SqlInstanceToBaseline) -and ($InventoryServer -ne $SqlInstanceAsDataDestination) ) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Adding entry into SqlInstanceAsDataDestination [$SqlInstanceAsDataDestination].[$DbaDatabase].[dbo].[instance_details].."
        $conSqlInstanceAsDataDestination | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlAddInstanceHostMapping -EnableException -Verbose:$false -Debug:$false | Format-Table -AutoSize
    }

    # Populate $InventoryServer
    if($InventoryServer -ne $SqlInstanceToBaseline) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Adding entry into InventoryServer [$InventoryServer].[$InventoryDatabase].[dbo].[instance_details].."
        $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlAddInstanceHostMappingInInventory -EnableException -Verbose:$false -Debug:$false | Format-Table -AutoSize
    }

    if($isExpressEdition -or (-not [String]::IsNullOrEmpty($RetentionDays)) ) 
    {
        if($isExpressEdition -and ([String]::IsNullOrEmpty($RetentionDays) -or $RetentionDays -gt 7) ) {
            $RetentionDays = 7
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Since Express Edition, setting retention to $RetentionDays days.." | Write-Host -ForegroundColor Cyan
        }
        else {
            if([String]::IsNullOrEmpty($RetentionDays) -or $RetentionDays -eq 0) {
                $RetentionDays = 14
            }
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Setting retention to $RetentionDays days.." | Write-Host -ForegroundColor Cyan
        }
        
        # Update retention only when table is recently added. For already existing tables, retention should be modified manually
        $sqlSetPurgeThreshold = @"
update dbo.purge_table 
set retention_days = case when table_name like 'dbo.BlitzIndex%' then $RetentionDays*6 
                            when table_name like 'dbo.Blitz%' then $RetentionDays*6 
                            when table_name = 'dbo.disk_space' then  $RetentionDays*6 
                            when table_name = 'dbo.file_io_stats' then  $RetentionDays*6 
                            when table_name = 'dbo.memory_clerks' then  $RetentionDays*6 
                            when table_name = 'dbo.wait_stats' then  $RetentionDays*6 
                            when table_name = 'dbo.xevent_metrics' then  $RetentionDays*6
                            else $RetentionDays
                            end
where 1=1
--and retention_days > $RetentionDays
and created_date >= DATEADD(hour,-2,getdate())
"@
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$sqlSetPurgeThreshold => `n`n`t$sqlSetPurgeThreshold"
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlSetPurgeThreshold -EnableException
    }



    # On $InventoryServer => Validate if dbo.params data has been customized as per need 
    if($InventoryServer -ne $SqlInstanceToBaseline) 
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Checking data for [$InventoryDatabase].dbo.sma_params on [$InventoryServer]."
        $sqlGetSmaParamsData = "select * from [$InventoryDatabase].dbo.sma_params"
        if($verbose) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$sqlGetSmaParamsData => `n`n`t$sqlGetSmaParamsData"
        }
        $resultGetSmaParamsData = @()
        $resultGetSmaParamsData += $conInventoryServer | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlGetSmaParamsData -EnableException

        if($resultGetSmaParamsData.Count -gt 0) {
            $dba_team_email_id = $resultGetSmaParamsData | Where-Object {$_.param_key -eq 'dba_team_email_id'} | Select-Object -ExpandProperty param_value
            $dba_manager_email_id = $resultGetSmaParamsData | Where-Object {$_.param_key -eq 'dba_manager_email_id'} | Select-Object -ExpandProperty param_value
            $sre_vp_email_id = $resultGetSmaParamsData | Where-Object {$_.param_key -eq 'sre_vp_email_id'} | Select-Object -ExpandProperty param_value
            $cto_email_id = $resultGetSmaParamsData | Where-Object {$_.param_key -eq 'cto_email_id'} | Select-Object -ExpandProperty param_value
            $noc_email_id = $resultGetSmaParamsData | Where-Object {$_.param_key -eq 'noc_email_id'} | Select-Object -ExpandProperty param_value
            $dba_team_email_id = $resultGetSmaParamsData | Where-Object {$_.param_key -eq 'dba_team_email_id'} | Select-Object -ExpandProperty param_value
            $GrafanaDashboardPortal = $resultGetSmaParamsData | Where-Object {$_.param_key -eq 'GrafanaDashboardPortal'} | Select-Object -ExpandProperty param_value
            $url_for_dba_slack_channel = $resultGetSmaParamsData | Where-Object {$_.param_key -eq 'url_for_dba_slack_channel'} | Select-Object -ExpandProperty param_value

            if($dba_team_email_id -eq 'dba_team@gmail.com' -or $dba_manager_email_id -eq 'dba.manager@gmail.com') {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly update the [param_value] in table [$InventoryDatabase].dbo.sma_params on [$InventoryServer]." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Then re-run this baselining of inventory server.`n" | Write-Host -ForegroundColor Red
                Start-Sleep -Seconds 1
                "Kindly fix error of above line" | Write-Error -ErrorAction Stop
            }
        }
        else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "No data found in table [$InventoryDatabase].dbo.sma_params on [$InventoryServer].`n" | Write-Host -ForegroundColor Red
            Start-Sleep -Seconds 1
            "Kindly fix error of above line" | Write-Error -ErrorAction Stop
        }
    }
}


# Save additional settings in [$InventoryServer].[$InventoryDatabase].[dbo].[instance_details].[more_info] as JSON
$stepName = 'Post_Step2_Block_Update_moreInfo'
if( ($ForceSetupOfTaskSchedulerJobs -or $HasCustomizedTsqlJobs -or $HasCustomizedPowerShellJobs -or $InstanceScopeFeaturesOnly) `
    -and ($isUpgradeScenario -or ($Steps2Execute -contains '2__AllDatabaseObjects'))
  )
{
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Working on [$stepName].."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Save additional settings in more_info columns.."
    $moreInfo = New-Object psobject -Property @{
        ForceSetupOfTaskSchedulerJobs = $ForceSetupOfTaskSchedulerJobs
        HasCustomizedTsqlJobs = $HasCustomizedTsqlJobs
        HasCustomizedPowerShellJobs = $HasCustomizedPowerShellJobs
        InstanceScopeFeaturesOnly = $InstanceScopeFeaturesOnly
    }
    $moreInfoJSON = $moreInfo | ConvertTo-Json


    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Adding [more_info] into [$SqlInstanceToBaseline].[$DbaDatabase].[dbo].[instance_details].."
    $sqlAddMoreInfo = @"
    update id
    set [more_info] = case when '$moreInfoJSON' = '' then null else '$moreInfoJSON' end
    --select * 
    from dbo.instance_details id
    where sql_instance = '$SqlInstanceToBaselineWithOutPort' and [host_name] = '$HostName'
"@
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlAddMoreInfo -EnableException | ft -AutoSize
}

$stepName = 'Post_Step2_Block_Initialize_newList'
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Working on [$stepName].."
if($Steps2Execute -is [array]) {
    [System.Collections.ArrayList]$newList = $Steps2Execute
}
else {
    [System.Collections.ArrayList]$newList = @()
    $newList.Add($Steps2Execute) | Out-Null
}

$stepName = 'Post_Step2_Block_HasCustomizedTsqlJobs'
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Working on [$stepName].."
if($HasCustomizedTsqlJobs) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "*****Based on `$HasCustomizedTsqlJobs, skipping TSQL Jobs.." | Write-Host -ForegroundColor Yellow
    foreach($tsqJobStepName in $TsqlJobSteps) {
        $newList.Remove($tsqJobStepName) | Out-Null
    }
}

$stepName = 'Post_Step2_Block_HasCustomizedPowerShellJobs'
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Working on [$stepName].."
if($HasCustomizedPowerShellJobs) {
    "$(if(-not $HasCustomizedTsqlJobs){"`n"})$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}`n" -f 'WARNING:', "*****Based on `$HasCustomizedPowerShellJobs, skipping PowerShell Jobs.." | Write-Host -ForegroundColor Yellow
    foreach($psJobStepName in $PowerShellJobSteps) {
        $newList.Remove($psJobStepName) | Out-Null
    }
}

$Steps2Execute = $newList


# 3__XEventSession
$stepName = '3__XEventSession'
if( ($stepName -in $Steps2Execute) -and ($MajorVersion -ge 11) ) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."

    if ($IsManagedInstance -or $InstanceScopeFeaturesOnly) {
        $XEventSessionFilePath = $XEventSessionRingBufferFilePath

        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$XEventSessionFilePath = '$XEventSessionFilePath'"
        $sqlXEventSession = [System.IO.File]::ReadAllText($XEventSessionFilePath)
    }
    else {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$XEventSessionFilePath = '$XEventSessionFilePath'"

        # If no XEvent directory provided, then use same as DBA database
        if( [String]::IsNullOrEmpty($XEventDirectory) )
        {
            $dbaDatabasePathParent = Split-Path $dbaDatabasePath -Parent
            $dbaDatabasePathParent = $dbaDatabasePathParent.Replace('\\','\');
            if($dbaDatabasePathParent.Length -eq 3) {
                $xEventTargetPathDirectory = "${dbaDatabasePathParent}xevents"
            }
            else {
                $xEventTargetPathDirectoryParent = Split-Path $dbaDatabasePathParent -Parent
                if($xEventTargetPathDirectoryParent.Length -eq 3) {
                    $xEventTargetPathDirectory = "$(Split-Path $dbaDatabasePathParent -Parent)xevents"
                }
                else {
                    $xEventTargetPathDirectory = "$($xEventTargetPathDirectoryParent)\xevents"
                }
            }
        }
        else {
            $xEventTargetPathDirectory = $XEventDirectory
            if($XEventDirectory.EndsWith('\')) {
                $xEventTargetPathDirectory = $XEventDirectory.TrimEnd('\')
            }
        }

        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Computed XEvent files directory -> '$xEventTargetPathDirectory'.."
        if(-not (Test-Path $($xEventTargetPathDirectory))) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Create directory '$xEventTargetPathDirectory' for XEvent files.."
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query "EXEC xp_create_subdir '$xEventTargetPathDirectory'" -EnableException
        }

        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Create XEvent session named [xevent_metrics].."
        $sqlXEventSession = [System.IO.File]::ReadAllText($XEventSessionFilePath).Replace('E:\Data\xevents', "$xEventTargetPathDirectory")
    }

    try {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database master -Query $sqlXEventSession -EnableException | Format-Table -AutoSize
    }
    catch {
        $errMessage = $_
        $errMessage | gm
        if($errMessage.Exception.Message -like "The value specified for event attribute or predicate source*") {
            $sqlXEventSession = $sqlXEventSession.Replace("WHERE ( ([duration]>=5000000) OR ([result]<>('OK')) ))", "WHERE ( ([duration]>=5000000) ))")
        }
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database master -Query $sqlXEventSession -EnableException | Format-Table -AutoSize
    }
}


# 4__FirstResponderKitObjects
$stepName = '4__FirstResponderKitObjects'
if($stepName -in $Steps2Execute) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating FirstResponderKit Objects in [master] database.."
    if([String]::IsNullOrEmpty($FirstResponderKitZipFile)) {
        Install-DbaFirstResponderKit -SqlInstance $conSqlInstanceToBaseline -Database master -EnableException -Verbose:$false -Debug:$false | Format-Table -AutoSize
    }
    else {
        Install-DbaFirstResponderKit -SqlInstance $conSqlInstanceToBaseline -Database master -LocalFile $FirstResponderKitZipFile -EnableException -Verbose:$false -Debug:$false | Format-Table -AutoSize
    }
}


# 5__DarlingDataObjects
$stepName = '5__DarlingDataObjects'
if($stepName -in $Steps2Execute) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating DarlingData Objects in [master] database.."
    if([String]::IsNullOrEmpty($DarlingDataZipFile)) {
        Install-DbaDarlingData -SqlInstance $conSqlInstanceToBaseline -Database master -EnableException -Procedure All | Format-Table -AutoSize
    }
    else {
        Install-DbaDarlingData -SqlInstance $conSqlInstanceToBaseline -Database master -LocalFile $DarlingDataZipFile -EnableException -Procedure All | Format-Table -AutoSize
    }
}


# 6__OlaHallengrenSolutionObjects
$stepName = '6__OlaHallengrenSolutionObjects'
if( ($stepName -in $Steps2Execute) -and ([String]::IsNullOrEmpty($OlaHallengrenSolutionZipFile) -eq $false) )  {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating objects of Ola Hallengren Maintenance Solution except jobs in [$DbaDatabase] database.."
    
    Install-DbaMaintenanceSolution -SqlInstance $conSqlInstanceToBaseline -Database $DbaDatabase -ReplaceExisting -LocalFile $OlaHallengrenSolutionZipFile -EnableException | Out-Null
}


# 7__sp_WhatIsRunning
$stepName = '7__sp_WhatIsRunning'
if($stepName -in $Steps2Execute) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$WhatIsRunningFilePath = '$WhatIsRunningFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating sp_WhatIsRunning procedure in [$DbaDatabase] database.."
    if($MajorVersion -ge 11) {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $WhatIsRunningFilePath) -EnableException | Format-Table -AutoSize
    }
    else {
        $sqlWhatIsRunning = [System.IO.File]::ReadAllText($WhatIsRunningFilePath)        
        $sqlWhatIsRunning = $sqlWhatIsRunning.Replace('open_transaction_count = s.open_transaction_count', "open_transaction_count = 0")
        $sqlWhatIsRunning = $sqlWhatIsRunning.Replace('s.database_id', "null")

        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlWhatIsRunning -EnableException
    }
}


# 8__usp_GetAllServerInfo
$stepName = '8__usp_GetAllServerInfo'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerInfoFilePath = '$GetAllServerInfoFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating usp_GetAllServerInfo procedure in [$InventoryServer].[$DbaDatabase] database.."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -File $(Get-Item -LiteralPath $GetAllServerInfoFilePath)

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspWrapperGetAllServerInfoFilePath = '$UspWrapperGetAllServerInfoFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating usp_wrapper_GetAllServerInfo procedure in [$InventoryServer].[$DbaDatabase] database.."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -File $(Get-Item -LiteralPath $UspWrapperGetAllServerInfoFilePath)

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerCollectedDataFilePath = '$GetAllServerCollectedDataFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating usp_GetAllServerCollectedData procedure in [$InventoryServer].[$DbaDatabase] database.."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -File $(Get-Item -LiteralPath $GetAllServerCollectedDataFilePath)

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspWrapperGetAllServerCollectedDataFilePath = '$UspWrapperGetAllServerCollectedDataFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating usp_wrapper_GetAllServerCollectedData procedure in [$InventoryServer].[$DbaDatabase] database.."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -File $(Get-Item -LiteralPath $UspWrapperGetAllServerCollectedDataFilePath)

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspGetAllServerDashboardMailFilePath = '$UspGetAllServerDashboardMailFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating usp_GetAllServerDashboardMail procedure in [$InventoryServer].[$DbaDatabase] database.."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -File $(Get-Item -LiteralPath $UspGetAllServerDashboardMailFilePath)

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCheckInstanceAvailabilityFilePath = '$UspCheckInstanceAvailabilityFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating usp_check_instance_availability procedure in [$InventoryServer].[$DbaDatabase] database.."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -File $(Get-Item -LiteralPath $GetAllServerInfoFilePath)

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspCollectAllServerLoginExpirationInfoFilePath = '$UspCollectAllServerLoginExpirationInfoFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating usp_collect_all_server_login_expiration_info procedure in [$InventoryServer].[$DbaDatabase] database.."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -File $(Get-Item -LiteralPath $UspCollectAllServerLoginExpirationInfoFilePath)

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspPopulateSmaSqlInstanceFilePath = '$UspPopulateSmaSqlInstanceFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating usp_populate_sma_sql_instance procedure in [$InventoryServer].[$DbaDatabase] database.."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -File $(Get-Item -LiteralPath $UspPopulateSmaSqlInstanceFilePath)

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspSendLoginExpiryEmailsFilePath = '$UspSendLoginExpiryEmailsFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating usp_send_login_expiry_emails procedure in [$InventoryServer].[$DbaDatabase] database.."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -File $(Get-Item -LiteralPath $UspSendLoginExpiryEmailsFilePath)

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspComputeAllServerVolatileInfoHistoryHourlyFilePath = '$UspComputeAllServerVolatileInfoHistoryHourlyFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating usp_compute_all_server_volatile_info_history_hourly procedure in [$InventoryServer].[$DbaDatabase] database.."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -File $(Get-Item -LiteralPath $UspComputeAllServerVolatileInfoHistoryHourlyFilePath)

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UspWrapperPopulateSmaSqlInstanceFilePath = '$UspWrapperPopulateSmaSqlInstanceFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating usp_wrapper_populate_sma_sql_instance procedure in [$InventoryServer].[$DbaDatabase] database.."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -File $(Get-Item -LiteralPath $UspWrapperPopulateSmaSqlInstanceFilePath)  -Verbose:$verbose -MessagesToOutput:$verbose

}


# 9__CopyDbaToolsModule2Host
$stepName = '9__CopyDbaToolsModule2Host'
if($stepName -in $Steps2Execute) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$DbaToolsFolderPath = '$DbaToolsFolderPath'"
    $dbaToolsLibraryFolderPath = "$DbaToolsFolderPath.library"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$dbaToolsLibraryFolderPath = '$dbaToolsLibraryFolderPath'"

    # Get dbatools config file path to retrive version info
    $dbatoolsConfigFileFullPath = if($DbaToolsFolderPath.EndsWith('\')) {$DbaToolsFolderPath+"dbatools.psd1"} else {"$DbaToolsFolderPath\dbatools.psd1"}
    if(-not (Test-Path $dbatoolsConfigFileFullPath)) {
        $dbatoolsConfigFile = Get-ChildItem -Path $DbaToolsFolderPath -Recurse -File `
                                | Where-Object {$_.Name -eq 'dbatools.psd1'} | Sort-Object -Property LastWriteTime -Descending `
                                | Select-Object -First 1
        $dbatoolsConfigFileDirectory =  $dbatoolsConfigFile.DirectoryName
        $dbatoolsConfigFileFullPath =  $dbatoolsConfigFile.FullName
    }

    # Read config file and parse data
    Import-LocalizedData -BaseDirectory $dbatoolsConfigFileDirectory -FileName 'dbatools.psd1' -BindingVariable configData

    # Check if module is post 2.x.x
    $isDbatoolsLibraryRequired = $true
    if([int]($configData.ModuleVersion).Split('.')[0] -lt 2) {
        $isDbatoolsLibraryRequired = $false
    }
    
    # Get PSModule path on HostName provided
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Finding valid PSModule path on [$HostName].."
    $remoteModulePath = Invoke-Command -Session $ssn4PerfmonSetup -ScriptBlock {
        $modulePath = $null
        if('C:\Program Files\WindowsPowerShell\Modules' -in $($env:PSModulePath -split ';')) {
            $modulePath = 'C:\Program Files\WindowsPowerShell\Modules'
        }
        else {
            $modulePath = $($env:PSModulePath -split ';') | Where-Object {$_ -like '*Microsoft SQL Server*'} | select -First 1
        }
        $modulePath
    }

    # Copy dbatools module on remote host
    if ($true) 
    {
        $dbatoolsRemotePath = Join-Path $remoteModulePath 'dbatools'
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Copy dbatools module from '$DbaToolsFolderPath' to host [$HostName] on '$dbatoolsRemotePath'.."
    
        if( (Invoke-Command -Session $ssn4PerfmonSetup -ScriptBlock {Test-Path $Using:dbatoolsRemotePath}) ) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "'$dbatoolsRemotePath' already exists on host [$HostName]."
        }
        else {
            Copy-Item $DbaToolsFolderPath -Destination $dbatoolsRemotePath -ToSession $ssn4PerfmonSetup -Recurse
        }

        # Copy dbatools folder on Jobs Server Host
        if( ($SqlInstanceToBaseline -ne $SqlInstanceForPowershellJobs) -and ($ssn4PerfmonSetup -ne $ssnJobServer) )
        {
            "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Finding valid PSModule path on [$($ssnJobServer.ComputerName)].."
            $remoteModulePath = Invoke-Command -Session $ssnJobServer -ScriptBlock {
                $modulePath = $null
                if('C:\Program Files\WindowsPowerShell\Modules' -in $($env:PSModulePath -split ';')) {
                    $modulePath = 'C:\Program Files\WindowsPowerShell\Modules'
                }
                else {
                    $modulePath = $($env:PSModulePath -split ';') | Where-Object {$_ -like '*Microsoft SQL Server*'} | select -First 1
                }
                $modulePath
            }

            $dbatoolsRemotePath = Join-Path $remoteModulePath 'dbatools'
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Copy dbatools module from '$DbaToolsFolderPath' to host [$($ssnJobServer.ComputerName)] on '$dbatoolsRemotePath'.."
    
            if( (Invoke-Command -Session $ssnJobServer -ScriptBlock {Test-Path $Using:dbatoolsRemotePath}) ) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "'$dbatoolsRemotePath' already exists on host [$($ssnJobServer.ComputerName)]."
            }
            else {
                Copy-Item $DbaToolsFolderPath -Destination $dbatoolsRemotePath -ToSession $ssnJobServer -Recurse
            }
        }
    }


    # Copy dbatools.library module on remote host
    if ($isDbatoolsLibraryRequired) 
    {
        $dbatoolsLibraryRemotePath = Join-Path $remoteModulePath 'dbatools.library'
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Copy dbatools.library module from '$dbaToolsLibraryFolderPath' to host [$HostName] on '$dbatoolsLibraryRemotePath'.."
    
        if( (Invoke-Command -Session $ssn4PerfmonSetup -ScriptBlock {Test-Path $Using:dbatoolsLibraryRemotePath}) ) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "'$dbatoolsLibraryRemotePath' already exists on host [$HostName]."
        }
        else {
            Copy-Item $dbaToolsLibraryFolderPath -Destination $dbatoolsLibraryRemotePath -ToSession $ssn4PerfmonSetup -Recurse
        }

        # Copy dbatools.library folder on Jobs Server Host
        if( ($SqlInstanceToBaseline -ne $SqlInstanceForPowershellJobs) -and ($ssn4PerfmonSetup -ne $ssnJobServer) )
        {
            "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Finding valid PSModule path on [$($ssnJobServer.ComputerName)].."
            $remoteModulePath = Invoke-Command -Session $ssnJobServer -ScriptBlock {
                $modulePath = $null
                if('C:\Program Files\WindowsPowerShell\Modules' -in $($env:PSModulePath -split ';')) {
                    $modulePath = 'C:\Program Files\WindowsPowerShell\Modules'
                }
                else {
                    $modulePath = $($env:PSModulePath -split ';') | Where-Object {$_ -like '*Microsoft SQL Server*'} | select -First 1
                }
                $modulePath
            }

            $dbatoolsLibraryRemotePath = Join-Path $remoteModulePath 'dbatools.library'
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Copy dbatools.library module from '$dbaToolsLibraryFolderPath' to host [$($ssnJobServer.ComputerName)] on '$dbatoolsLibraryRemotePath'.."
    
            if( (Invoke-Command -Session $ssnJobServer -ScriptBlock {Test-Path $Using:dbatoolsLibraryRemotePath}) ) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "'$dbatoolsLibraryRemotePath' already exists on host [$($ssnJobServer.ComputerName)]."
            }
            else {
                Copy-Item $dbaToolsLibraryFolderPath -Destination $dbatoolsLibraryRemotePath -ToSession $ssnJobServer -Recurse
            }
        }
    }
}


# 10__CopyPerfmonFolder2Host
$stepName = '10__CopyPerfmonFolder2Host'
if($stepName -in $Steps2Execute) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$psScriptPath = '$psScriptPath'"
    
    # Copy SQLMonitor folder on HostName provided
    if( (Invoke-Command -Session $ssn4PerfmonSetup -ScriptBlock {Test-Path $Using:RemoteSQLMonitorPath}) ) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Sync '$RemoteSQLMonitorPath' on [$HostName] from local copy '$psScriptPath'.."
        Copy-Item "$psScriptPath\*" -Destination "$RemoteSQLMonitorPath" -ToSession $ssn4PerfmonSetup -Exclude "*.blg" -Recurse -Force
    }else {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Copy '$psScriptPath' to '$RemoteSQLMonitorPath' on [$HostName].."
        Copy-Item $psScriptPath -Destination $RemoteSQLMonitorPath -ToSession $ssn4PerfmonSetup -Exclude "*.blg" -Recurse -Force
    }

    # Copy SQLMonitor folder on Jobs Server Host
    if( ($SqlInstanceToBaseline -ne $SqlInstanceForPowershellJobs) -and ($ssn4PerfmonSetup -ne $ssnJobServer) )
    {
        if( (Invoke-Command -Session $ssn4PerfmonSetup -ScriptBlock {Test-Path $Using:RemoteSQLMonitorPath}) ) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Sync '$RemoteSQLMonitorPath' on [$HostName] from local copy '$psScriptPath'.."
            Copy-Item "$psScriptPath\*" -Destination "$RemoteSQLMonitorPath" -ToSession $ssn4PerfmonSetup -Exclude "*.blg" -Recurse -Force
        }else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Copy '$psScriptPath' to '$RemoteSQLMonitorPath' on [$HostName].."
            Copy-Item $psScriptPath -Destination $RemoteSQLMonitorPath -ToSession $ssn4PerfmonSetup -Exclude "*.blg" -Recurse -Force
        }
    }
}


# 11__SetupPerfmonDataCollector
$stepName = '11__SetupPerfmonDataCollector'
if($stepName -in $Steps2Execute) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Setup Data Collector set 'DBA' on host '$HostName'.."
    Invoke-Command -Session $ssn4PerfmonSetup -ScriptBlock {
        # Set execution policy
        Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Unrestricted -Force 
        & "$Using:RemoteSQLMonitorPath\perfmon-collector-logman.ps1" -TemplatePath "$Using:RemoteSQLMonitorPath\DBA_PerfMon_All_Counters_Template.xml" -ReSetupCollector $true
    }
}

# If non-domain server, then added HostName in credential name
$stepName = 'Post_Step11_Block_CredentialName_Formatter'
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Working on [$stepName].."
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Add hostname in Credential name if non-domain server.."
if(-not [String]::IsNullOrEmpty($WindowsCredential))
{
    if( $domain -in @('WORKGROUP','WORKGROUP.com') -and (-not $WindowsCredential.UserName.Contains('\')) ) {
        $credentialName = "$HostName\$($WindowsCredential.UserName)"
    }
    else {
        $credentialName = $WindowsCredential.UserName
    }
    $credentialPassword = $WindowsCredential.Password
}

# 12__CreateCredentialProxy. Create Credential & Proxy on SQL Server. If Instance being baselined is same as data collector job owner
$stepName = '12__CreateCredentialProxy'
if( $requireProxy -and ($stepName -in $Steps2Execute) ) 
{
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."

    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Create new SQL Credential [$credentialName] on [$SqlInstanceForPowershellJobs].."
    $dbaCredential = @()
    $dbaCredential += Get-DbaCredential -SqlInstance $conSqlInstanceForPowershellJobs -Name $credentialName -EnableException
    if($dbaCredential.Count -eq 0) {
        New-DbaCredential -SqlInstance $conSqlInstanceForPowershellJobs -Identity $credentialName -SecurePassword $credentialPassword -EnableException
    } else {
        "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "SQL Credential [$credentialName] already exists on [$SqlInstanceForPowershellJobs].."
    }
    $dbaAgentProxy = @()
    $dbaAgentProxy += Get-DbaAgentProxy -SqlInstance $conSqlInstanceForPowershellJobs -Proxy $credentialName -EnableException
    if($dbaAgentProxy.Count -eq 0) {
        New-DbaAgentProxy -SqlInstance $conSqlInstanceForPowershellJobs -Name $credentialName -ProxyCredential $credentialName -SubSystem CmdExec, PowerShell -EnableException
    } else {
        "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "SQL Agent Proxy [$credentialName] already exists on [$SqlInstanceForPowershellJobs].."
    }
}

# 13__CreateJobCollectDiskSpace
$stepName = '13__CreateJobCollectDiskSpace'
if($stepName -in $Steps2Execute) {
    if ($SkipPowerShellJobs4Host -and ($InstanceScopeFeaturesOnly -eq $false)) {
        "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Skipping step '$stepName'.."
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Required DiskSpace Collection job for host [$HostName] should already be present on [$hostOwnerServer] server."
    }
}
if($stepName -in $Steps2Execute -and $SkipPowerShellJobs4Host -eq $false -and $InstanceScopeFeaturesOnly -eq $false) 
{
    $jobName = '(dba) Collect-DiskSpace'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CollectDiskSpaceJobFilePath = '$CollectDiskSpaceJobFilePath'"

    # Append HostName if Job Server is different
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceAsDataDestinationWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceAsDataDestination"
    if( ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForPowershellJobsWithOutPort) -or ($HostName -ne $jobServerDbServiceInfo.host_name) -or ($isClustered -eq $true) ) {
        $jobNameNew = "$jobName - $HostName"
        #$sqlInstanceOnJobStep = $SqlInstanceAsDataDestination
    }    

    # Replace defaults
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForPowershellJobs].."
    $sqlCreateJobCollectDiskSpace = [System.IO.File]::ReadAllText($CollectDiskSpaceJobFilePath) #.Replace('-SqlInstance localhost', "-SqlInstance ''$sqlInstanceOnJobStep''")
    $sqlCreateJobCollectDiskSpace = $sqlCreateJobCollectDiskSpace.Replace('-Database DBA', "-Database `"$DbaDatabase`"")
    $sqlCreateJobCollectDiskSpace = $sqlCreateJobCollectDiskSpace.Replace('-SqlInstance localhost', "-SqlInstance `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobCollectDiskSpace = $sqlCreateJobCollectDiskSpace.Replace('-HostName localhost', "-HostName `"$HostName`"")

    if($jobNameNew -ne $jobName) {
        $sqlCreateJobCollectDiskSpace = $sqlCreateJobCollectDiskSpace.Replace($jobName, $jobNameNew)
    }

    if($RemoteSQLMonitorPath -ne 'C:\SQLMonitor') {
        $sqlCreateJobCollectDiskSpace = $sqlCreateJobCollectDiskSpace.Replace('C:\SQLMonitor', $RemoteSQLMonitorPath)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CollectDiskSpaceJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobCollectDiskSpace -match "@command=N'powershell.exe(?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 10 minutes"
        $timeIntervalMinutes = "00:10"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'powershell' -Argument "$jobArguments *> '$logsPath\$jobName.txt'"
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($DropCreatePowerShellJobs) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
    }
    else # If not express edition
    {
        if($DropCreatePowerShellJobs) {
            $tsqlSSMSValidation = "and APP_NAME() = 'Microsoft SQL Server Management Studio - Query'"
            $sqlCreateJobCollectDiskSpace = $sqlCreateJobCollectDiskSpace.Replace($tsqlSSMSValidation, "--$tsqlSSMSValidation")
        }
        $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobCollectDiskSpace -EnableException

        if($requireProxy) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update job [$jobNameNew] to run under proxy [$credentialName].."
            $sqlUpdateJob = "EXEC msdb.dbo.sp_update_jobstep @job_name=N'$jobNameNew', @step_id=1 ,@proxy_name=N'$credentialName';"
            $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query $sqlUpdateJob -EnableException
        }
        $sqlStartJob = "EXEC msdb.dbo.sp_start_job @job_name=N'$jobNameNew';"
        $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query $sqlStartJob -EnableException
    }
}
elseif($stepName -in $Steps2Execute -and $InstanceScopeFeaturesOnly -eq $true)
{
    $jobName = '(dba) Collect-DiskSpace'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CollectDiskSpaceJobFilePath = '$CollectDiskSpaceJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobCollectDiskSpace = [System.IO.File]::ReadAllText($CollectDiskSpaceJobFilePath)

    if($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) 
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCreateJobCollectDiskSpace = $sqlCreateJobCollectDiskSpace.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlCreateJobCollectDiskSpace)
        #$sqlCreateJobCollectDiskSpace = $flagsPattern.Replace($sqlCreateJobCollectDiskSpace, "@flags=`${flagsValue}, @database_name=N'$DbaDatabase'")
        $sqlCreateJobCollectDiskSpace = $flagsPattern.Replace($sqlCreateJobCollectDiskSpace, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while($cmdPattern.IsMatch($sqlCreateJobCollectDiskSpace)) {
            #$Matches = $cmdPattern.Match($testJobStepSQL)
            #$Matches.Groups
            $sqlCreateJobCollectDiskSpace = $cmdPattern.Replace($sqlCreateJobCollectDiskSpace, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCreateJobCollectDiskSpace = $sqlCreateJobCollectDiskSpace.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobCollectDiskSpace = $sqlCreateJobCollectDiskSpace.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCreateJobCollectDiskSpace = $sqlCreateJobCollectDiskSpace.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
   
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobCollectDiskSpace = $sqlCreateJobCollectDiskSpace.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CollectDiskSpaceJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobCollectDiskSpace -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 1 minutes"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobCollectDiskSpace -EnableException
    }
}


# 14__CreateJobCollectPerfmonData
$stepName = '14__CreateJobCollectPerfmonData'
if($stepName -in $Steps2Execute) {
    if ($SkipPowerShellJobs4Host -and $InstanceScopeFeaturesOnly -eq $false) {
        "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Skipping step '$stepName'.."
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Required Perfmon Data Collection job for host [$HostName] should already be present on [$hostOwnerServer] server."
    }
}
if($stepName -in $Steps2Execute -and $SkipPowerShellJobs4Host -eq $false -and $InstanceScopeFeaturesOnly -eq $false) 
{
    $jobName = '(dba) Collect-PerfmonData'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CollectPerfmonDataJobFilePath = '$CollectPerfmonDataJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceAsDataDestinationWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceAsDataDestination"
    if( ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForPowershellJobsWithOutPort) -or ($HostName -ne $jobServerDbServiceInfo.host_name) -or ($isClustered -eq $true) ) {
        $jobNameNew = "$jobName - $HostName"
        #$sqlInstanceOnJobStep = $SqlInstanceAsDataDestination
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForPowershellJobs].."
    $sqlCreateJobCollectPerfmonData = [System.IO.File]::ReadAllText($CollectPerfmonDataJobFilePath) #.Replace('-SqlInstance localhost', "-SqlInstance ''$sqlInstanceOnJobStep''")
    $sqlCreateJobCollectPerfmonData = $sqlCreateJobCollectPerfmonData.Replace('-Database DBA', "-Database `"$DbaDatabase`"")
    $sqlCreateJobCollectPerfmonData = $sqlCreateJobCollectPerfmonData.Replace('-SqlInstance localhost', "-SqlInstance `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobCollectPerfmonData = $sqlCreateJobCollectPerfmonData.Replace('-HostName localhost', "-HostName `"$HostName`"")
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobCollectPerfmonData = $sqlCreateJobCollectPerfmonData.Replace($jobName, $jobNameNew)
    }
    
    if($RemoteSQLMonitorPath -ne 'C:\SQLMonitor') {
        $sqlCreateJobCollectPerfmonData = $sqlCreateJobCollectPerfmonData.Replace('C:\SQLMonitor', $RemoteSQLMonitorPath)
    }
    
    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CollectPerfmonDataJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobCollectPerfmonData -match "@command=N'powershell.exe(?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Provide WindowsCredential for accessing server [$ssnHostName] of domain '$($sqlServerInfo.domain)'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 1 minutes"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'powershell' -Argument "$jobArguments *> '$logsPath\$jobName.txt'"
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($DropCreatePowerShellJobs) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
    }
    else # If not express edition
    {
        if($DropCreatePowerShellJobs) {
            $tsqlSSMSValidation = "and APP_NAME() = 'Microsoft SQL Server Management Studio - Query'"
            $sqlCreateJobCollectPerfmonData = $sqlCreateJobCollectPerfmonData.Replace($tsqlSSMSValidation, "--$tsqlSSMSValidation")
        }
        $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobCollectPerfmonData -EnableException

        if($requireProxy) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update job [$jobNameNew] to run under proxy [$credentialName].."
            $sqlUpdateJob = "EXEC msdb.dbo.sp_update_jobstep @job_name=N'$jobNameNew', @step_id=1 ,@proxy_name=N'$credentialName';"
            $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query $sqlUpdateJob -EnableException
        }
        $sqlStartJob = "EXEC msdb.dbo.sp_start_job @job_name=N'$jobNameNew';"
        $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query $sqlStartJob -EnableException
    }
}
elseif($stepName -in $Steps2Execute -and $InstanceScopeFeaturesOnly -eq $true)
{
    $jobName = '(dba) Collect-PerfmonData'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CollectPerfmonDataJobFilePath = '$CollectPerfmonDataJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobCollectPerfmonData = [System.IO.File]::ReadAllText($CollectPerfmonDataJobFilePath)

    if($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) 
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCreateJobCollectPerfmonData = $sqlCreateJobCollectPerfmonData.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlCreateJobCollectPerfmonData)
        #$sqlCreateJobCollectPerfmonData = $flagsPattern.Replace($sqlCreateJobCollectPerfmonData, "@flags=`${flagsValue}, @database_name=N'$DbaDatabase'")
        $sqlCreateJobCollectPerfmonData = $flagsPattern.Replace($sqlCreateJobCollectPerfmonData, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while($cmdPattern.IsMatch($sqlCreateJobCollectPerfmonData)) {
            #$Matches = $cmdPattern.Match($testJobStepSQL)
            #$Matches.Groups
            $sqlCreateJobCollectPerfmonData = $cmdPattern.Replace($sqlCreateJobCollectPerfmonData, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCreateJobCollectPerfmonData = $sqlCreateJobCollectPerfmonData.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobCollectPerfmonData = $sqlCreateJobCollectPerfmonData.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCreateJobCollectPerfmonData = $sqlCreateJobCollectPerfmonData.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
   
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobCollectPerfmonData = $sqlCreateJobCollectPerfmonData.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CollectPerfmonDataJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobCollectPerfmonData -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 1 minutes"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobCollectPerfmonData -EnableException
    }
}


# 15__CreateJobCollectWaitStats
$stepName = '15__CreateJobCollectWaitStats'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Collect-WaitStats'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CollectWaitStatsJobFilePath = '$CollectWaitStatsJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobCollectWaitStats = [System.IO.File]::ReadAllText($CollectWaitStatsJobFilePath)

    if($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) 
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCreateJobCollectWaitStats = $sqlCreateJobCollectWaitStats.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlCreateJobCollectWaitStats)
        #$sqlCreateJobCollectWaitStats = $flagsPattern.Replace($sqlCreateJobCollectWaitStats, "@flags=`${flagsValue}, @database_name=N'$DbaDatabase'")
        $sqlCreateJobCollectWaitStats = $flagsPattern.Replace($sqlCreateJobCollectWaitStats, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while($cmdPattern.IsMatch($sqlCreateJobCollectWaitStats)) {
            #$Matches = $cmdPattern.Match($testJobStepSQL)
            #$Matches.Groups
            $sqlCreateJobCollectWaitStats = $cmdPattern.Replace($sqlCreateJobCollectWaitStats, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCreateJobCollectWaitStats = $sqlCreateJobCollectWaitStats.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobCollectWaitStats = $sqlCreateJobCollectWaitStats.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCreateJobCollectWaitStats = $sqlCreateJobCollectWaitStats.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    $sqlCreateJobCollectWaitStats = $sqlCreateJobCollectWaitStats.Replace('sqlmonitor_inventory_server', $InventoryServer)

    if($jobNameNew -ne $jobName) {
        $sqlCreateJobCollectWaitStats = $sqlCreateJobCollectWaitStats.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CollectWaitStatsJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobCollectWaitStats -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 10 minutes"
        $timeIntervalMinutes = "00:10"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobCollectWaitStats -EnableException
    }
}


# 16__CreateJobCollectXEvents
$stepName = '16__CreateJobCollectXEvents'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Collect-XEvents'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CollectXEventsJobFilePath = '$CollectXEventsJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobCollectXEvents = [System.IO.File]::ReadAllText($CollectXEventsJobFilePath)

    if($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) 
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCreateJobCollectXEvents = $sqlCreateJobCollectXEvents.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlCreateJobCollectXEvents)
        #$sqlCreateJobCollectXEvents = $flagsPattern.Replace($sqlCreateJobCollectXEvents, "@flags=`${flagsValue}, @database_name=N'$DbaDatabase'")
        $sqlCreateJobCollectXEvents = $flagsPattern.Replace($sqlCreateJobCollectXEvents, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while($cmdPattern.IsMatch($sqlCreateJobCollectXEvents)) {
            #$Matches = $cmdPattern.Match($sqlCreateJobCollectXEvents)
            #$Matches.Groups
            $sqlCreateJobCollectXEvents = $cmdPattern.Replace($sqlCreateJobCollectXEvents, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCreateJobCollectXEvents = $sqlCreateJobCollectXEvents.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobCollectXEvents = $sqlCreateJobCollectXEvents.Replace('-d DBA', "-d `"$DbaDatabase`"")
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobCollectXEvents = $sqlCreateJobCollectXEvents.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CollectXEventsJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobCollectXEvents -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 1 minutes"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:01" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:01" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobCollectXEvents -EnableException
    }
}


# 17__CreateJobCollectFileIOStats
$stepName = '17__CreateJobCollectFileIOStats'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Collect-FileIOStats'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CollectFileIOStatsJobFilePath = '$CollectFileIOStatsJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobFileIOStats = [System.IO.File]::ReadAllText($CollectFileIOStatsJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCreateJobFileIOStats = $sqlCreateJobFileIOStats.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlCreateJobFileIOStats)
        #$sqlCreateJobFileIOStats = $flagsPattern.Replace($sqlCreateJobFileIOStats, "@flags=`${flagsValue}, @database_name=N'$DbaDatabase'")
        $sqlCreateJobFileIOStats = $flagsPattern.Replace($sqlCreateJobFileIOStats, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlCreateJobFileIOStats)) {
            #$Matches = $cmdPattern.Match($sqlCreateJobFileIOStats)
            #$Matches.Groups
            $sqlCreateJobFileIOStats = $cmdPattern.Replace($sqlCreateJobFileIOStats, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCreateJobFileIOStats = $sqlCreateJobFileIOStats.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobFileIOStats = $sqlCreateJobFileIOStats.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCreateJobFileIOStats = $sqlCreateJobFileIOStats.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    $sqlCreateJobFileIOStats = $sqlCreateJobFileIOStats.Replace('sqlmonitor_inventory_server', $InventoryServer)
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobFileIOStats = $sqlCreateJobFileIOStats.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CollectFileIOStatsJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobFileIOStats -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 10 minutes"
        $timeIntervalMinutes = "00:10"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop        
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobFileIOStats -EnableException
    }
}


# 18__CreateJobPartitionsMaintenance
$stepName = '18__CreateJobPartitionsMaintenance'
if($stepName -in $Steps2Execute -and $IsNonPartitioned -eq $false) 
{
    $jobName = '(dba) Partitions-Maintenance'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$PartitionsMaintenanceJobFilePath = '$PartitionsMaintenanceJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlPartitionsMaintenance = [System.IO.File]::ReadAllText($PartitionsMaintenanceJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlPartitionsMaintenance = $sqlPartitionsMaintenance.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlPartitionsMaintenance)
        #$sqlPartitionsMaintenance = $flagsPattern.Replace($sqlPartitionsMaintenance, "@flags=`${flagsValue}, @database_name=N'$DbaDatabase'")
        $sqlPartitionsMaintenance = $flagsPattern.Replace($sqlPartitionsMaintenance, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlPartitionsMaintenance)) {
            #$Matches = $cmdPattern.Match($sqlPartitionsMaintenance)
            #$Matches.Groups
            $sqlPartitionsMaintenance = $cmdPattern.Replace($sqlPartitionsMaintenance, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlPartitionsMaintenance = $sqlPartitionsMaintenance.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlPartitionsMaintenance = $sqlPartitionsMaintenance.Replace('-d DBA', "-d `"$DbaDatabase`"")
    if($jobNameNew -ne $jobName) {
        $sqlPartitionsMaintenance = $sqlPartitionsMaintenance.Replace($jobName, $jobNameNew)
    }
    
    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$PartitionsMaintenanceJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlPartitionsMaintenance -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily at 12:00 am"
        $timeIntervalMinutes = "00:10"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                #$durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')
                $durationString = '00:00'

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                #$timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlPartitionsMaintenance -EnableException
    }
}


# 19__CreateJobPurgeTables
$stepName = '19__CreateJobPurgeTables'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Purge-Tables'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$PurgeTablesJobFilePath = '$PurgeTablesJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlPurgeDbaMetrics = [System.IO.File]::ReadAllText($PurgeTablesJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlPurgeDbaMetrics = $sqlPurgeDbaMetrics.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlPurgeDbaMetrics)
        $sqlPurgeDbaMetrics = $flagsPattern.Replace($sqlPurgeDbaMetrics, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlPurgeDbaMetrics)) {
            #$Matches = $cmdPattern.Match($sqlPurgeDbaMetrics)
            #$Matches.Groups
            $sqlPurgeDbaMetrics = $cmdPattern.Replace($sqlPurgeDbaMetrics, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlPurgeDbaMetrics = $sqlPurgeDbaMetrics.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlPurgeDbaMetrics = $sqlPurgeDbaMetrics.Replace('-d DBA', "-d `"$DbaDatabase`"")
    if($jobNameNew -ne $jobName) {
        $sqlPurgeDbaMetrics = $sqlPurgeDbaMetrics.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$PurgeTablesJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlPurgeDbaMetrics -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily at 12:00 am"
        $timeIntervalMinutes = "00:10"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                #$durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')
                $durationString = '00:00'

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                #$timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlPurgeDbaMetrics -EnableException
    }
}


# 20__CreateJobRemoveXEventFiles
$stepName = '20__CreateJobRemoveXEventFiles'
if($InstanceScopeFeaturesOnly -and $stepName -in $Steps2Execute) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Skipping step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "This job is not required on Managed Instance."
}
if($stepName -in $Steps2Execute -and $InstanceScopeFeaturesOnly -eq $false) 
{
    $jobName = '(dba) Remove-XEventFiles'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$RemoveXEventFilesJobFilePath = '$RemoveXEventFilesJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if( ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForPowershellJobsWithOutPort) ) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForPowershellJobs].."
    $sqlCreateJobRemoveXEventFiles = [System.IO.File]::ReadAllText($RemoveXEventFilesJobFilePath)

    $sqlCreateJobRemoveXEventFiles = $sqlCreateJobRemoveXEventFiles.Replace('-SqlInstance localhost', "-SqlInstance `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobRemoveXEventFiles = $sqlCreateJobRemoveXEventFiles.Replace('-Database DBA', "-Database `"$DbaDatabase`"")
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobRemoveXEventFiles = $sqlCreateJobRemoveXEventFiles.Replace($jobName, $jobNameNew)
    }

    if($RemoteSQLMonitorPath -ne 'C:\SQLMonitor') {
        $sqlCreateJobRemoveXEventFiles = $sqlCreateJobRemoveXEventFiles.Replace('C:\SQLMonitor', $RemoteSQLMonitorPath)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$RemoveXEventFilesJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobRemoveXEventFiles -match "@command=N'powershell.exe(?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 10 minutes"
        $timeIntervalMinutes = "00:10"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'powershell' -Argument "$jobArguments *> '$logsPath\$jobName.txt'"
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($DropCreatePowerShellJobs) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        if($DropCreatePowerShellJobs) {
            $tsqlSSMSValidation = "and APP_NAME() = 'Microsoft SQL Server Management Studio - Query'"
            $sqlCreateJobRemoveXEventFiles = $sqlCreateJobRemoveXEventFiles.Replace($tsqlSSMSValidation, "--$tsqlSSMSValidation")
        }
        $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobRemoveXEventFiles -EnableException

        if($requireProxy) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update job [$jobNameNew] to run under proxy [$credentialName].."
            $sqlUpdateJob = "EXEC msdb.dbo.sp_update_jobstep @job_name=N'$jobNameNew', @step_id=1 ,@proxy_name=N'$credentialName';"
            $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query $sqlUpdateJob -EnableException
        }
        $sqlStartJob = "EXEC msdb.dbo.sp_start_job @job_name=N'$jobNameNew';"
        $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query $sqlStartJob -EnableException
    }
}


# 21__CreateJobRunLogSaver
$stepName = '21__CreateJobRunLogSaver'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Run-LogSaver'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$RunLogSaverJobFilePath = '$RunLogSaverJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobRunLogSaver = [System.IO.File]::ReadAllText($RunLogSaverJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCreateJobRunLogSaver = $sqlCreateJobRunLogSaver.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlCreateJobRunLogSaver)
        $sqlCreateJobRunLogSaver = $flagsPattern.Replace($sqlCreateJobRunLogSaver, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlCreateJobRunLogSaver)) {
            #$Matches = $cmdPattern.Match($sqlCreateJobRunLogSaver)
            #$Matches.Groups
            $sqlCreateJobRunLogSaver = $cmdPattern.Replace($sqlCreateJobRunLogSaver, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCreateJobRunLogSaver = $sqlCreateJobRunLogSaver.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobRunLogSaver = $sqlCreateJobRunLogSaver.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCreateJobRunLogSaver = $sqlCreateJobRunLogSaver.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    $sqlCreateJobRunLogSaver = $sqlCreateJobRunLogSaver.Replace('sqlmonitor_inventory_server', $InventoryServer)
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobRunLogSaver = $sqlCreateJobRunLogSaver.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$RunLogSaverJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobRunLogSaver -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 5 minutes"
        $timeIntervalMinutes = "00:05"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobRunLogSaver -EnableException
    }
}


# 22__CreateJobRunTempDbSaver
$stepName = '22__CreateJobRunTempDbSaver'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Run-TempDbSaver'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$RunTempDbSaverJobFilePath = '$RunTempDbSaverJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobRunTempDbSaver = [System.IO.File]::ReadAllText($RunTempDbSaverJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCreateJobRunTempDbSaver = $sqlCreateJobRunTempDbSaver.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlCreateJobRunTempDbSaver)
        $sqlCreateJobRunTempDbSaver = $flagsPattern.Replace($sqlCreateJobRunTempDbSaver, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlCreateJobRunTempDbSaver)) {
            #$Matches = $cmdPattern.Match($sqlCreateJobRunTempDbSaver)
            #$Matches.Groups
            $sqlCreateJobRunTempDbSaver = $cmdPattern.Replace($sqlCreateJobRunTempDbSaver, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCreateJobRunTempDbSaver = $sqlCreateJobRunTempDbSaver.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobRunTempDbSaver = $sqlCreateJobRunTempDbSaver.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCreateJobRunTempDbSaver = $sqlCreateJobRunTempDbSaver.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    $sqlCreateJobRunTempDbSaver = $sqlCreateJobRunTempDbSaver.Replace('sqlmonitor_inventory_server', $InventoryServer)
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobRunTempDbSaver = $sqlCreateJobRunTempDbSaver.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$RunTempDbSaverJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobRunTempDbSaver -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 5 minutes"
        $timeIntervalMinutes = "00:05"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobRunTempDbSaver -EnableException
    }
}


# 23__CreateJobRunWhoIsActive
$stepName = '23__CreateJobRunWhoIsActive'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Run-WhoIsActive'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$RunWhoIsActiveJobFilePath = '$RunWhoIsActiveJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlRunWhoIsActive = [System.IO.File]::ReadAllText($RunWhoIsActiveJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlRunWhoIsActive = $sqlRunWhoIsActive.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlRunWhoIsActive)
        $sqlRunWhoIsActive = $flagsPattern.Replace($sqlRunWhoIsActive, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlRunWhoIsActive)) {
            #$Matches = $cmdPattern.Match($sqlRunWhoIsActive)
            #$Matches.Groups
            $sqlRunWhoIsActive = $cmdPattern.Replace($sqlRunWhoIsActive, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlRunWhoIsActive = $sqlRunWhoIsActive.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlRunWhoIsActive = $sqlRunWhoIsActive.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlRunWhoIsActive = $sqlRunWhoIsActive.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    $sqlRunWhoIsActive = $sqlRunWhoIsActive.Replace('sqlmonitor_inventory_server', $InventoryServer)
    if($isExpressEdition) {
        $sqlRunWhoIsActive = $sqlRunWhoIsActive.Replace('@retention_day = 7,', "@retention_day = 2,")
    }
    
    if($jobNameNew -ne $jobName) {
        $sqlRunWhoIsActive = $sqlRunWhoIsActive.Replace($jobName, $jobNameNew)
    }
    
    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$RunWhoIsActiveJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlRunWhoIsActive -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 2 minutes"
        $timeIntervalMinutes = "00:02"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlRunWhoIsActive -EnableException
    }
}


# 24__CreateJobRunBlitz
$stepName = '24__CreateJobRunBlitz'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Run-Blitz'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$RunBlitzJobFilePath = '$RunBlitzJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlRunBlitzJob = [System.IO.File]::ReadAllText($RunBlitzJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlRunBlitzJob = $sqlRunBlitzJob.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlRunBlitzJob)
        $sqlRunBlitzJob = $flagsPattern.Replace($sqlRunBlitzJob, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlRunBlitzJob)) {
            #$Matches = $cmdPattern.Match($sqlRunBlitzJob)
            #$Matches.Groups
            $sqlRunBlitzJob = $cmdPattern.Replace($sqlRunBlitzJob, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlRunBlitzJob = $sqlRunBlitzJob.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlRunBlitzJob = $sqlRunBlitzJob.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlRunBlitzJob = $sqlRunBlitzJob.Replace("''DBA''", "''$DbaDatabase''" )
    $sqlRunBlitzJob = $sqlRunBlitzJob.Replace("'DBA'", "'$DbaDatabase'" )
    $sqlRunBlitzJob = $sqlRunBlitzJob.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlRunBlitzJob = $sqlRunBlitzJob.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$RunBlitzJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlRunBlitzJob -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily at 10:00 pm"
        $timeIntervalMinutes = "00:00"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                #$durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')
                $durationString = '22:00'

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                #$timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlRunBlitzJob -EnableException
    }
}


# 24__CreateJobRunBlitzIndex
$stepName = '24__CreateJobRunBlitzIndex'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Run-BlitzIndex'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$RunBlitzIndexJobFilePath = '$RunBlitzIndexJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlRunBlitzIndexJob = [System.IO.File]::ReadAllText($RunBlitzIndexJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlRunBlitzIndexJob = $sqlRunBlitzIndexJob.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlRunBlitzIndexJob)
        $sqlRunBlitzIndexJob = $flagsPattern.Replace($sqlRunBlitzIndexJob, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlRunBlitzIndexJob)) {
            #$Matches = $cmdPattern.Match($sqlRunBlitzIndexJob)
            #$Matches.Groups
            $sqlRunBlitzIndexJob = $cmdPattern.Replace($sqlRunBlitzIndexJob, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlRunBlitzIndexJob = $sqlRunBlitzIndexJob.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlRunBlitzIndexJob = $sqlRunBlitzIndexJob.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlRunBlitzIndexJob = $sqlRunBlitzIndexJob.Replace("''DBA''", "''$DbaDatabase''" )
    $sqlRunBlitzIndexJob = $sqlRunBlitzIndexJob.Replace("'DBA'", "'$DbaDatabase'" )
    $sqlRunBlitzIndexJob = $sqlRunBlitzIndexJob.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlRunBlitzIndexJob = $sqlRunBlitzIndexJob.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$RunBlitzIndexJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlRunBlitzIndexJob -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily at 7:00 pm"
        $timeIntervalMinutes = "00:00"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                #$durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')
                $durationString = '19:00'

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                #$timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlRunBlitzIndexJob -EnableException
    }
}


# 26__CreateJobRunBlitzIndexWeekly
$stepName = '26__CreateJobRunBlitzIndexWeekly'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Run-BlitzIndex - Weekly'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$RunBlitzIndexWeeklyJobFilePath = '$RunBlitzIndexWeeklyJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlRunBlitzIndexWeeklyJob = [System.IO.File]::ReadAllText($RunBlitzIndexWeeklyJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlRunBlitzIndexWeeklyJob = $sqlRunBlitzIndexWeeklyJob.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlRunBlitzIndexWeeklyJob)
        $sqlRunBlitzIndexWeeklyJob = $flagsPattern.Replace($sqlRunBlitzIndexWeeklyJob, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlRunBlitzIndexWeeklyJob)) {
            #$Matches = $cmdPattern.Match($sqlRunBlitzIndexWeeklyJob)
            #$Matches.Groups
            $sqlRunBlitzIndexWeeklyJob = $cmdPattern.Replace($sqlRunBlitzIndexWeeklyJob, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlRunBlitzIndexWeeklyJob = $sqlRunBlitzIndexWeeklyJob.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlRunBlitzIndexWeeklyJob = $sqlRunBlitzIndexWeeklyJob.Replace('-S "localhost"', "-S `"$sqlInstanceOnJobStep`"")
    $sqlRunBlitzIndexWeeklyJob = $sqlRunBlitzIndexWeeklyJob.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlRunBlitzIndexWeeklyJob = $sqlRunBlitzIndexWeeklyJob.Replace('-d "DBA"', "-d `"$DbaDatabase`"")
    $sqlRunBlitzIndexWeeklyJob = $sqlRunBlitzIndexWeeklyJob.Replace("''DBA''", "''$DbaDatabase''" )
    $sqlRunBlitzIndexWeeklyJob = $sqlRunBlitzIndexWeeklyJob.Replace("'DBA'", "'$DbaDatabase'" )
    $sqlRunBlitzIndexWeeklyJob = $sqlRunBlitzIndexWeeklyJob.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlRunBlitzIndexWeeklyJob = $sqlRunBlitzIndexWeeklyJob.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$RunBlitzIndexWeeklyJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlRunBlitzIndexWeeklyJob -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily at 8:00 PM"
        $timeIntervalMinutes = "00:00"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                #$durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')
                $durationString = '20:00'

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek Friday -At $durationString
                #$timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        if($verbose -or $debug) {
            $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlRunBlitzIndexWeeklyJob -EnableException -MessagesToOutput | Write-Verbose
        }
        else {
            $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlRunBlitzIndexWeeklyJob -EnableException
        }
    }
}


# 27__CreateJobCollectMemoryClerks
$stepName = '27__CreateJobCollectMemoryClerks'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Collect-MemoryClerks'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CollectMemoryClerksJobFilePath = '$CollectMemoryClerksJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobCollectMemoryClerks = [System.IO.File]::ReadAllText($CollectMemoryClerksJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCreateJobCollectMemoryClerks = $sqlCreateJobCollectMemoryClerks.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlCreateJobCollectMemoryClerks)
        $sqlCreateJobCollectMemoryClerks = $flagsPattern.Replace($sqlCreateJobCollectMemoryClerks, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlCreateJobCollectMemoryClerks)) {
            #$Matches = $cmdPattern.Match($sqlCreateJobCollectMemoryClerks)
            #$Matches.Groups
            $sqlCreateJobCollectMemoryClerks = $cmdPattern.Replace($sqlCreateJobCollectMemoryClerks, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCreateJobCollectMemoryClerks = $sqlCreateJobCollectMemoryClerks.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobCollectMemoryClerks = $sqlCreateJobCollectMemoryClerks.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCreateJobCollectMemoryClerks = $sqlCreateJobCollectMemoryClerks.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    $sqlCreateJobCollectMemoryClerks = $sqlCreateJobCollectMemoryClerks.Replace('sqlmonitor_inventory_server', $InventoryServer)
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobCollectMemoryClerks = $sqlCreateJobCollectMemoryClerks.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CollectMemoryClerksJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobCollectMemoryClerks -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 2 minutes"
        $timeIntervalMinutes = "00:02"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobCollectMemoryClerks -EnableException
    }
}



# 28__CreateJobCollectPrivilegedInfo
$stepName = '28__CreateJobCollectPrivilegedInfo'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Collect-PrivilegedInfo'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CollectPrivilegedInfoJobFilePath = '$CollectPrivilegedInfoJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobCollectPrivilegedInfo = [System.IO.File]::ReadAllText($CollectPrivilegedInfoJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCreateJobCollectPrivilegedInfo = $sqlCreateJobCollectPrivilegedInfo.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlCreateJobCollectPrivilegedInfo)
        $sqlCreateJobCollectPrivilegedInfo = $flagsPattern.Replace($sqlCreateJobCollectPrivilegedInfo, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlCreateJobCollectPrivilegedInfo)) {
            #$Matches = $cmdPattern.Match($sqlCreateJobCollectPrivilegedInfo)
            #$Matches.Groups
            $sqlCreateJobCollectPrivilegedInfo = $cmdPattern.Replace($sqlCreateJobCollectPrivilegedInfo, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCreateJobCollectPrivilegedInfo = $sqlCreateJobCollectPrivilegedInfo.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobCollectPrivilegedInfo = $sqlCreateJobCollectPrivilegedInfo.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCreateJobCollectPrivilegedInfo = $sqlCreateJobCollectPrivilegedInfo.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    $sqlCreateJobCollectPrivilegedInfo = $sqlCreateJobCollectPrivilegedInfo.Replace('sqlmonitor_inventory_server', $InventoryServer)
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobCollectPrivilegedInfo = $sqlCreateJobCollectPrivilegedInfo.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CollectPrivilegedInfoJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobCollectPrivilegedInfo -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 5 minutes"
        $timeIntervalMinutes = "00:05"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobCollectPrivilegedInfo -EnableException
    }
}


# 29__CreateJobCollectAgHealthState
$stepName = '29__CreateJobCollectAgHealthState'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Collect-AgHealthState'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CollectAgHealthStateJobFilePath = '$CollectAgHealthStateJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobCollectAgHealthState = [System.IO.File]::ReadAllText($CollectAgHealthStateJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCreateJobCollectAgHealthState = $sqlCreateJobCollectAgHealthState.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlCreateJobCollectAgHealthState)
        $sqlCreateJobCollectAgHealthState = $flagsPattern.Replace($sqlCreateJobCollectAgHealthState, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlCreateJobCollectAgHealthState)) {
            #$Matches = $cmdPattern.Match($sqlCreateJobCollectAgHealthState)
            #$Matches.Groups
            $sqlCreateJobCollectAgHealthState = $cmdPattern.Replace($sqlCreateJobCollectAgHealthState, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCreateJobCollectAgHealthState = $sqlCreateJobCollectAgHealthState.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobCollectAgHealthState = $sqlCreateJobCollectAgHealthState.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCreateJobCollectAgHealthState = $sqlCreateJobCollectAgHealthState.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    $sqlCreateJobCollectAgHealthState = $sqlCreateJobCollectAgHealthState.Replace('sqlmonitor_inventory_server', $InventoryServer)
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobCollectAgHealthState = $sqlCreateJobCollectAgHealthState.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CollectPrivilegedInfoJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobCollectPrivilegedInfo -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 5 minutes"
        $timeIntervalMinutes = "00:05"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobCollectAgHealthState -EnableException
    }
}



# 30__CreateJobCheckSQLAgentJobs
$stepName = '30__CreateJobCheckSQLAgentJobs'
if($stepName -in $Steps2Execute -and $isExpressEdition -eq $false) 
{
    $jobName = '(dba) Check-SQLAgentJobs'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CheckSQLAgentJobsJobFilePath = '$CheckSQLAgentJobsJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobCheckSQLAgentJobs = [System.IO.File]::ReadAllText($CheckSQLAgentJobsJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCreateJobCheckSQLAgentJobs = $sqlCreateJobCheckSQLAgentJobs.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlCreateJobCheckSQLAgentJobs)
        $sqlCreateJobCheckSQLAgentJobs = $flagsPattern.Replace($sqlCreateJobCheckSQLAgentJobs, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlCreateJobCheckSQLAgentJobs)) {
            #$Matches = $cmdPattern.Match($sqlCreateJobCheckSQLAgentJobs)
            #$Matches.Groups
            $sqlCreateJobCheckSQLAgentJobs = $cmdPattern.Replace($sqlCreateJobCheckSQLAgentJobs, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCreateJobCheckSQLAgentJobs = $sqlCreateJobCheckSQLAgentJobs.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobCheckSQLAgentJobs = $sqlCreateJobCheckSQLAgentJobs.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCreateJobCheckSQLAgentJobs = $sqlCreateJobCheckSQLAgentJobs.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    $sqlCreateJobCheckSQLAgentJobs = $sqlCreateJobCheckSQLAgentJobs.Replace('sqlmonitor_inventory_server', $InventoryServer)
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobCheckSQLAgentJobs = $sqlCreateJobCheckSQLAgentJobs.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CheckSQLAgentJobsJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlCreateJobCheckSQLAgentJobs -match "@command=N'sqlcmd (?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
            $jobArguments = $jobArguments.Replace("''","'").Replace(";;",";")
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 5 minutes"
        $timeIntervalMinutes = "00:05"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$jobArguments -o `"$logsPath\$jobName.txt`""
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobCheckSQLAgentJobs -EnableException
    }
}
else {
    if($isExpressEdition) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Step [$stepName] is skipped as not valid for SQLServer Express Edition."
    }
}



# 31__CreateJobCaptureAlertMessages
$stepName = '31__CreateJobCaptureAlertMessages'
if($stepName -in $Steps2Execute) 
{
    $jobName = '(dba) Capture-AlertMessages'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CaptureAlertMessagesJobFilePath = '$CaptureAlertMessagesJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline" 
    
    if($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort -and $stepName -ne '31__CreateJobCaptureAlertMessages') {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCreateJobCaptureAlertMessages = [System.IO.File]::ReadAllText($CaptureAlertMessagesJobFilePath)
    $sqlCreateJobCaptureAlertMessages = $sqlCreateJobCaptureAlertMessages.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCreateJobCaptureAlertMessages = $sqlCreateJobCaptureAlertMessages.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCreateJobCaptureAlertMessages = $sqlCreateJobCaptureAlertMessages.Replace("@database_name=N'DBA'", "@database_name=N'$DbaDatabase'")
    $sqlCreateJobCaptureAlertMessages = $sqlCreateJobCaptureAlertMessages.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    $sqlCreateJobCaptureAlertMessages = $sqlCreateJobCaptureAlertMessages.Replace('sqlmonitor_inventory_server', $InventoryServer)
    if($jobNameNew -ne $jobName) {
        $sqlCreateJobCaptureAlertMessages = $sqlCreateJobCaptureAlertMessages.Replace($jobName, $jobNameNew)
    }

    # If Express edition, and Task scheduler jobs are required
    if( $isExpressEdition )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "SQL Agent Job [$jobName] can not be created on Express Edition." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query $sqlCreateJobCaptureAlertMessages -EnableException
    }
}


# 32__CreateSQLAgentAlerts
$stepName = '32__CreateSQLAgentAlerts'
if($stepName -in $Steps2Execute -and $IsNonPartitioned -eq $false) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."

    if($isExpressEdition) {
        "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Skipping SQLAgent Alerts creation for SQLExpress edition.."    
    }
    else {
        $sqlCreateAgentAlerts = "exec dbo.usp_create_agent_alerts; "
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlCreateAgentAlerts -EnableException
        "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "SQLAgent Alerts created using dbo.usp_create_agent_alerts.."    
    }
}


# 33__CreateJobUpdateSqlServerVersions
$stepName = '33__CreateJobUpdateSqlServerVersions'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Update-SqlServerVersions'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$UpdateSqlServerVersionsJobFilePath = '$UpdateSqlServerVersionsJobFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobName] on [$SqlInstanceToBaseline].."
    $sqlUpdateSqlServerVersions = [System.IO.File]::ReadAllText($UpdateSqlServerVersionsJobFilePath).Replace('-SqlInstance localhost', "-SqlInstance ''$SqlInstanceToBaselineWithOutPort''")

    if($RemoteSQLMonitorPath -ne 'C:\SQLMonitor') {
        $sqlUpdateSqlServerVersions = $sqlUpdateSqlServerVersions.Replace('C:\SQLMonitor', $RemoteSQLMonitorPath)
    }

    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$UpdateSqlServerVersionsJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlUpdateSqlServerVersions -match "@command=N'powershell.exe(?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 1 minutes"
        $timeIntervalMinutes = "00:00"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'powershell' -Argument "$jobArguments *> '$logsPath\$jobName.txt'"
                $timeToDoStuff = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek Wednesday, Friday  -At $durationString
                #$timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($DropCreatePowerShellJobs) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        if($DropCreatePowerShellJobs) {
            $tsqlSSMSValidation = "and APP_NAME() = 'Microsoft SQL Server Management Studio - Query'"
            $sqlUpdateSqlServerVersions = $sqlUpdateSqlServerVersions.Replace($tsqlSSMSValidation, "--$tsqlSSMSValidation")
        }
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlUpdateSqlServerVersions -EnableException

        if($requireProxy) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update job [$jobName] to run under proxy [$credentialName].."
            $sqlUpdateJob = "EXEC msdb.dbo.sp_update_jobstep @job_name=N'$jobName', @step_id=1 ,@proxy_name=N'$credentialName';"
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlUpdateJob -EnableException
        }
        $sqlStartJob = "EXEC msdb.dbo.sp_start_job @job_name=N'$jobName';"
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlStartJob -EnableException
    }
}


# 34__CreateJobCheckInstanceAvailability
$stepName = '34__CreateJobCheckInstanceAvailability'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Check-InstanceAvailability'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CheckInstanceAvailabilityJobFilePath = '$CheckInstanceAvailabilityJobFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobName] on [$SqlInstanceToBaseline].."
    $sqlGetInstanceAvailability = [System.IO.File]::ReadAllText($CheckInstanceAvailabilityJobFilePath)
    $sqlGetInstanceAvailability = $sqlGetInstanceAvailability.Replace('-InventoryServer localhost', "-InventoryServer `"$SqlInstanceToBaselineWithOutPort`"")
    $sqlGetInstanceAvailability = $sqlGetInstanceAvailability.Replace('-InventoryDatabase DBA', "-InventoryDatabase `"$InventoryDatabase`"")

    if($RemoteSQLMonitorPath -ne 'C:\SQLMonitor') {
        $sqlGetInstanceAvailability = $sqlGetInstanceAvailability.Replace('C:\SQLMonitor', $RemoteSQLMonitorPath)
    }
    
    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CheckInstanceAvailabilityJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlGetInstanceAvailability -match "@command=N'powershell.exe(?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
        }

        $jobDescription = "Run Job [$jobName] daily every 2 minutes"
        $timeIntervalMinutes = "00:02"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'powershell' -Argument "$jobArguments *> '$logsPath\$jobName.txt'"
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($DropCreatePowerShellJobs) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        if($DropCreatePowerShellJobs) {
            $tsqlSSMSValidation = "and APP_NAME() = 'Microsoft SQL Server Management Studio - Query'"
            $sqlGetInstanceAvailability = $sqlGetInstanceAvailability.Replace($tsqlSSMSValidation, "--$tsqlSSMSValidation")
        }
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetInstanceAvailability -EnableException

        if($requireProxy) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update job [$jobName] to run under proxy [$credentialName].."
            $sqlUpdateJob = "EXEC msdb.dbo.sp_update_jobstep @job_name=N'$jobName', @step_id=1 ,@proxy_name=N'$credentialName';"
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlUpdateJob -EnableException
        }
        $sqlStartJob = "EXEC msdb.dbo.sp_start_job @job_name=N'$jobName';"
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlStartJob -EnableException
    }
}


# 35__CreateJobGetAllServerStableInfo
$stepName = '35__CreateJobGetAllServerStableInfo'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerStableInfo'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerStableInfoJobFilePath = '$GetAllServerStableInfoJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerStableInfoJobFileText = [System.IO.File]::ReadAllText($GetAllServerStableInfoJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerStableInfoJobFileText = $sqlGetAllServerStableInfoJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerStableInfoJobFileText)
        $sqlGetAllServerStableInfoJobFileText = $flagsPattern.Replace($sqlGetAllServerStableInfoJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerStableInfoJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerStableInfoJobFileText)
            #$Matches.Groups
            $sqlGetAllServerStableInfoJobFileText = $cmdPattern.Replace($sqlGetAllServerStableInfoJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerStableInfoJobFileText = $sqlGetAllServerStableInfoJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerStableInfoJobFileText = $sqlGetAllServerStableInfoJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerStableInfoJobFileText = $sqlGetAllServerStableInfoJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerStableInfoJobFileText = $sqlGetAllServerStableInfoJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerStableInfoJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 10 minutes"
        $timeIntervalMinutes = "00:10"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerStableInfoJobFileText -EnableException
    }
}


# 36__CreateJobGetAllServerVolatileInfo
$stepName = '36__CreateJobGetAllServerVolatileInfo'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerVolatileInfo'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerVolatileInfoJobFilePath = '$GetAllServerVolatileInfoJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerVolatileInfoJobFileText = [System.IO.File]::ReadAllText($GetAllServerVolatileInfoJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerVolatileInfoJobFileText = $sqlGetAllServerVolatileInfoJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerVolatileInfoJobFileText)
        $sqlGetAllServerVolatileInfoJobFileText = $flagsPattern.Replace($sqlGetAllServerVolatileInfoJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerVolatileInfoJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerVolatileInfoJobFileText)
            #$Matches.Groups
            $sqlGetAllServerVolatileInfoJobFileText = $cmdPattern.Replace($sqlGetAllServerVolatileInfoJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerVolatileInfoJobFileText = $sqlGetAllServerVolatileInfoJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerVolatileInfoJobFileText = $sqlGetAllServerVolatileInfoJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerVolatileInfoJobFileText = $sqlGetAllServerVolatileInfoJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerVolatileInfoJobFileText = $sqlGetAllServerVolatileInfoJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerVolatileInfoJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 20 seconds"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerVolatileInfoJobFileText -EnableException
    }
}


# 37__CreateJobGetAllServerCollectionLatencyInfo
$stepName = '37__CreateJobGetAllServerCollectionLatencyInfo'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerCollectionLatencyInfo'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerCollectionLatencyInfoJobFilePath = '$GetAllServerCollectionLatencyInfoJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerCollectionLatencyInfoJobFileText = [System.IO.File]::ReadAllText($GetAllServerCollectionLatencyInfoJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerCollectionLatencyInfoJobFileText = $sqlGetAllServerCollectionLatencyInfoJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerCollectionLatencyInfoJobFileText)
        $sqlGetAllServerCollectionLatencyInfoJobFileText = $flagsPattern.Replace($sqlGetAllServerCollectionLatencyInfoJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerCollectionLatencyInfoJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerCollectionLatencyInfoJobFileText)
            #$Matches.Groups
            $sqlGetAllServerCollectionLatencyInfoJobFileText = $cmdPattern.Replace($sqlGetAllServerCollectionLatencyInfoJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerCollectionLatencyInfoJobFileText = $sqlGetAllServerCollectionLatencyInfoJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerCollectionLatencyInfoJobFileText = $sqlGetAllServerCollectionLatencyInfoJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerCollectionLatencyInfoJobFileText = $sqlGetAllServerCollectionLatencyInfoJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerCollectionLatencyInfoJobFileText = $sqlGetAllServerCollectionLatencyInfoJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerCollectionLatencyInfoJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 20 seconds"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerCollectionLatencyInfoJobFileText -EnableException
    }
}


# 38__CreateJobGetAllServerSqlAgentJobs
$stepName = '38__CreateJobGetAllServerSqlAgentJobs'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerSqlAgentJobs'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerSqlAgentJobsJobFilePath = '$GetAllServerSqlAgentJobsJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerSqlAgentJobsJobFileText = [System.IO.File]::ReadAllText($GetAllServerSqlAgentJobsJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerSqlAgentJobsJobFileText = $sqlGetAllServerSqlAgentJobsJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerSqlAgentJobsJobFileText)
        $sqlGetAllServerSqlAgentJobsJobFileText = $flagsPattern.Replace($sqlGetAllServerSqlAgentJobsJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerSqlAgentJobsJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerSqlAgentJobsJobFileText)
            #$Matches.Groups
            $sqlGetAllServerSqlAgentJobsJobFileText = $cmdPattern.Replace($sqlGetAllServerSqlAgentJobsJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerSqlAgentJobsJobFileText = $sqlGetAllServerSqlAgentJobsJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerSqlAgentJobsJobFileText = $sqlGetAllServerSqlAgentJobsJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerSqlAgentJobsJobFileText = $sqlGetAllServerSqlAgentJobsJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerSqlAgentJobsJobFileText = $sqlGetAllServerSqlAgentJobsJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerSqlAgentJobsJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 20 seconds"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerSqlAgentJobsJobFileText -EnableException
    }
}


# 39__CreateJobGetAllServerDiskSpace
$stepName = '39__CreateJobGetAllServerDiskSpace'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerDiskSpace'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerDiskSpaceJobFilePath = '$GetAllServerDiskSpaceJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerDiskSpaceJobFileText = [System.IO.File]::ReadAllText($GetAllServerDiskSpaceJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerDiskSpaceJobFileText = $sqlGetAllServerDiskSpaceJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerDiskSpaceJobFileText)
        $sqlGetAllServerDiskSpaceJobFileText = $flagsPattern.Replace($sqlGetAllServerDiskSpaceJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerDiskSpaceJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerDiskSpaceJobFileText)
            #$Matches.Groups
            $sqlGetAllServerDiskSpaceJobFileText = $cmdPattern.Replace($sqlGetAllServerDiskSpaceJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerDiskSpaceJobFileText = $sqlGetAllServerDiskSpaceJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerDiskSpaceJobFileText = $sqlGetAllServerDiskSpaceJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerDiskSpaceJobFileText = $sqlGetAllServerDiskSpaceJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerDiskSpaceJobFileText = $sqlGetAllServerDiskSpaceJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerDiskSpaceJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 20 seconds"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerDiskSpaceJobFileText -EnableException
    }
}


# 40__CreateJobGetAllServerLogSpaceConsumers
$stepName = '40__CreateJobGetAllServerLogSpaceConsumers'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerLogSpaceConsumers'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerLogSpaceConsumersJobFilePath = '$GetAllServerLogSpaceConsumersJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerLogSpaceConsumersJobFileText = [System.IO.File]::ReadAllText($GetAllServerLogSpaceConsumersJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerLogSpaceConsumersJobFileText = $sqlGetAllServerLogSpaceConsumersJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerLogSpaceConsumersJobFileText)
        $sqlGetAllServerLogSpaceConsumersJobFileText = $flagsPattern.Replace($sqlGetAllServerLogSpaceConsumersJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerLogSpaceConsumersJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerLogSpaceConsumersJobFileText)
            #$Matches.Groups
            $sqlGetAllServerLogSpaceConsumersJobFileText = $cmdPattern.Replace($sqlGetAllServerLogSpaceConsumersJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerLogSpaceConsumersJobFileText = $sqlGetAllServerLogSpaceConsumersJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerLogSpaceConsumersJobFileText = $sqlGetAllServerLogSpaceConsumersJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerLogSpaceConsumersJobFileText = $sqlGetAllServerLogSpaceConsumersJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerLogSpaceConsumersJobFileText = $sqlGetAllServerLogSpaceConsumersJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerLogSpaceConsumersJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 20 seconds"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerLogSpaceConsumersJobFileText -EnableException
    }
}



# 41__CreateJobGetAllServerTempdbSpaceUsage
$stepName = '41__CreateJobGetAllServerTempdbSpaceUsage'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerTempdbSpaceUsage'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerTempdbSpaceUsageJobFilePath = '$GetAllServerTempdbSpaceUsageJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerTempdbSpaceUsageJobFileText = [System.IO.File]::ReadAllText($GetAllServerTempdbSpaceUsageJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerTempdbSpaceUsageJobFileText = $sqlGetAllServerTempdbSpaceUsageJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerTempdbSpaceUsageJobFileText)
        $sqlGetAllServerTempdbSpaceUsageJobFileText = $flagsPattern.Replace($sqlGetAllServerTempdbSpaceUsageJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerTempdbSpaceUsageJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerTempdbSpaceUsageJobFileText)
            #$Matches.Groups
            $sqlGetAllServerTempdbSpaceUsageJobFileText = $cmdPattern.Replace($sqlGetAllServerTempdbSpaceUsageJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerTempdbSpaceUsageJobFileText = $sqlGetAllServerTempdbSpaceUsageJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerTempdbSpaceUsageJobFileText = $sqlGetAllServerTempdbSpaceUsageJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerTempdbSpaceUsageJobFileText = $sqlGetAllServerTempdbSpaceUsageJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerTempdbSpaceUsageJobFileText = $sqlGetAllServerTempdbSpaceUsageJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerTempdbSpaceUsageJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 20 seconds"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerTempdbSpaceUsageJobFileText -EnableException
    }
}


# 42__CreateJobGetAllServerAgHealthState
$stepName = '42__CreateJobGetAllServerAgHealthState'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerAgHealthState'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerAgHealthStateJobFilePath = '$GetAllServerAgHealthStateJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerAgHealthStateJobFileText = [System.IO.File]::ReadAllText($GetAllServerAgHealthStateJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerAgHealthStateJobFileText = $sqlGetAllServerAgHealthStateJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerAgHealthStateJobFileText)
        $sqlGetAllServerAgHealthStateJobFileText = $flagsPattern.Replace($sqlGetAllServerAgHealthStateJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerAgHealthStateJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerAgHealthStateJobFileText)
            #$Matches.Groups
            $sqlGetAllServerAgHealthStateJobFileText = $cmdPattern.Replace($sqlGetAllServerAgHealthStateJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerAgHealthStateJobFileText = $sqlGetAllServerAgHealthStateJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerAgHealthStateJobFileText = $sqlGetAllServerAgHealthStateJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerAgHealthStateJobFileText = $sqlGetAllServerAgHealthStateJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerAgHealthStateJobFileText = $sqlGetAllServerAgHealthStateJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerAgHealthStateJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 20 seconds"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerAgHealthStateJobFileText -EnableException
    }
}


# 43__CreateJobGetAllServerServices
$stepName = '43__CreateJobGetAllServerServices'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerServices'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerServicesJobFilePath = '$GetAllServerServicesJobFilePath'"

    # Append HostName if Job Server is different
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerServicesJobFileText = [System.IO.File]::ReadAllText($GetAllServerServicesJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerServicesJobFileText = $sqlGetAllServerServicesJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerServicesJobFileText)
        $sqlGetAllServerServicesJobFileText = $flagsPattern.Replace($sqlGetAllServerServicesJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerServicesJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerServicesJobFileText)
            #$Matches.Groups
            $sqlGetAllServerServicesJobFileText = $cmdPattern.Replace($sqlGetAllServerServicesJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerServicesJobFileText = $sqlGetAllServerServicesJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerServicesJobFileText = $sqlGetAllServerServicesJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerServicesJobFileText = $sqlGetAllServerServicesJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerServicesJobFileText = $sqlGetAllServerServicesJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerServicesJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 20 seconds"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerServicesJobFileText -EnableException
    }
}


# 44__CreateJobGetAllServerBackups
$stepName = '44__CreateJobGetAllServerBackups'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerBackups'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerBackupsJobFilePath = '$GetAllServerBackupsJobFilePath'"

    # Append HostName if Job Server is different
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerBackupsJobFileText = [System.IO.File]::ReadAllText($GetAllServerBackupsJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerBackupsJobFileText = $sqlGetAllServerBackupsJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerBackupsJobFileText)
        $sqlGetAllServerBackupsJobFileText = $flagsPattern.Replace($sqlGetAllServerBackupsJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerBackupsJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerBackupsJobFileText)
            #$Matches.Groups
            $sqlGetAllServerBackupsJobFileText = $cmdPattern.Replace($sqlGetAllServerBackupsJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerBackupsJobFileText = $sqlGetAllServerBackupsJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerBackupsJobFileText = $sqlGetAllServerBackupsJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerBackupsJobFileText = $sqlGetAllServerBackupsJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerBackupsJobFileText = $sqlGetAllServerBackupsJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerBackupsJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 20 seconds"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerBackupsJobFileText -EnableException
    }
}


# 45__CreateJobGetAllServerAlertHistory
$stepName = '45__CreateJobGetAllServerAlertHistory'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerAlertHistory'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerAlertHistoryJobFilePath = '$GetAllServerAlertHistoryJobFilePath'"

    # Append HostName if Job Server is different
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerAlertHistoryJobFileText = [System.IO.File]::ReadAllText($GetAllServerAlertHistoryJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerAlertHistoryJobFileText = $sqlGetAllServerAlertHistoryJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerAlertHistoryJobFileText)
        $sqlGetAllServerAlertHistoryJobFileText = $flagsPattern.Replace($sqlGetAllServerAlertHistoryJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerAlertHistoryJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerAlertHistoryJobFileText)
            #$Matches.Groups
            $sqlGetAllServerAlertHistoryJobFileText = $cmdPattern.Replace($sqlGetAllServerAlertHistoryJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerAlertHistoryJobFileText = $sqlGetAllServerAlertHistoryJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerAlertHistoryJobFileText = $sqlGetAllServerAlertHistoryJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerAlertHistoryJobFileText = $sqlGetAllServerAlertHistoryJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerAlertHistoryJobFileText = $sqlGetAllServerAlertHistoryJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerAlertHistoryJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 20 seconds"
        $timeIntervalMinutes = "00:01"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "00:10" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "00:10" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerAlertHistoryJobFileText -EnableException
    }
}


# 46__CreateJobGetAllServerDashboardMail
$stepName = '46__CreateJobGetAllServerDashboardMail'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Get-AllServerDashboardMail'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$GetAllServerDashboardMailJobFilePath = '$GetAllServerDashboardMailJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlGetAllServerDashboardMailJobFileText = [System.IO.File]::ReadAllText($GetAllServerDashboardMailJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlGetAllServerDashboardMailJobFileText = $sqlGetAllServerDashboardMailJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerDashboardMailJobFileText)
        $sqlGetAllServerDashboardMailJobFileText = $flagsPattern.Replace($sqlGetAllServerDashboardMailJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlGetAllServerDashboardMailJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerDashboardMailJobFileText)
            #$Matches.Groups
            $sqlGetAllServerDashboardMailJobFileText = $cmdPattern.Replace($sqlGetAllServerDashboardMailJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlGetAllServerDashboardMailJobFileText = $sqlGetAllServerDashboardMailJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlGetAllServerDashboardMailJobFileText = $sqlGetAllServerDashboardMailJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlGetAllServerDashboardMailJobFileText = $sqlGetAllServerDashboardMailJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlGetAllServerDashboardMailJobFileText = $sqlGetAllServerDashboardMailJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$GetAllServerDashboardMailJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlGetAllServerDashboardMailJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 8 hours"
        $timeIntervalMinutes = "08:00"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "08:00" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "08:00" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlGetAllServerDashboardMailJobFileText -EnableException
    }
}


# 47__CreateJobStopStuckSQLMonitorJobs
$stepName = '47__CreateJobStopStuckSQLMonitorJobs'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Stop-StuckSQLMonitorJobs'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$StopStuckSQLMonitorJobsJobFilePath = '$StopStuckSQLMonitorJobsJobFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobName] on [$SqlInstanceToBaseline].."
    $sqlStopStuckSQLMonitorJobs = [System.IO.File]::ReadAllText($StopStuckSQLMonitorJobsJobFilePath)

    $sqlStopStuckSQLMonitorJobs = $sqlStopStuckSQLMonitorJobs.Replace('-InventoryServer localhost', "-InventoryServer `"$SqlInstanceToBaselineWithOutPort`"")
    $sqlStopStuckSQLMonitorJobs = $sqlStopStuckSQLMonitorJobs.Replace('-InventoryDatabase DBA', "-InventoryDatabase `"$InventoryDatabase`"")
    $sqlStopStuckSQLMonitorJobs = $sqlStopStuckSQLMonitorJobs.Replace('-CredentialManagerDatabase DBA', "-CredentialManagerDatabase `"$InventoryDatabase`"")

    if($RemoteSQLMonitorPath -ne 'C:\SQLMonitor') {
        $sqlStopStuckSQLMonitorJobs = $sqlStopStuckSQLMonitorJobs.Replace('C:\SQLMonitor', $RemoteSQLMonitorPath)
    }
    
    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$StopStuckSQLMonitorJobsJobFilePath'.."
        [String]$jobArguments = $null
        if($sqlStopStuckSQLMonitorJobs -match "@command=N'powershell.exe(?'arguments'.*)',") {
            $jobArguments = $Matches['arguments']
        }

        if([String]::IsNullOrEmpty($jobArguments)) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 1 hour"
        $timeIntervalMinutes = "01:00"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = New-ScheduledTaskAction -Execute 'powershell' -Argument "$jobArguments *> '$logsPath\$jobName.txt'"
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "01:00" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "01:00" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($DropCreatePowerShellJobs) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $DropCreatePowerShellJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error." #| Write-Host -ForegroundColor Yellow
    }
    else # If not express edition
    {
        if($DropCreatePowerShellJobs) {
            $tsqlSSMSValidation = "and APP_NAME() = 'Microsoft SQL Server Management Studio - Query'"
            $sqlStopStuckSQLMonitorJobs = $sqlStopStuckSQLMonitorJobs.Replace($tsqlSSMSValidation, "--$tsqlSSMSValidation")
        }
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlStopStuckSQLMonitorJobs -EnableException

        if($requireProxy) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update job [$jobName] to run under proxy [$credentialName].."
            $sqlUpdateJob = "EXEC msdb.dbo.sp_update_jobstep @job_name=N'$jobName', @step_id=1 ,@proxy_name=N'$credentialName';"
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlUpdateJob -EnableException
        }
        $sqlStartJob = "EXEC msdb.dbo.sp_start_job @job_name=N'$jobName';"
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlStartJob -EnableException
    }
}


# 48__CreateJobCollectLoginExpirationInfo
$stepName = '48__CreateJobCollectLoginExpirationInfo'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Collect Login Expiration Info'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$CollectLoginExpirationInfoJobFilePath = '$CollectLoginExpirationInfoJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlCollectLoginExpirationInfoJobFile = [System.IO.File]::ReadAllText($CollectLoginExpirationInfoJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlCollectLoginExpirationInfoJobFile = $sqlCollectLoginExpirationInfoJobFile.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlGetAllServerDashboardMailJobFileText)
        $sqlCollectLoginExpirationInfoJobFile = $flagsPattern.Replace($sqlCollectLoginExpirationInfoJobFile, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlCollectLoginExpirationInfoJobFile)) {
            #$Matches = $cmdPattern.Match($sqlGetAllServerDashboardMailJobFileText)
            #$Matches.Groups
            $sqlCollectLoginExpirationInfoJobFile = $cmdPattern.Replace($sqlCollectLoginExpirationInfoJobFile, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlCollectLoginExpirationInfoJobFile = $sqlCollectLoginExpirationInfoJobFile.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCollectLoginExpirationInfoJobFile = $sqlCollectLoginExpirationInfoJobFile.Replace('-S Localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCollectLoginExpirationInfoJobFile = $sqlCollectLoginExpirationInfoJobFile.Replace('-S "localhost"', "-S `"$sqlInstanceOnJobStep`"")
    $sqlCollectLoginExpirationInfoJobFile = $sqlCollectLoginExpirationInfoJobFile.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlCollectLoginExpirationInfoJobFile = $sqlCollectLoginExpirationInfoJobFile.Replace('-d "DBA"', "-d `"$DbaDatabase`"")
    $sqlCollectLoginExpirationInfoJobFile = $sqlCollectLoginExpirationInfoJobFile.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if($jobNameNew -ne $jobName) {
        $sqlCollectLoginExpirationInfoJobFile = $sqlCollectLoginExpirationInfoJobFile.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$CollectLoginExpirationInfoJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlCollectLoginExpirationInfoJobFile -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 8 hours"
        $timeIntervalMinutes = "08:00"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "08:00" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "08:00" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlCollectLoginExpirationInfoJobFile -EnableException
    }
}


# 49__CreateJobPopulateInventoryTables
$stepName = '49__CreateJobPopulateInventoryTables'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Populate Inventory Tables'
    # This job is special as it contains one step of PowerShell script & another one of TSQL Procedure

    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$PopulateInventoryTablesJobFilePath = '$PopulateInventoryTablesJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlPopulateInventoryTablesJobFileText = [System.IO.File]::ReadAllText($PopulateInventoryTablesJobFilePath)

    # For TSQL Type, extract TSQL Query if required
    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        
        [regex]$cmdPattern = "@subsystem=N'CmdExec',(?<LineSpaceCharacters>\s+)@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlPopulateInventoryTablesJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlPopulateInventoryTablesJobFileText)
            #$Matches.Groups
            $sqlPopulateInventoryTablesJobFileText = $cmdPattern.Replace($sqlPopulateInventoryTablesJobFileText, "@command=N'`${JobStepTsqlCode}', `${LineSpaceCharacters}@subsystem=N'TSQL',", 1)
        }

        [regex]$subsystemPattern = "@subsystem=N'TSQL',(?<LineSpaceCharacters>\s+)@flags=(?<flagsValue>\d+)[^,]"
        while ($subsystemPattern.IsMatch($sqlPopulateInventoryTablesJobFileText)) {
            #$Matches = $subsystemPattern.Match($sqlPopulateInventoryTablesJobFileText)
            #$Matches.Groups
            $sqlPopulateInventoryTablesJobFileText = $subsystemPattern.Replace($sqlPopulateInventoryTablesJobFileText, "@subsystem=N'TSQL', `${LineSpaceCharacters}@database_name=N'$DbaDatabase', `${LineSpaceCharacters}@flags=12", 1)
        }
    }

    # Since the job has 2 steps. One TSQL Type & Another PowerShell type
        # Update PowerShell parameters
    $sqlPopulateInventoryTablesJobFileText = $sqlPopulateInventoryTablesJobFileText.Replace('-InventoryServer localhost', "-InventoryServer `"$SqlInstanceToBaselineWithOutPort`"")
    $sqlPopulateInventoryTablesJobFileText = $sqlPopulateInventoryTablesJobFileText.Replace('-InventoryDatabase DBA', "-InventoryDatabase `"$InventoryDatabase`"")
    $sqlPopulateInventoryTablesJobFileText = $sqlPopulateInventoryTablesJobFileText.Replace('-CredentialManagerDatabase DBA', "-CredentialManagerDatabase `"$InventoryDatabase`"")
    if($sqlPopulateInventoryTablesJobFileText -ne 'C:\SQLMonitor') {
        $sqlPopulateInventoryTablesJobFileText = $sqlPopulateInventoryTablesJobFileText.Replace('C:\SQLMonitor', $RemoteSQLMonitorPath)
    }

    # Update TSQL parameters
    $sqlPopulateInventoryTablesJobFileText = $sqlPopulateInventoryTablesJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlPopulateInventoryTablesJobFileText = $sqlPopulateInventoryTablesJobFileText.Replace('-S Localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlPopulateInventoryTablesJobFileText = $sqlPopulateInventoryTablesJobFileText.Replace('-S "localhost"', "-S `"$sqlInstanceOnJobStep`"")
    $sqlPopulateInventoryTablesJobFileText = $sqlPopulateInventoryTablesJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlPopulateInventoryTablesJobFileText = $sqlPopulateInventoryTablesJobFileText.Replace('-d "DBA"', "-d `"$DbaDatabase`"")
    $sqlPopulateInventoryTablesJobFileText = $sqlPopulateInventoryTablesJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )

    if($jobNameNew -ne $jobName) {
        $sqlPopulateInventoryTablesJobFileText = $sqlPopulateInventoryTablesJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$PopulateInventoryTablesJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlPopulateInventoryTablesJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 8 hours"
        $timeIntervalMinutes = "08:00"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "08:00" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "08:00" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conInventoryServer | Invoke-DbaQuery -Database msdb -Query $sqlPopulateInventoryTablesJobFileText -EnableException
    }
}


# 50__CreateJobSendLoginExpiryEmails
$stepName = '50__CreateJobSendLoginExpiryEmails'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Send Login Expiry EMails'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$SendLoginExpiryEmailsJobFilePath = '$SendLoginExpiryEmailsJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlSendLoginExpiryEmailsJobFileText = [System.IO.File]::ReadAllText($SendLoginExpiryEmailsJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlSendLoginExpiryEmailsJobFileText = $sqlSendLoginExpiryEmailsJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlSendLoginExpiryEmailsJobFileText)
        $sqlSendLoginExpiryEmailsJobFileText = $flagsPattern.Replace($sqlSendLoginExpiryEmailsJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlSendLoginExpiryEmailsJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlSendLoginExpiryEmailsJobFileText)
            #$Matches.Groups
            $sqlSendLoginExpiryEmailsJobFileText = $cmdPattern.Replace($sqlSendLoginExpiryEmailsJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlSendLoginExpiryEmailsJobFileText = $sqlSendLoginExpiryEmailsJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlSendLoginExpiryEmailsJobFileText = $sqlSendLoginExpiryEmailsJobFileText.Replace('-S Localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlSendLoginExpiryEmailsJobFileText = $sqlSendLoginExpiryEmailsJobFileText.Replace('-S "localhost"', "-S `"$sqlInstanceOnJobStep`"")
    $sqlSendLoginExpiryEmailsJobFileText = $sqlSendLoginExpiryEmailsJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlSendLoginExpiryEmailsJobFileText = $sqlSendLoginExpiryEmailsJobFileText.Replace('-d "DBA"', "-d `"$DbaDatabase`"")
    $sqlSendLoginExpiryEmailsJobFileText = $sqlSendLoginExpiryEmailsJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if(-not $GrafanaDashboardPortal.EndsWith('/')) {
        $GrafanaDashboardPortal = "$GrafanaDashboardPortal/"
    }
    $sqlSendLoginExpiryEmailsJobFileText = $sqlSendLoginExpiryEmailsJobFileText.Replace("https://sqlmonitor.ajaydwivedi.com:3000/", "$GrafanaDashboardPortal" )

    if($jobNameNew -ne $jobName) {
        $sqlSendLoginExpiryEmailsJobFileText = $sqlSendLoginExpiryEmailsJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$SendLoginExpiryEmailsJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlSendLoginExpiryEmailsJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 8 hours"
        $timeIntervalMinutes = "08:00"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "08:00" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "08:00" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlSendLoginExpiryEmailsJobFileText -EnableException
    }
}


# 51__CreateJobComputeAllServerVolatileInfoHistoryHourly
$stepName = '51__CreateJobComputeAllServerVolatileInfoHistoryHourly'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaseline -eq $InventoryServer) 
{
    $jobName = '(dba) Compute-AllServerVolatileInfoHistoryHourly'
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$ComputeAllServerVolatileInfoHistoryHourlyJobFilePath = '$ComputeAllServerVolatileInfoHistoryHourlyJobFilePath'"

    # Append HostName if Job Server is different    
    $jobNameNew = $jobName
    #$sqlInstanceOnJobStep = "$SqlInstanceToBaselineWithOutPort"
    $sqlInstanceOnJobStep = "$SqlInstanceToBaseline"
    if ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceForTsqlJobsWithOutPort) {
        $jobNameNew = "$jobName - $SqlInstanceToBaselineWithOutPort"
        #$sqlInstanceOnJobStep = $SqlInstanceToBaseline
    }

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating job [$jobNameNew] on [$SqlInstanceForTsqlJobs].."
    $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = [System.IO.File]::ReadAllText($ComputeAllServerVolatileInfoHistoryHourlyJobFilePath)

    if ($ForceTSQLStepType4TsqlJobs -and $SqlInstanceToBaseline -eq $SqlInstanceForTsqlJobs) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Parameter ForceTSQLStepType4TsqlJobs is set to true. So creating TSQL step type job.."
        $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText.Replace("@subsystem=N'CmdExec'", "@subsystem=N'TSQL'")

        [regex]$flagsPattern = "@flags=(?<flagsValue>\d+)[^,]"
        #$flagsPattern.IsMatch($sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText)
        $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = $flagsPattern.Replace($sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText, "@flags=12, @database_name=N'$DbaDatabase'")

        [regex]$cmdPattern = "@command=N'(?<CommandText>sqlcmd.* -Q `"(?<JobStepTsqlCode>.*)`")',"
        while ($cmdPattern.IsMatch($sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText)) {
            #$Matches = $cmdPattern.Match($sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText)
            #$Matches.Groups
            $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = $cmdPattern.Replace($sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText, "@command=N'`${JobStepTsqlCode}',", 1)
        }
    }

    $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText.Replace('-S localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText.Replace('-S Localhost', "-S `"$sqlInstanceOnJobStep`"")
    $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText.Replace('-S "localhost"', "-S `"$sqlInstanceOnJobStep`"")
    $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText.Replace('-d DBA', "-d `"$DbaDatabase`"")
    $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText.Replace('-d "DBA"', "-d `"$DbaDatabase`"")
    $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText.Replace("''dba_team@gmail.com''", "''$($DbaGroupMailId -join ';')''" )
    if(-not $GrafanaDashboardPortal.EndsWith('/')) {
        $GrafanaDashboardPortal = "$GrafanaDashboardPortal/"
    }
    $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText.Replace("https://sqlmonitor.ajaydwivedi.com:3000/", "$GrafanaDashboardPortal" )

    if($jobNameNew -ne $jobName) {
        $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText = $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText.Replace($jobName, $jobNameNew)
    }


    # If Express edition, and Task scheduler jobs are required
    if( ((-not [String]::IsNullOrEmpty($WindowsCredential)) -or ($ssnHostName -eq $env:COMPUTERNAME)) `
        -and ($isExpressEdition -or $ForceSetupOfTaskSchedulerJobs) -and ($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) )
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Extract Job command from '$ComputeAllServerVolatileInfoHistoryHourlyJobFilePath'.."
        [System.Collections.ArrayList]$jobArguments = @()
        foreach($line in $($sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText -split "`n")) 
        {
            if($line -match "@command=N'sqlcmd (?'arguments'.*)',") {
                $command = $Matches['arguments']
                $command = $command.Replace("''","'").Replace(";;",";")
                $jobArguments.Add($command) | Out-Null
            }
        }

        if($jobArguments.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
		        "Failure in extracting Job command in '$stepName'." | Write-Error
	        }
	        else {            
		        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Failure in extracting Job command in '$stepName'." | Write-Host -ForegroundColor Red
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly Resolve above error." | Write-Error
            }
        }

        $jobDescription = "Run Job [$jobName] daily every 8 hours"
        $timeIntervalMinutes = "08:00"
        $taskPath = '\DBA\'
        $parameters = @{
            Session = $ssn4PerfmonSetup
            ScriptBlock = {
                Param ($jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs)

                $currentTime = Get-Date
                $durationString = ($currentTime.AddMinutes(1)).ToString('HH:mm')

                $doStuff = @()
                $counter = 1
                foreach($command in $jobArguments) {
                    $doStuff += $(New-ScheduledTaskAction -Execute 'sqlcmd' -Argument "$command -o `"$logsPath\$jobName-$counter.txt`"")
                    $counter += 1
                }
                $timeToDoStuff = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -RandomDelay "08:00" -At $durationString
                $timeToDoStuff.Repetition = $(New-ScheduledTaskTrigger -Once -RandomDelay "08:00" -At $durationString -RepetitionDuration "23:59" -RepetitionInterval $timeIntervalMinutes).Repetition
                $settingsForTheStuff = New-ScheduledTaskSettingsSet
                $runAsUser = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $finalBuildOfTheStuff = New-ScheduledTask -Action $doStuff -Trigger $timeToDoStuff -Settings $settingsForTheStuff -Principal $runAsUser -Description $jobDescription

                $taskObj = @()
                try { $taskObj += Get-ScheduledTask -TaskName $jobName -TaskPath $taskPath -ErrorAction SilentlyContinue }
                catch { "Some Error" | Out-Null }

                $isCreated = $false
                if([String]::IsNullOrEmpty($taskObj)) {
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] created in Windows Task Scheduler."
                }
                elseif ($SkipTsqlJobs -eq $false) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Recreate Job [$jobName] in Windows Task Scheduler.."
                    $taskObj | Unregister-ScheduledTask -Confirm:$false | Out-Null
                    Register-ScheduledTask -TaskName $jobName -InputObject $finalBuildOfTheStuff -TaskPath $taskPath | Out-Null
                    $isCreated = $true
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] dropped & recreated in Windows Task Scheduler."
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [$jobName] already exists in Windows Task Scheduler."
                }

                if($isCreated) {
                    Start-ScheduledTask -TaskName $jobName -TaskPath $taskPath | Out-Null
                }
            }
            ArgumentList = $jobName, $jobDescription, $jobArguments, $timeIntervalMinutes, $taskPath, $logsPath, $SkipTsqlJobs
        }

        Invoke-Command @parameters -ErrorAction Stop
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly ensure windows Task Scheduler job [$($taskPath)$($jobName)] is running without Error."
    }
    else # If not express edition
    {
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database msdb -Query $sqlComputeAllServerVolatileInfoHistoryHourlyJobFileText -EnableException
    }
}


# 52__WhoIsActivePartition
$stepName = '52__WhoIsActivePartition'
if($stepName -in $Steps2Execute -and $IsNonPartitioned -eq $false) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$WhoIsActivePartitionFilePath = '$WhoIsActivePartitionFilePath'"
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "ALTER [dbo].[WhoIsActive] table to partitioned table on [$SqlInstanceToBaseline].."
    $sqlPartitionWhoIsActive = [System.IO.File]::ReadAllText($WhoIsActivePartitionFilePath).Replace("[DBA]", "[$DbaDatabase]")

    # If Task Scheduler Job, then execute one time manually
    if($ForceSetupOfTaskSchedulerJobs) {
        $sql2RunManually = "EXEC dbo.usp_run_WhoIsActive @recipients = '$($DbaGroupMailId -join ';')';"
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Since ForceSetupOfTaskSchedulerJobs is enabled, executing dbo.usp_run_WhoIsActive once manually.."
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sql2RunManually -EnableException
    }
    
    $whoIsActiveExists = @()
    $loopStartTime = Get-Date
    $sleepDurationSeconds = 30
    $loopTotalDurationThresholdSeconds = 300    
    
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Check for existance of table [dbo].[WhoIsActive] on [$SqlInstanceToBaseline].."
    while ($whoIsActiveExists.Count -eq 0 -and $( (New-TimeSpan $loopStartTime $(Get-Date)).TotalSeconds -le $loopTotalDurationThresholdSeconds ) )
    {
        $whoIsActiveExists += $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase `
                                    -Query "if OBJECT_ID('dbo.WhoIsActive') is not null select OBJECT_ID('dbo.WhoIsActive') as WhoIsActiveObjectID" -EnableException

        if($whoIsActiveExists.Count -eq 0) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Wait for $sleepDurationSeconds seconds as table dbo.WhoIsActive still does not exist.."
            Start-Sleep -Seconds $sleepDurationSeconds
        }
    }

    if($whoIsActiveExists.Count -eq 0) {
        if ($ReturnInlineErrorMessage) {
            "Table [dbo].[WhoIsActive] does not exist.`nKindly ensure job [(dba) Run-WhoIsActive] is running successfully." | Write-Error
        }
        else {            
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Table [dbo].[WhoIsActive] does not exist." | Write-Host -ForegroundColor Red
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly ensure job [(dba) Run-WhoIsActive] is running successfully." | Write-Host -ForegroundColor Red
        
            "STOP here, and fix above issue." | Write-Error
        }
    }
    else {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Seems table exists now. Convert [dbo].[WhoIsActive] into partitioned table.."
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlPartitionWhoIsActive -EnableException
    }
}


# 53__BlitzIndexPartition
$stepName = '53__BlitzIndexPartition'
if($stepName -in $Steps2Execute) 
{
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."

    # Create CX on dbo.BlitzIndex
    $tableName = "BlitzIndex"
    if($tableName -eq "BlitzIndex")
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$BlitzIndexPartitionFilePath = '$BlitzIndexPartitionFilePath'"
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "ALTER [dbo].[BlitzIndex] table to partitioned table on [$SqlInstanceToBaseline].."
        $sqlPartitionBlitzIndex = [System.IO.File]::ReadAllText($BlitzIndexPartitionFilePath).Replace("[DBA]", "[$DbaDatabase]")

        # If Task Scheduler Job, then execute one time manually
        if($ForceSetupOfTaskSchedulerJobs) {
            $sql2RunManually = "EXEC master.dbo.sp_BlitzIndex @GetAllDatabases = 1, @Mode = 2, @BringThePain = 1, @OutputDatabaseName = '$DbaDatabase', @OutputSchemaName = 'dbo', @OutputTableName = 'BlitzIndex';"
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Since ForceSetupOfTaskSchedulerJobs is enabled, executing 'dbo.sp_BlitzIndex @Mode = 2' once manually.."
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sql2RunManually -EnableException
        }

        # Modify content if SQL Server does not support Partitioning
        if($IsNonPartitioned) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Remove partition flag from '$BlitzIndexPartitionFileName'.."
            $sqlPartitionBlitzIndex = $sqlPartitionBlitzIndex.Replace('declare @is_partitioned bit = 1;', 'declare @is_partitioned bit = 0;')
        }
    
        $BlitzIndexExists = @()
        $loopStartTime = Get-Date
        $sleepDurationSeconds = 30
        $loopTotalDurationThresholdSeconds = $JobsExecutionWaitTimeoutMinutes*60
    
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Check for existance of table [dbo].[BlitzIndex] on [$SqlInstanceToBaseline].."
        while ($BlitzIndexExists.Count -eq 0 -and $( (New-TimeSpan $loopStartTime $(Get-Date)).TotalSeconds -le $loopTotalDurationThresholdSeconds ) )
        {
            $BlitzIndexExists += $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase `
                                        -Query "if OBJECT_ID('dbo.BlitzIndex') is not null select OBJECT_ID('dbo.BlitzIndex') as BlitzIndexObjectID" -EnableException

            if($BlitzIndexExists.Count -eq 0) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Wait for $sleepDurationSeconds seconds as table dbo.BlitzIndex still does not exist.."
                Start-Sleep -Seconds $sleepDurationSeconds
            }
        }

        if($BlitzIndexExists.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
                "Run job [(dba) Run-BlitzIndex] to create table dbo.BlitzIndex." | Write-Error
            }
            else {            
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Run job [(dba) Run-BlitzIndex] to create table dbo.BlitzIndex." | Write-Host -ForegroundColor Red        
                "STOP here, and fix above issue." | Write-Error
            }
        }
        else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Seems table exists now. Convert [dbo].[BlitzIndex] into partitioned table.."
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlPartitionBlitzIndex -EnableException
        }
    }

    # Create CX on dbo.BlitzIndex_Mode0
    $tableName = "BlitzIndex_Mode0"
    if($tableName -eq "BlitzIndex_Mode0")
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$BlitzIndexMode0PartitionFilePath = '$BlitzIndexMode0PartitionFilePath'"
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "ALTER [dbo].[BlitzIndex_Mode0] table to partitioned table on [$SqlInstanceToBaseline].."
        $sqlPartitionBlitzIndexMode0 = [System.IO.File]::ReadAllText($BlitzIndexMode0PartitionFilePath).Replace("[DBA]", "[$DbaDatabase]")

        # If Task Scheduler Job, then execute one time manually
        if($ForceSetupOfTaskSchedulerJobs) {
            $sql2RunManually = "EXEC master.dbo.sp_BlitzIndex @GetAllDatabases = 1, @Mode = 0, @BringThePain = 1, @OutputDatabaseName = '$DbaDatabase', @OutputSchemaName = 'dbo', @OutputTableName = 'BlitzIndex_Mode0';"
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Since ForceSetupOfTaskSchedulerJobs is enabled, executing 'dbo.sp_BlitzIndex @Mode = 0' once manually.."
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sql2RunManually -EnableException
        }

        # Modify content if SQL Server does not support Partitioning
        if($IsNonPartitioned) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Remove partition flag from '$BlitzIndexMode0PartitionFileName'.."
            $sqlPartitionBlitzIndexMode0 = $sqlPartitionBlitzIndexMode0.Replace('declare @is_partitioned bit = 1;', 'declare @is_partitioned bit = 0;')
        }
    
        $BlitzIndexMode0Exists = @()
        $loopStartTime = Get-Date
        $sleepDurationSeconds = 30
        $loopTotalDurationThresholdSeconds = $JobsExecutionWaitTimeoutMinutes*60
    
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Check for existence of table [dbo].[BlitzIndex_Mode0] on [$SqlInstanceToBaseline].."
        while ($BlitzIndexMode0Exists.Count -eq 0 -and $( (New-TimeSpan $loopStartTime $(Get-Date)).TotalSeconds -le $loopTotalDurationThresholdSeconds ) )
        {
            $BlitzIndexMode0Exists += $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase `
                                        -Query "if OBJECT_ID('dbo.BlitzIndex_Mode0') is not null select OBJECT_ID('dbo.BlitzIndex_Mode0') as BlitzIndexObjectID" -EnableException

            if($BlitzIndexMode0Exists.Count -eq 0) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Wait for $sleepDurationSeconds seconds as table dbo.BlitzIndex_Mode0 still does not exist.."
                Start-Sleep -Seconds $sleepDurationSeconds
            }
        }

        if($BlitzIndexMode0Exists.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
                "Run job [(dba) Run-BlitzIndex - Weekly] to create table dbo.BlitzIndex_Mode0." | Write-Error
            }
            else {            
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Run job [(dba) Run-BlitzIndex - Weekly] to create table dbo.BlitzIndex_Mode0." | Write-Host -ForegroundColor Red        
                "STOP here, and fix above issue." | Write-Error
            }
        }
        else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Seems table exists now. Convert [dbo].[BlitzIndex_Mode0] into partitioned table.."
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlPartitionBlitzIndexMode0 -EnableException
        }
    }

    # Create CX on dbo.BlitzIndex_Mode1
    $tableName = "BlitzIndex_Mode1"
    if($tableName -eq "BlitzIndex_Mode1")
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$BlitzIndexMode1PartitionFilePath = '$BlitzIndexMode1PartitionFilePath'"
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "ALTER [dbo].[BlitzIndex_Mode1] table to partitioned table on [$SqlInstanceToBaseline].."
        $sqlPartitionBlitzIndexMode1 = [System.IO.File]::ReadAllText($BlitzIndexMode1PartitionFilePath).Replace("[DBA]", "[$DbaDatabase]")

        # If Task Scheduler Job, then execute one time manually
        if($ForceSetupOfTaskSchedulerJobs) {
            $sql2RunManually = "EXEC master.dbo.sp_BlitzIndex @GetAllDatabases = 1, @Mode = 1, @BringThePain = 1, @OutputDatabaseName = '$DbaDatabase', @OutputSchemaName = 'dbo', @OutputTableName = 'BlitzIndex_Mode1';"
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Since ForceSetupOfTaskSchedulerJobs is enabled, executing 'dbo.sp_BlitzIndex @Mode = 1' once manually.."
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sql2RunManually -EnableException
        }

        # Modify content if SQL Server does not support Partitioning
        if($IsNonPartitioned) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Remove partition flag from '$BlitzIndexMode1PartitionFileName'.."
            $sqlPartitionBlitzIndexMode1 = $sqlPartitionBlitzIndexMode1.Replace('declare @is_partitioned bit = 1;', 'declare @is_partitioned bit = 0;')
        }
    
        $BlitzIndexMode1Exists = @()
        $loopStartTime = Get-Date
        $sleepDurationSeconds = 30
        $loopTotalDurationThresholdSeconds = $JobsExecutionWaitTimeoutMinutes*60
    
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Check for existence of table [dbo].[BlitzIndex_Mode1] on [$SqlInstanceToBaseline].."
        while ($BlitzIndexMode1Exists.Count -eq 0 -and $( (New-TimeSpan $loopStartTime $(Get-Date)).TotalSeconds -le $loopTotalDurationThresholdSeconds ) )
        {
            $BlitzIndexMode1Exists += $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase `
                                        -Query "if OBJECT_ID('dbo.BlitzIndex_Mode1') is not null select OBJECT_ID('dbo.BlitzIndex_Mode1') as BlitzIndexObjectID" -EnableException

            if($BlitzIndexMode1Exists.Count -eq 0) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Wait for $sleepDurationSeconds seconds as table dbo.BlitzIndex_Mode1 still does not exist.."
                Start-Sleep -Seconds $sleepDurationSeconds
            }
        }

        if($BlitzIndexMode1Exists.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
                "Run job [(dba) Run-BlitzIndex - Weekly] to create table dbo.BlitzIndex_Mode1." | Write-Error
            }
            else {            
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Run job [(dba) Run-BlitzIndex - Weekly] to create table dbo.BlitzIndex_Mode1." | Write-Host -ForegroundColor Red        
                "STOP here, and fix above issue." | Write-Error
            }
        }
        else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Seems table exists now. Convert [dbo].[BlitzIndex_Mode1] into partitioned table.."
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlPartitionBlitzIndexMode1 -EnableException
        }
    }

    # Create CX on dbo.BlitzIndex_Mode4
    $tableName = "BlitzIndex_Mode4"
    if($tableName -eq "BlitzIndex_Mode4")
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$BlitzIndexMode4PartitionFilePath = '$BlitzIndexMode4PartitionFilePath'"
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "ALTER [dbo].[BlitzIndex_Mode4] table to partitioned table on [$SqlInstanceToBaseline].."
        $sqlPartitionBlitzIndexMode4 = [System.IO.File]::ReadAllText($BlitzIndexMode4PartitionFilePath).Replace("[DBA]", "[$DbaDatabase]")

        # If Task Scheduler Job, then execute one time manually
        if($ForceSetupOfTaskSchedulerJobs) {
            $sql2RunManually = "EXEC master.dbo.sp_BlitzIndex @GetAllDatabases = 1, @Mode = 4, @BringThePain = 1, @OutputDatabaseName = '$DbaDatabase', @OutputSchemaName = 'dbo', @OutputTableName = 'BlitzIndex_Mode4';"
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Since ForceSetupOfTaskSchedulerJobs is enabled, executing 'dbo.sp_BlitzIndex @Mode = 4' once manually.."
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sql2RunManually -EnableException
        }

        # Modify content if SQL Server does not support Partitioning
        if($IsNonPartitioned) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Remove partition flag from '$BlitzIndexMode4PartitionFileName'.."
            $sqlPartitionBlitzIndexMode4 = $sqlPartitionBlitzIndexMode4.Replace('declare @is_partitioned bit = 1;', 'declare @is_partitioned bit = 0;')
        }
    
        $BlitzIndexMode4Exists = @()
        $loopStartTime = Get-Date
        $sleepDurationSeconds = 30
        $loopTotalDurationThresholdSeconds = $JobsExecutionWaitTimeoutMinutes*60
    
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Check for existence of table [dbo].[BlitzIndex_Mode4] on [$SqlInstanceToBaseline].."
        while ($BlitzIndexMode4Exists.Count -eq 0 -and $( (New-TimeSpan $loopStartTime $(Get-Date)).TotalSeconds -le $loopTotalDurationThresholdSeconds ) )
        {
            $BlitzIndexMode4Exists += $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase `
                                        -Query "if OBJECT_ID('dbo.BlitzIndex_Mode4') is not null select OBJECT_ID('dbo.BlitzIndex_Mode4') as BlitzIndexObjectID" -EnableException

            if($BlitzIndexMode4Exists.Count -eq 0) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Wait for $sleepDurationSeconds seconds as table dbo.BlitzIndex_Mode4 still does not exist.."
                Start-Sleep -Seconds $sleepDurationSeconds
            }
        }

        if($BlitzIndexMode4Exists.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
                "Run job [(dba) Run-BlitzIndex - Weekly] to create table dbo.BlitzIndex_Mode4." | Write-Error
            }
            else {            
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Run job [(dba) Run-BlitzIndex - Weekly] to create table dbo.BlitzIndex_Mode4." | Write-Host -ForegroundColor Red        
                "STOP here, and fix above issue." | Write-Error
            }
        }
        else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Seems table exists now. Convert [dbo].[BlitzIndex_Mode4] into partitioned table.."
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlPartitionBlitzIndexMode4 -EnableException
        }
    }
}


# 54__BlitzPartition
$stepName = '54__BlitzPartition'
#if($stepName -in $Steps2Execute -and $IsNonPartitioned -eq $false) {
if($stepName -in $Steps2Execute) 
{
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."

    # Create CX on dbo.Blitz
    $tableName = "Blitz"
    if($tableName -eq "Blitz")
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$BlitzPartitionFilePath = '$BlitzPartitionFilePath'"
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "ALTER [dbo].[Blitz] table to partitioned table on [$SqlInstanceToBaseline].."
        $sqlPartitionBlitz = [System.IO.File]::ReadAllText($BlitzPartitionFilePath).Replace("[DBA]", "[$DbaDatabase]")

        # If Task Scheduler Job, then execute one time manually
        if($ForceSetupOfTaskSchedulerJobs) {
            $sql2RunManually = "EXEC master.dbo.sp_Blitz @CheckUserDatabaseObjects = 1, @BringThePain = 1, @CheckServerInfo = 1, @OutputDatabaseName = '$DbaDatabase', @OutputSchemaName = 'dbo', @OutputTableName = 'Blitz';"
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Since ForceSetupOfTaskSchedulerJobs is enabled, executing 'dbo.sp_Blitz' once manually.."
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sql2RunManually -EnableException
        }

        # Modify content if SQL Server does not support Partitioning
        if($IsNonPartitioned) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Remove partition flag from '$BlitzPartitionFileName'.."
            $sqlPartitionBlitz = $sqlPartitionBlitz.Replace('declare @is_partitioned bit = 1;', 'declare @is_partitioned bit = 0;')
        }
    
        $BlitzExists = @()
        $loopStartTime = Get-Date
        $sleepDurationSeconds = 30
        $loopTotalDurationThresholdSeconds = $JobsExecutionWaitTimeoutMinutes*60
    
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Check for existence of table [dbo].[Blitz] on [$SqlInstanceToBaseline].."
        while ($BlitzExists.Count -eq 0 -and $( (New-TimeSpan $loopStartTime $(Get-Date)).TotalSeconds -le $loopTotalDurationThresholdSeconds ) )
        {
            $BlitzExists += $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase `
                                        -Query "if OBJECT_ID('dbo.Blitz') is not null select OBJECT_ID('dbo.Blitz') as BlitzIndexObjectID" -EnableException

            if($BlitzExists.Count -eq 0) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Wait for $sleepDurationSeconds seconds as table dbo.Blitz still does not exist.."
                Start-Sleep -Seconds $sleepDurationSeconds
            }
        }

        if($BlitzExists.Count -eq 0) {
            if ($ReturnInlineErrorMessage) {
                "Run job [(dba) Run-Blitz] to create table dbo.Blitz." | Write-Error
            }
            else {            
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Run job [(dba) Run-Blitz] to create table dbo.Blitz." | Write-Host -ForegroundColor Red        
                "STOP here, and fix above issue." | Write-Error
            }
        }
        else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Seems table exists now. Convert [dbo].[Blitz] into partitioned table.."
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlPartitionBlitz -EnableException
        }
    }
}


# 55__EnablePageCompression
$stepName = '55__EnablePageCompression'
if( ($stepName -in $Steps2Execute) -and ($SkipPageCompression -eq $false) -and $IsCompressionSupported) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."

    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Execute procedure [usp_enable_page_compression] on [$SqlInstanceToBaseline].."
    $sqlExecuteUspEnablePageCompression = "exec dbo.usp_enable_page_compression;"
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlExecuteUspEnablePageCompression -EnableException
}

# 56__GrafanaLogin
$stepName = '56__GrafanaLogin'
if ($stepName -in $Steps2Execute) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.." | Write-Host

    $sqlResidualDcl = [System.IO.File]::ReadAllText($GrafanaLoginFilePath).Replace("[DBA]", "[$DbaDatabase]")

    if ($SqlInstanceToBaselineWithOutPort -ne $InventoryServerWithOutPort) {
        # ── Non-inventory: copy login from inventory, then apply residual DCL ──
        try {
            Copy-DbaLogin -Source $conInventoryServer -Destination $conSqlInstanceToBaseline `
                -Login 'grafana' -Force -EnableException
        }
        catch {
            $errMessage = $_.Exception.Message
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', `
                "Error occurred while using Copy-DbaLogin for [grafana].`n`n$errMessage" |
                Write-Host -ForegroundColor Red
            "STOP here, and fix above issue." | Write-Error
        }
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database master -Query $sqlResidualDcl -EnableException
    }
    else {
        # ── Inventory server: create or conditionally rotate ──
        $grafanaLogin = Get-DbaLogin -SqlInstance $conSqlInstanceToBaseline -Login 'grafana'

        if ($null -eq $grafanaLogin) {
            # Branch 1 — fresh install: create login
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', `
                "Creating [grafana] login (password value redacted).." | Write-Host
            New-DbaLogin -SqlInstance $conSqlInstanceToBaseline `
                -Login 'grafana' `
                -SecurePassword $GrafanaLoginPassword `
                -DefaultDatabase $DbaDatabase `
                -DisablePasswordPolicy `
                -DisablePasswordExpiration `
                -EnableException
        }
        elseif ($RotateGrafanaLoginPassword) {
            # Branch 2 — rotation
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', `
                "Rotating [grafana] login password (value redacted).." | Write-Host
            Set-DbaLogin -SqlInstance $conSqlInstanceToBaseline `
                -Login 'grafana' `
                -SecurePassword $GrafanaLoginPassword `
                -EnableException
        }
        else {
            # Branch 3 — login exists, no rotation flag: leave login untouched
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', `
                "[grafana] login exists; -RotateGrafanaLoginPassword not set. Login left untouched." | Write-Host
        }

        # Residual DCL — always; recovers grant drift idempotently
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database master -Query $sqlResidualDcl -EnableException
    }
}

# 57__LinkedServerOnInventory
$stepName = '57__LinkedServerOnInventory'
if($stepName -in $Steps2Execute -and $SqlInstanceToBaselineWithOutPort -ne $InventoryServerWithOutPort) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.." | Write-Host
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$LinkedServerOnInventoryFilePath = '$LinkedServerOnInventoryFilePath'" | Write-Host
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating linked server for [$SqlInstanceToBaseline] on [$InventoryServer].." | Write-Host
    #$sqlLinkedServerOnInventory = [System.IO.File]::ReadAllText($LinkedServerOnInventoryFilePath).Replace("'YourSqlInstanceNameHere'", "'$SqlInstanceToBaseline'")
    $sqlLinkedServerOnInventory = [System.IO.File]::ReadAllText($LinkedServerOnInventoryFilePath)
    $sqlLinkedServerOnInventory = $sqlLinkedServerOnInventory.Replace("@server = N'YourSqlInstanceNameHere'", "@server = N'$SqlInstanceToBaselineWithOutPort'")
    $sqlLinkedServerOnInventory = $sqlLinkedServerOnInventory.Replace("@server=N'YourSqlInstanceNameHere'", "@server=N'$SqlInstanceToBaselineWithOutPort'")
    $sqlLinkedServerOnInventory = $sqlLinkedServerOnInventory.Replace("@rmtsrvname=N'YourSqlInstanceNameHere'", "@rmtsrvname=N'$SqlInstanceToBaselineWithOutPort'")
    $sqlLinkedServerOnInventory = $sqlLinkedServerOnInventory.Replace("@datasrc=N'YourSqlInstanceNameHere'", "@datasrc=N'$SqlInstanceToBaseline'")
    $sqlLinkedServerOnInventory = $sqlLinkedServerOnInventory.Replace("@catalog=N'DBA'", "@catalog=N'$DbaDatabase'")
    
    $dbaLinkedServer = @()
    $dbaLinkedServer += Get-DbaLinkedServer -SqlInstance $conInventoryServer -LinkedServer $SqlInstanceToBaselineWithOutPort

    # Take action if Linked Server exists on Inventory
    if($dbaLinkedServer.Count -ne 0) 
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Linked server for [$SqlInstanceToBaseline] on [$InventoryServer] already exists.." | Write-Host

        if($DropCreateLinkedServerOnInventory) 
        {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Create linked server for [$SqlInstanceToBaseline] on [$InventoryServer].."
            $sqlDropCreateLinkedServerOnInventory = "EXEC master.dbo.sp_dropserver @server=N'$SqlInstanceToBaselineWithOutPort', @droplogins='droplogins'"
            try {
                $conInventoryServer | Invoke-DbaQuery -Database master -Query $sqlDropCreateLinkedServerOnInventory -EnableException -MessagesToOutput:$verbose
            }
            catch {
                $errMessage = $_.Exception.Message
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Error occurred while dropping linked server [$SqlInstanceToBaselineWithOutPort] on Inventory.`n`n$errMessage" | Write-Host -ForegroundColor Red
                Start-Sleep -Seconds 1
                "STOP here, and fix above issue." | Write-Error
            } 
            
        }
    }

    if( ($dbaLinkedServer.Count -eq 0) -or ($DropCreateLinkedServerOnInventory) ) {
        try {
            $_bstr57 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($GrafanaLoginPassword)
            $_pwd57 = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($_bstr57)
            $conInventoryServer | Invoke-DbaQuery -Database master -Query $sqlLinkedServerOnInventory -SqlParameter @{ GrafanaLoginPassword= $_pwd57 } -EnableException
        }
        finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($_bstr57)
            Remove-Variable _pwd57 -ErrorAction SilentlyContinue
        }
    }
}


# 58__LinkedServerForDataDestinationInstance
$stepName = '58__LinkedServerForDataDestinationInstance'
if( ($stepName -in $Steps2Execute) -and ($SqlInstanceToBaseline -ne $SqlInstanceAsDataDestination) )
{
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.." | Write-Host
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "`$LinkedServerOnInventoryFilePath = '$LinkedServerOnInventoryFilePath'" | Write-Host

    $sqlLinkedServerForDataDestinationInstance = [System.IO.File]::ReadAllText($LinkedServerOnInventoryFilePath)
    $sqlLinkedServerForDataDestinationInstance = $sqlLinkedServerForDataDestinationInstance.Replace("@server = N'YourSqlInstanceNameHere'", "@server = N'$SqlInstanceAsDataDestinationWithOutPort'")
    $sqlLinkedServerForDataDestinationInstance = $sqlLinkedServerForDataDestinationInstance.Replace("@server=N'YourSqlInstanceNameHere'", "@server=N'$SqlInstanceAsDataDestinationWithOutPort'")
    $sqlLinkedServerForDataDestinationInstance = $sqlLinkedServerForDataDestinationInstance.Replace("@rmtsrvname=N'YourSqlInstanceNameHere'", "@rmtsrvname=N'$SqlInstanceAsDataDestinationWithOutPort'")
    $sqlLinkedServerForDataDestinationInstance = $sqlLinkedServerForDataDestinationInstance.Replace("@datasrc=N'YourSqlInstanceNameHere'", "@datasrc=N'$SqlInstanceAsDataDestination'")
    $sqlLinkedServerForDataDestinationInstance = $sqlLinkedServerForDataDestinationInstance.Replace("@catalog=N'DBA'", "@catalog=N'$DbaDatabase'")
    
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Checking if linked server already exists.." | Write-Host
    $dbaLinkedServer = @()
    $dbaLinkedServer += Get-DbaLinkedServer -SqlInstance $conSqlInstanceToBaseline -LinkedServer $SqlInstanceAsDataDestinationWithOutPort -EnableException

    # Take action if Linked Server exists for DataDestinationServer
    if($dbaLinkedServer.Count -ne 0) 
    {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Linked server for [SqlInstanceAsDataDestination] on [$SqlInstanceToBaseline] already exists.." | Write-Host

        if($DropCreateLinkedServerForDataDestinationServer) 
        {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Drop/Create linked server for [$SqlInstanceAsDataDestination] on [$SqlInstanceToBaseline].."
            $sqlDropCreateLinkedServerForDataDestination = "EXEC master.dbo.sp_dropserver @server=N'$SqlInstanceAsDataDestinationWithOutPort', @droplogins='droplogins'"
            try {
                $conSqlInstanceToBaseline | Invoke-DbaQuery -Database master -Query $sqlDropCreateLinkedServerForDataDestination -EnableException -MessagesToOutput:$verbose
            }
            catch {
                $errMessage = $_.Exception.Message
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Error occurred while dropping linked server [$SqlInstanceAsDataDestinationWithOutPort] on [$SqlInstanceToBaseline].`n`n$errMessage" | Write-Host -ForegroundColor Red
                Start-Sleep -Seconds 1
                "STOP here, and fix above issue." | Write-Error
            }
        }
    }

    if( ($dbaLinkedServer.Count -eq 0) -or ($DropCreateLinkedServerForDataDestinationServer) ) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Creating linked server named [$SqlInstanceAsDataDestination] on [$SqlInstanceToBaseline].." | Write-Host
        try {
            $_bstr58 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($GrafanaLoginPassword)
            $_pwd58 = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($_bstr58)
            $conSqlInstanceToBaseline | Invoke-DbaQuery -Database master -Query $sqlLinkedServerForDataDestinationInstance -SqlParameter @{ GrafanaLoginPassword= $_pwd58 } -EnableException
        }
        finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($_bstr58)
            Remove-Variable _pwd58 -ErrorAction SilentlyContinue
        }
    } else 
    {
        if( ($dbaLinkedServer.Count -ne 0) -and ($DropCreateLinkedServerForDataDestinationServer -eq $false) ) {
		    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Linked server named [$SqlInstanceAsDataDestination] already exists on [$SqlInstanceToBaseline]." | Write-Host -ForegroundColor Red
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly validate if linked server is able to access data of [$SqlInstanceAsDataDestination].[$DbaDatabase] database." | Write-Host -ForegroundColor Red
            "STOP and check above error message" | Write-Error
        }
    }
}


# 59__AlterViewsForDataDestinationInstance
$stepName = '59__AlterViewsForDataDestinationInstance'
if( ($stepName -in $Steps2Execute) -and ($SqlInstanceToBaselineWithOutPort -ne $SqlInstanceAsDataDestinationWithOutPort) )
{
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."

    # Alter dbo.vw_performance_counters
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Alter view [dbo].[vw_performance_counters].."
    $sqlAlterViewPerformanceCounters = @"
alter view dbo.vw_performance_counters
as
with cte_counters_local as (select collection_time_utc, host_name, object, counter, value, instance from dbo.performance_counters)
,cte_counters_datasource as (select collection_time_utc, host_name, object, counter, value, instance from [$SqlInstanceAsDataDestinationWithOutPort].[$DbaDatabase].dbo.performance_counters)

select collection_time_utc, host_name, object, counter, value, instance from cte_counters_local
union all
select collection_time_utc, host_name, object, counter, value, instance from cte_counters_datasource
"@
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlAlterViewPerformanceCounters -EnableException


    # Alter dbo.vw_disk_space
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Alter view [dbo].[vw_disk_space].."
    $sqlAlterViewDiskSpace = @"
alter view dbo.vw_disk_space
as
with cte_disk_space_local as (select collection_time_utc, host_name, disk_volume, label, capacity_mb, free_mb, block_size, filesystem from dbo.disk_space)
,cte_disk_space_datasource as (select collection_time_utc, host_name, disk_volume, label, capacity_mb, free_mb, block_size, filesystem from [$SqlInstanceAsDataDestinationWithOutPort].[$DbaDatabase].dbo.disk_space)

select collection_time_utc, host_name, disk_volume, label, capacity_mb, free_mb, block_size, filesystem from cte_disk_space_local
union all
select collection_time_utc, host_name, disk_volume, label, capacity_mb, free_mb, block_size, filesystem from cte_disk_space_datasource
go
"@
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlAlterViewDiskSpace -EnableException
}


# Update SQLMonitor Jobs Thresholds
if($UpdateSQLAgentJobsThreshold) 
{
    $sqlStartJob = @"
declare @object_id int;
set @object_id = OBJECT_ID('dbo.sql_agent_job_thresholds');

if @object_id is null
begin
    $(if($ForceSetupOfTaskSchedulerJobs){'--'}else{''})exec msdb.dbo.sp_start_job @job_name = '(dba) Check-SQLAgentJobs';
    $(if($ForceSetupOfTaskSchedulerJobs){''}else{'--'})EXEC dbo.usp_check_sql_agent_jobs @default_mail_recipient = '$($DbaGroupMailId -join ';')';
end
else
    select [object_id] = @object_id;
"@

    try
    {
        if($true)
        {
            "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Start job [(dba) Check-SQLAgentJobs] on [$SqlInstanceForTsqlJobs].." | Write-Host -ForegroundColor Cyan
            $sqlResult = @()
            if($ForceSetupOfTaskSchedulerJobs) {
                $sqlResult += $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlStartJob -EnableException -MessagesToOutput:$verbose
            } else {
                $sqlResult += $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database msdb -Query "exec msdb.dbo.sp_start_job @job_name = '(dba) Check-SQLAgentJobs';" -EnableException -MessagesToOutput:$verbose
            }

            if ($sqlResult.Count -eq 0) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [(dba) Check-SQLAgentJobs] on [$SqlInstanceForTsqlJobs] started & waiting for 10 seconds.."
                Start-Sleep -Seconds 10
            }
        }

        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update SQLMonitor jobs thresholds on [$SqlInstanceForTsqlJobs].[$DbaDatabase]..[sql_agent_job_thresholds] using '$UpdateSQLAgentJobsThresholdFileName'.." | Write-Host -ForegroundColor Cyan   
        $conSqlInstanceForTsqlJobs | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UpdateSQLAgentJobsThresholdFilePath) -EnableException
    }
    catch {
        $errMessage = $_.Exception.Message

        if ($ReturnInlineErrorMessage) {
            if($errMessage -like "Invalid object name 'dbo.sql_agent_job_thresholds*") {
                $errMessage = "Kindly ensure all SQLAgent jobs on [$SqlInstanceForTsqlJobs], and then finally [(dba) Check-SQLAgentJobs] is executed at least once, and then retry from this step.`n$errMessage";
            }
            "$errMessage.`n" | Write-Error
        }
        else {            
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "$errMessage" | Write-Host -ForegroundColor Red
            if($errMessage -like "Invalid object name 'dbo.sql_agent_job_thresholds*") {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly ensure all SQLAgent jobs on [$SqlInstanceForTsqlJobs], and then finally [(dba) Check-SQLAgentJobs] is executed at least once, and then retry from this step." | Write-Host -ForegroundColor Red
            }
            "STOP here, and fix above issue." | Write-Error
        }
    }


    try 
    {
        if($SqlInstanceForPowershellJobs -ne $SqlInstanceForTsqlJobs)
        {
            "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Start job [(dba) Check-SQLAgentJobs] on [$SqlInstanceForPowershellJobs].." | Write-Host -ForegroundColor Cyan
            $sqlResult = @()
            if($ForceSetupOfTaskSchedulerJobs) {
                $sqlResult += $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlStartJob -EnableException -MessagesToOutput:$verbose
            } else {
                $sqlResult += $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database msdb -Query "exec msdb.dbo.sp_start_job @job_name = '(dba) Check-SQLAgentJobs';" -EnableException -MessagesToOutput:$verbose
            }

            if ($sqlResult.Count -eq 0) {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Job [(dba) Check-SQLAgentJobs] on [$SqlInstanceForPowershellJobs] started & waiting for 10 seconds.."
                Start-Sleep -Seconds 10
            }
        }
        
        if($SqlInstanceForPowershellJobs -ne $SqlInstanceForTsqlJobs) {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update SQLMonitor jobs thresholds on [$SqlInstanceForPowershellJobs].[$DbaDatabase]..[sql_agent_job_thresholds] using '$UpdateSQLAgentJobsThresholdFileName'.." | Write-Host -ForegroundColor Cyan   
            $conSqlInstanceForPowershellJobs | Invoke-DbaQuery -Database $DbaDatabase -File $(Get-Item -LiteralPath $UpdateSQLAgentJobsThresholdFilePath) -EnableException
        }
    }
    catch {
        $errMessage = $_.Exception.Message
        if ($ReturnInlineErrorMessage) {
            if($errMessage -like "Invalid object name 'dbo.sql_agent_job_thresholds*") {
                $errMessage = "Kindly ensure all SQLAgent jobs on [$SqlInstanceForPowershellJobs], and then finally [(dba) Check-SQLAgentJobs] is executed at least once, and then retry from this step.`n$errMessage";
            }
            "$errMessage.`n" | Write-Error
        }
        else {
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "$errMessage" | Write-Host -ForegroundColor Red
            if($errMessage -like "Invalid object name 'dbo.sql_agent_job_thresholds*") {
                "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', "Kindly ensure all SQLAgent jobs on [$SqlInstanceForPowershellJobs], and then finally [(dba) Check-SQLAgentJobs] is executed at least once, and then retry from this step." | Write-Host -ForegroundColor Red
            }
            "STOP here, and fix above issue." | Write-Error
        }
    }
}

# Fix Broken Alerts if any
if ($true) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Executing PostQuery dbo.usp_create_agent_alerts on [$SqlInstanceToBaseline].." | Write-Host -ForegroundColor Cyan
    $sqlQuery = 'exec dbo.usp_create_agent_alerts'
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlQuery -EnableException    
}

# Execute PostQuery
if(-not [String]::IsNullOrEmpty($PostQuery)) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Executing PostQuery on [$SqlInstanceToBaseline].." | Write-Host -ForegroundColor Cyan
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $PostQuery -EnableException
}

# Update Version No
if( $true )
{
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update SQLMonitor Version Number.."

    $sqlUpdateSQLMonitorVersion = @"
update dbo.instance_details 
set [sqlmonitor_version] = '$sqlmonitorVersion'
where sql_instance = '$SqlInstanceToBaselineWithOutPort'
and host_name = '$HostName'
"@

    $sqlUpdateInventoryVersionEntry = @"
update dbo.instance_details 
set [sqlmonitor_version] = '$sqlmonitorVersion', [is_enabled] = 1
where sql_instance = '$SqlInstanceToBaselineWithOutPort'
and host_name = '$HostName'
"@

    # Update dbo.instance_details on SqlInstanceToBaseline
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update SQLMonitor version on [$SqlInstanceToBaseline].."
    $conSqlInstanceToBaseline | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlUpdateSQLMonitorVersion -EnableException

    # Update dbo.instance_details on SqlInstanceAsDataDestination
    if($SqlInstanceAsDataDestination -ne $SqlInstanceToBaseline) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update SQLMonitor version on [$SqlInstanceAsDataDestination].."
        $conSqlInstanceAsDataDestination | Invoke-DbaQuery -Database $DbaDatabase -Query $sqlUpdateSQLMonitorVersion -EnableException
    }

    # Update dbo.instance_details on InventoryServer
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Update SQLMonitor version on [$InventoryServer].."
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlUpdateInventoryVersionEntry -EnableException

    # Populate dbo.sma_server inventory tables
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Execute dbo.usp_populate_sma_sql_instance for [$SqlInstanceToBaselineWithOutPort].."
    $sqlExecuteUspPopulateSmaSqlInstance = @"
IF OBJECT_ID('dbo.usp_populate_sma_sql_instance', 'P') IS NOT NULL
    exec dbo.usp_populate_sma_sql_instance @server = '$SqlInstanceToBaselineWithOutPort' ,@execute = 1 ,@verbose = 1;
"@
    $allServerInfoQuery = "select 1 from dbo.vw_all_server_info where srv_name = '$SqlInstanceToBaselineWithOutPort'"
    $resultAllServerInfo = @()
    $resultAllServerInfo += $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $allServerInfoQuery -EnableException
    
    $loopStartTime = Get-Date
    $loopTimeoutSeconds = 180
    while($resultAllServerInfo.Count -eq 0 -and ($loopStartTime.AddSeconds($loopTimeoutSeconds) -gt $(Get-Date))) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "[$SqlInstanceToBaselineWithOutPort] yet not present in [$InventoryServer].[$InventoryDatabase].dbo.vw_all_server_info. So wait for few seconds.."
        
        Start-Sleep -Seconds 10

        $resultAllServerInfo += $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $allServerInfoQuery -EnableException
    }
    if($resultAllServerInfo.Count -eq 0) {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "[$SqlInstanceToBaselineWithOutPort] yet not present in [$InventoryServer].[$InventoryDatabase].dbo.vw_all_server_info." | Write-Host -ForegroundColor Yellow
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "So below statement might fail.." | Write-Host -ForegroundColor Yellow
        "`nexec dbo.usp_populate_sma_sql_instance @server = '$SqlInstanceToBaselineWithOutPort' ,@execute = 1 ,@verbose = 2;`n"
    }
    $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlExecuteUspPopulateSmaSqlInstance -EnableException
    
}


"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Clearing old PSSessions.."
Get-PSSession | Remove-PSSession

"`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Baselining of [$SqlInstanceToBaseline] completed."

if(-not ($IsManagedInstance -or $InstanceScopeFeaturesOnly)) 
{
    $agentServiceAccount = $agentServiceInfo.service_account
    $addAgentAccountToWindowsGroups = @"
`t net localgroup administrators "$agentServiceAccount" /add
`t net localgroup "Performance Log Users" "$agentServiceAccount" /add
`t net localgroup "Performance Monitor Users" "$agentServiceAccount" /add
"@
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Kindly RDP [$SqlInstanceForPowershellJobs], and execute following commands in Elevated Command Prompt-" | Write-Host -ForegroundColor Cyan
    "$addAgentAccountToWindowsGroups`n" | Write-Host -ForegroundColor Yellow
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Post executing above commands, restart SQLAgent service." | Write-Host -ForegroundColor Cyan
}

if($ConfirmSetupOfTaskSchedulerJobs -or $ForceSetupOfTaskSchedulerJobs) {
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Kindly validate DBA scheduled tasks in windows Task Scheduler." | Write-Host -ForegroundColor Yellow
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Ensure all the sechduled tasks are running successfully." | Write-Host -ForegroundColor Yellow
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Logs can be found on '$SQLMonitorPath\Logs' directory." | Write-Host -ForegroundColor Yellow    
    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "If required, change service account, and ensure task is scheduled to 'Run whether user is logged on or not'." | Write-Host -ForegroundColor Yellow
}

if( ($SqlInstanceToBaseline -eq $InventoryServer) -and ([String]::IsNullOrEmpty($url_for_dba_slack_channel) -eq $false) ) {
    if($url_for_dba_slack_channel -eq 'workspace.slack.com/archives/unique_id') {
        "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'WARNING:', "Kindly update the [param_value] in table [$InventoryDatabase].dbo.sma_params on [$InventoryServer]." | Write-Host -ForegroundColor Yellow
    }
}

$timeTaken = New-TimeSpan -Start $startTime -End $(Get-Date)
"$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "Execution completed in $($timeTaken.TotalSeconds) seconds."



<#
    .SYNOPSIS
    Baselines the SQL Server instance by creating all required objects, Perfmon data collector, and required SQL Agent jobs. Adds linked server on inventory instance.
    .DESCRIPTION
    This function accepts various parameters and perform baselining of the SQLInstance with creation of required tables, views, procedures, jobs, perfmon data collector, and linked server.
    .PARAMETER SqlInstanceToBaseline
    Name/IP of SQL Instance that has to be baselined. Instances should be capable of connecting from remove machine SSMS using this name/ip.
    .PARAMETER DbaDatabase
    Name of DBA database on the SQL Instance being baseline, and as well target on [SqlInstanceAsDataDestination].
    .PARAMETER SqlInstanceAsDataDestination
    Name/IP of SQL Instance that would store the data caputured using Perfmon data collection. Generally same as [SqlInstanceToBaseline]. But, this could be different from [SqlInstanceToBaseline] in central repository scenario.
    .PARAMETER SqlInstanceForTsqlJobs
    Name/IP of SQL Instance that could be used to host SQL Agent jobs that call tsql scripts. Generally same as [SqlInstanceToBaseline]. This can be used in case of Express Edition as agent services are not available.
    .PARAMETER SqlInstanceForPowershellJobs
    Name/IP of SQL Instance that could be used to host SQL Agent jobs that call tsql scripts. Generally same as [SqlInstanceToBaseline]. This can be used when [SqlInstanceToBaseline] is Express Edition, or not capable of running PowerShell Jobs successfully due to being old version of powershell, or incapability of install modules like dbatools.
    .PARAMETER InventoryServer
    Name/IP of SQL Instance that would act as inventory server and is the data source on Grafana application. A linked server would be created for [SqlInstanceToBaseline] on this server.
    .PARAMETER InventoryDatabase
    Name of DBA database on the InventoryServer. Default to be same as DbaDatabase.
    .PARAMETER HostName
    Name of server where Perfmon data collection & other OS level settings would be done. For standalone SQL Instances, this is not required as this value can be retrieved from tsql. But for active/passive SQLCluster setup where SQL Cluster instance may have other passive nodes, this value can be explictly passed to setup perfmon collection of other passive hosts.
    .PARAMETER IsNonPartitioned
    Switch to signify if Partitioning of table should NOT be done even if supported.
    .PARAMETER SQLMonitorPath
    Path of SQLMonitor tool parent folder. This is the folder that contains other folders/files like Alerting, Credential-Manager, DDLs, SQLMonitor, Inventory etc.
    .PARAMETER DbaToolsFolderPath
    Local directory path of dbatools powershell module that was downloaded locally from github https://github.com/dataplat/dbatools or using Save-Module.
    .PARAMETER FirstResponderKitZipFile
    Specifies the path to a local file to install FRK from. This should be the zip file as distributed by the maintainers. If this parameter is not specified, the latest version will be downloaded and installed from https://github.com/BrentOzarULTD/SQL-Server-First-Responder-Kit
    .PARAMETER DarlingDataZipFile
    Specifies the path to a local file to install from. This should be the zip file as distributed by the maintainers. If this parameter is not specified, the latest version will be downloaded and installed from https://github.com/erikdarlingdata/DarlingData
    .PARAMETER OlaHallengrenSolutionZipFile
    Specifies the path to a local file to be used for installation of Ola Hallengren Mailtenance Solution objects except jobs. It can be download using https://github.com/olahallengren/sql-server-maintenance-solution/archive/refs/heads/master.zip
    .PARAMETER RemoteSQLMonitorPath
    Desired SQLMonitor folder location on [SqlInstanceToBaseline] or [SqlInstanceForDataCollectionJobs]. At this path, folder SQLMonitor\SQLMonitor would be copied. On target instance, all the SQL Agent jobs would call the scripts from this folder.
    .PARAMETER DbaGroupMailId
    List of DBA/group email ids that should receive job failure alerts.
    .PARAMETER StartAtStep
    Starts the baselining automation on this step. If no value provided, then baselining starts with 1st step.
    .PARAMETER SkipSteps
    List of steps that should be skipped in the baselining automation.
    .PARAMETER OnlySteps
    List of steps that should be the only steps to be executed. This parameter has highest precedence and overrides other settings.
    .PARAMETER StopAtStep
    End the baselining automation on this step. If no value provided, then baselining finishes with last step.
    .PARAMETER SqlCredential
    PowerShell credential object to execute queries any SQL Servers. If no value provided, then connectivity is tried using Windows Integrated Authentication.
    .PARAMETER WindowsCredential
    PowerShell credential object that could be used to perform OS interactives tasks. If no value provided, then connectivity is tried using Windows Integrated Authentication. This is important when [SqlInstanceToBaseline] is not in same domain as current host.
    .PARAMETER RetentionDays 
    No of days as data retention threshold in tables  of SQLMonitor. Data older than this value would be purged daily once.
    .PARAMETER DropCreatePowerShellJobs
    When enabled, drops the existing SQL Agent jobs having CmdExec steps, and creates them from scratch. By default, Jobs running CmdExec step are not dropped if found existing.
    .PARAMETER DropCreateWhoIsActiveTable
    When enabled, drops the existing WhoIsActive table, and creates it from scratch. This might be required in case of change in sp_WhoIsActive features usage.
    .PARAMETER SkipPowerShellJobs
    When enabled, baselining steps involving create of SQL Agent jobs having CmdExec steps are skipped.
    .PARAMETER SkipMultiServerviewsUpgrade
    Default enabled. This skips alter of views like vw_performance_counters, vw_disk_space, vw_os_tasks_list etc which interact with multiple hosts in many cases.
    .PARAMETER SkipTsqlJobs
    When enabled, skips creation of all the SQL Agent jobs that execute tsql stored procedures.
    .PARAMETER SkipRDPSessionSteps
    When enabled, any steps that need OS level interaction is skipped. This includes copy of dbatools powershell module, SQLMonitor folder on remove path, creation of Perfmon Data Collector etc.
    .PARAMETER SkipWindowsAdminAccessTest
    When enabled, script does not check if Proxy/Credential is required for running PowerShell jobs.
    .PARAMETER SkipMailProfileCheck 
    When enabled, script does not look for default global mail profile.
    .PARAMETER SkipCollationCheck 
    When enabled, database collations checks are skipped. This means we don't validate if the collation of DBA database  is same as system databases or not.
    .PARAMETER SkipPageCompression
    When enabled, page data compression of SQLMonitor tables is skipped.
    .PARAMETER SkipDriveCheck
    When enabled, script ignores the location of DBA database even if its present in C:\ drive
    .PARAMETER SkipPingCheck
    When enabled, ping operation is not done for connectivity validation
    .PARAMETER HasCustomizedTsqlJobs
    When enabled, assumption is that SQL Agent jobs running sqlcmd/tsql queries have modified version. This flag ensures that these jobs/steps are skipped in upgrades unless OverrideCustomizedTsqlJobs parameter is used
    .PARAMETER ForceTSQLStepType4TsqlJobs
    When enabled, TSQL jobs running tsql queries are created as TSQL step type. Default is disabled. When disabled, then jobs are created as CmdExec step type running sqlcmd utility.
    .PARAMETER HasCustomizedPowerShellJobs
    When enabled, assumption is that SQL Agent jobs running powershell scripts have modified version. This flag ensures that these jobs/steps are skipped in upgrades unless OverrideCustomizedPowerShellJobs parameter is used
    .PARAMETER OverrideCustomizedTsqlJobs
    When enabled, any upgrade of SQLMonitor will drop/create the SQL Agent jobs running sqlcmd/tsql queries even if they are modified on server
    .PARAMETER OverrideCustomizedPowerShellJobs
    When enabled, any upgrade of SQLMonitor will drop/create the SQL Agent jobs running Powershell scripts even if they are modified on server
    .PARAMETER ForceSetupOfTaskSchedulerJobs
    When enabled, all the jobs are created in windows Task Scheduler instead of SQLAgent
    .PARAMETER ConfirmValidationOfMultiInstance
    If required for confirmation from end user in case multiple SQL Instances are found on same host. At max, perfmon data can be pushed to only one SQL Instance.
    .PARAMETER ConfirmSetupOfTaskSchedulerJobs
    If SQLInstance is SQL Express edition, then this parameter ensures that local SQLInstance level jobs are created in Windows Task Scheduler.
    .PARAMETER UpdateSQLAgentJobsThreshold
    when enabled, appropriate thresholds are set in table dbo.sql_agent_job_thresholds using script \DDLs\Update-SQLAgentJobsThreshold.sql
    .PARAMETER XEventDirectory
    Directory for saving Extended Events files. By default, a directory named xevents is created on parent path of DBA database files directory
    .PARAMETER JobsExecutionWaitTimeoutMinutes
    Time in minutes. Script will wait for maximum this duration for execution of SQLAgent jobs like [(dba) Run-BlitzIndex] or [(dba) Run-BlitzIndex - Weekly] etc
    .PARAMETER DryRun
    When enabled, only messages are printed, but actual changes are NOT made.
    .PARAMETER PreQuery
    TSQL String that should be executed before actual SQLMonitor scripts are run. This is useful when specific pre-changes are required for SQLMonitor. For example, drop/create few columns etc.
    .PARAMETER PostQuery
    TSQL String that should be executed after actual SQLMonitor scripts are run. This is useful when specific post-changes are required due to environment specific needs.
    .PARAMETER ReturnInlineErrorMessage
    On PS Script failure, when set to true, actual error message is returned. Otherwise, by default, user friendly error message is returned.
    .EXAMPLE
$params = @{
    SqlInstanceToBaseline = 'Workstation'
    DbaDatabase = 'DBA'
    DbaToolsFolderPath = 'F:\GitHub\dbatools'
    RemoteSQLMonitorPath = 'C:\SQLMonitor'
    InventoryServer = 'SQLMonitor'
    DbaGroupMailId = 'sqlagentservice@gmail.com'
    #SqlCredential = $saAdmin
    #WindowsCredential = $LabCredential
    #SkipSteps = @("11__SetupPerfmonDataCollector", "12__CreateJobCollectOSProcesses","13__CreateJobCollectPerfmonData")
    #StartAtStep = '56__GrafanaLogin'
    #StopAtStep = '21__WhoIsActivePartition'
    #DropCreatePowerShellJobs = $true
    #DryRun = $false
    #SkipRDPSessionSteps = $true
    #SkipPowerShellJobs = $true
    #SkipAllJobs = $true
}
F:\GitHub\SQLMonitor\SQLMonitor\Install-SQLMonitor.ps1 @Params

Baseline SQLInstance [Workstation] using [DBA] database. Use [SQLMonitor] as Inventory SQLInstance, and alerts should go to 'sqlagentservice@gmail.com'.
    .EXAMPLE
$LabCredential = Get-Credential -UserName 'Lab\SQLServices' -Message 'AD Account'
$saAdmin = Get-Credential -UserName 'sa' -Message 'sa'
#$localAdmin = Get-Credential -UserName 'Administrator' -Message 'Local Admin'

cls
$params = @{
    SqlInstanceToBaseline = 'Workstation'
    DbaDatabase = 'DBA'
    DbaToolsFolderPath = 'F:\GitHub\dbatools'
    RemoteSQLMonitorPath = 'C:\SQLMonitor'
    InventoryServer = 'SQLMonitor'
    DbaGroupMailId = 'sqlagentservice@gmail.com'
    SqlCredential = $saAdmin
    WindowsCredential = $LabCredential
    #SkipSteps = @("11__SetupPerfmonDataCollector", "12__CreateJobCollectOSProcesses","13__CreateJobCollectPerfmonData")
    #StartAtStep = '56__GrafanaLogin'
    #StopAtStep = '21__WhoIsActivePartition'
    #DropCreatePowerShellJobs = $true
    #DryRun = $false
    #SkipRDPSessionSteps = $true
    #SkipPowerShellJobs = $true
    #SkipAllJobs = $true
}
F:\GitHub\SQLMonitor\SQLMonitor\Install-SQLMonitor.ps1 @Params

Baseline SQLInstance [Workstation] using [DBA] database. Use [SQLMonitor] as Inventory SQLInstance. Alerts are sent to 'sqlagentservice@gmail.com'. Using $saAdmin credential while interacting with SQLInstance. Similary, for performing OS interactive task, use $LabCredential.
    .NOTES
Owner Ajay Kumar Dwivedi (ajay.dwivedi2007@gmail.com)
    .LINK
    https://ajaydwivedi.com/github/sqlmonitor
    https://ajaydwivedi.com/youtube/sqlmonitor
    https://ajaydwivedi.com/blog/sqlmonitor    
#>


