
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading.Tasks;

namespace AxialSqlTools
{

    public class HealthDashboardServerMetric
    {

        public class DiskInfo
        {
            public string VolumeDescription { get; set; }
            public long TotalCapacity { get; set; }
            public long TotalCapacityGb { get; set; }
            public long FreeSpace { get; set; }
            public long FreeSpaceGb { get; set; }
            public long UsedSpaceGb { get; set; }
        }

        public void AddDiskInfo(string volumeMountPoint, string volumeName, long Capacity, long FreeSpace)
        {
            DiskInfo diskInfo = new DiskInfo();

            diskInfo.VolumeDescription = volumeMountPoint;
            if (!string.IsNullOrEmpty(volumeName))
                diskInfo.VolumeDescription += $" ({volumeName})";

            diskInfo.TotalCapacity = Capacity;
            diskInfo.FreeSpace = FreeSpace;

            diskInfo.TotalCapacityGb = diskInfo.TotalCapacity / 1024 / 1024 / 1024;
            diskInfo.FreeSpaceGb = diskInfo.FreeSpace / 1024 / 1024 / 1024;
            diskInfo.UsedSpaceGb = (diskInfo.TotalCapacity - diskInfo.FreeSpace) / 1024 / 1024 / 1024;

            DisksInfo.Add(diskInfo);
        }

        public class WaitsInfo
        {
            public string WaitName { get; set; }
            public decimal WaitSec { get; set; }
        }

        public void AddWaitInfo(string WaitName, decimal WaitSec)
        {
            WaitsInfo waitsInfo = new WaitsInfo();
            waitsInfo.WaitName = WaitName;
            waitsInfo.WaitSec = WaitSec;

            WaitStatsInfo.Add(waitsInfo);
        }

        public HealthDashboardServerMetric ()
        {
            Iteration = 0;
            PerfCounter_RefreshDateTime = DateTime.Now;

            DisksInfo = new List<DiskInfo>();
            WaitStatsInfo = new List<WaitsInfo>();

        }
        // Example properties of the Metric class
        // public int Id { get; set; }
        public int Iteration { get; set; }
        public string ServerName { get; set; }
        public string ServiceName { get; set; }
        public string ServerVersion { get; set; }
        public string ServerVersionShort { get; set; }
        public DateTime UtcStartTime { get; set; }
        public int CpuUtilization { get; set; }
        public int ConnectionCountTotal { get; set; }
        public int ConnectionCountEnc { get; set; }

        public int CountUserDatabasesTotal { get; set; }
        public int CountUserDatabasesOkay { get; set; }


        public int BlockedRequestsCount { get; set; }
        public int BlockingTotalWaitTime { get; set; }

        public DateTime PerfCounter_RefreshDateTime { get; set; }
        public long PerfCounter_BatchRequestsSec_Total { get; set; }
        public long PerfCounter_BatchRequestsSec { get; set; }
        public long PerfCounter_SQLCompilationsSec_Total { get; set; }
        public long PerfCounter_SQLCompilationsSec { get; set; }

        public long PerfCounter_PageReadsSec_Total { get; set; }
        public long PerfCounter_PageReadsSec { get; set; }
        public long PerfCounter_PageWritesSec_Total { get; set; }
        public long PerfCounter_PageWritesSec { get; set; }
        public long PerfCounter_LogFlushesSec_Total { get; set; }
        public long PerfCounter_LogFlushesSec { get; set; }
        public long PerfCounter_TransactionsSec_Total { get; set; }
        public long PerfCounter_TransactionsSec { get; set; }
        public long PerfCounter_LockWaitsSec_Total { get; set; }
        public long PerfCounter_LockWaitsSec { get; set; }
        public long PerfCounter_MemoryGrantsPending { get; set; }

        public long PerfCounter_PLE { get; set; }
        public long PerfCounter_DataFileSize { get; set; }
        public long PerfCounter_LogFileSize { get; set; }
        public long PerfCounter_UsedLogFileSize { get; set; }
        public long ServerMemoryTotal { get; set; }
        public long PerfCounter_TargetServerMemory { get; set; }
        public long PerfCounter_TotalServerMemory { get; set; }
        public long PerfCounter_AlwaysOn_LogSendQueue { get; set; }

        public bool AlwaysOn_Exists { get; set; }
        public int AlwaysOn_Health { get; set; }
        public int AlwaysOn_MaxLatency { get; set; }
        public long AlwaysOn_TotalRedoQueueSize { get; set; }
        public long AlwaysOn_TotalLogSentQueueSize { get; set; }

        public bool spWhoIsActiveExists { get; set; }


        public DateTime SlowMetrics_RefreshDateTime { get; set; }
        public List<DiskInfo> DisksInfo { get; set; }
        public List<WaitsInfo> WaitStatsInfo { get; set; }

        // infra
        public int ServerResponseTimeMs { get; set; }
        public bool Completed { get; set; }
        public bool HasException { get; set; }
        public string ExecutionException { get; set; }



    }
    
    public static class MetricsService
    {


        public static async Task<HealthDashboardServerMetric> FetchServerMetricsAsync(string connectionString, HealthDashboardServerMetric prev_metrics)
        {
            var metrics = new HealthDashboardServerMetric {};

            metrics.Iteration = prev_metrics.Iteration + 1;
            metrics.SlowMetrics_RefreshDateTime = prev_metrics.SlowMetrics_RefreshDateTime;

            try
            {

                string perfCounterObjectName = "SQLServer";

                Stopwatch stopwatch = Stopwatch.StartNew();

                // Create and open a connection to SQL Server
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    // Open the connection
                    await connection.OpenAsync();

                    string queryText_1 = @"
                    SELECT @@SERVERNAME AS ServerName,
                           @@SERVICENAME AS ServiceName,
                           DATEADD(hh, DATEDIFF(hh, GETDATE(), GETUTCDATE()), sqlserver_start_time) AS UtcStartTime,
                           @@VERSION,
                           SERVERPROPERTY('edition'),
                           CASE
                                WHEN OBJECT_ID('dbo.sp_WhoIsActive') IS NULL
                                    THEN 0
                                ELSE 1
                           END AS spWhoIsActiveExists,
                           SERVERPROPERTY('productversion') 
                    FROM sys.dm_os_sys_info;
                    ";

                    using (SqlCommand command = new SqlCommand(queryText_1, connection))
                    { using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        { if (reader.HasRows)
                            { while (await reader.ReadAsync())
                                {
                                    metrics.ServerName = reader.GetString(0);
                                    metrics.ServiceName = reader.GetString(1);

                                    if (metrics.ServiceName != "MSSQLSERVER")
                                        perfCounterObjectName = "MSSQL$" + metrics.ServiceName;

                                    metrics.UtcStartTime = DateTime.SpecifyKind(reader.GetDateTime(2), DateTimeKind.Utc);  

                                    metrics.ServerVersion = reader.GetString(3);
                                    metrics.ServerVersionShort = reader.GetString(6);

                                    int index = metrics.ServerVersion.IndexOf("Copyright");
                                    if (index != -1)
                                        metrics.ServerVersion = metrics.ServerVersion.Substring(0, index).Trim().Replace("\r", "").Replace("\n", "");

                                    metrics.ServerVersion += " | " + reader.GetString(4);

                                    metrics.spWhoIsActiveExists = (reader.GetInt32(5) == 1);

                                }
                            }
                        }
                    }

                    string queryText_2 = @"
                    SELECT COUNT(*) AS ConnectionCountTotal,
                           SUM(CASE WHEN encrypt_option = 'TRUE' THEN 1 ELSE 0 END) AS ConnectionCountEnc
                    FROM sys.dm_exec_connections;
                    ";

                    using (SqlCommand command = new SqlCommand(queryText_2, connection))
                    { using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        { if (reader.HasRows)
                            { while (await reader.ReadAsync())
                                {
                                    metrics.ConnectionCountTotal = reader.GetInt32(0);
                                    metrics.ConnectionCountEnc = reader.GetInt32(1);
                                }
                            }
                        }
                    }

                    string queryText_3 = @"
                    SELECT
                        AVG(SQLProcessUtilization)
                    FROM (
                        SELECT 
                            record.value('(./Record/@id)[1]', 'int') AS record_id,
                            record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]', 'int') AS SystemIdle,
                            record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]', 'int') AS SQLProcessUtilization,
                            timestamp
                        FROM (
                            SELECT 
                                timestamp, 
                                CONVERT(xml, record) AS record 
                            FROM sys.dm_os_ring_buffers 
                            WHERE ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR' 
                            AND record LIKE '%<SystemHealth>%'
                        ) AS x
                    ) AS y;
                    ";

                    using (SqlCommand command = new SqlCommand(queryText_3, connection))
                    { using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        { if (reader.HasRows)
                            { while (await reader.ReadAsync())
                                {
                                    metrics.CpuUtilization = reader.GetInt32(0);
                                }
                            }
                        }
                    }

                    string queryText_4 = @"                
                    SELECT ISNULL(COUNT(*), 0) AS BlockedRequestCount,
                           ISNULL(SUM(wait_time) / 1000, 0) AS TotalWaitTime
                    FROM sys.dm_exec_requests AS er
                    WHERE blocking_session_id > 0;
                    ";

                    using (SqlCommand command = new SqlCommand(queryText_4, connection))
                    {
                        using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        {
                            if (reader.HasRows)
                            {
                                while (await reader.ReadAsync())
                                {
                                    metrics.BlockedRequestsCount = reader.GetInt32(0);
                                    metrics.BlockingTotalWaitTime = reader.GetInt32(1);
                                }
                            }
                        }
                    }

                    string queryText_5a = $@"                
                    SELECT [object_name], [counter_name], [cntr_value]
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:Buffer Manager'
                          AND [counter_name] = 'Page life expectancy'
                          AND [instance_name] = ''
                    UNION ALL

                    SELECT [object_name], [counter_name], SUM([cntr_value])
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:Databases'
                          AND [counter_name] = 'Log File(s) Used Size (KB)'
                          AND [instance_name] NOT IN ('_Total', 'tempdb', 'master', 'model', 'msdb', 'mssqlsystemresource')
                    GROUP BY [object_name], [counter_name]
                    UNION ALL

                    SELECT [object_name], [counter_name], SUM([cntr_value])
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:Databases'
                          AND [counter_name] IN ('Data File(s) Size (KB)', 'Log File(s) Size (KB)')
                          AND [instance_name] NOT IN ('_Total', 'tempdb', 'master', 'model', 'msdb', 'mssqlsystemresource')
                    GROUP BY [object_name], [counter_name]
                    UNION ALL
    
                    SELECT [object_name], [counter_name], [cntr_value]
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:Memory Manager'
                          AND [counter_name] IN ('Total Server Memory (KB)', 'Target Server Memory (KB)')
                    UNION ALL
    
                    SELECT [object_name], [counter_name], [cntr_value]
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:Database Replica'
                          AND [counter_name] = 'Log Send Queue'
                          AND [instance_name] = '_Total'
                    UNION ALL

                    SELECT 'TOTAL-SERVER-MEMORY', 'TOTAL-SERVER-MEMORY', physical_memory_kb 
                    FROM sys.dm_os_sys_info
                    ";

                    using (SqlCommand command = new SqlCommand(queryText_5a, connection))
                    {
                        using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        {
                            if (reader.HasRows)
                            {
                                while (await reader.ReadAsync())
                                {

                                    string CounterGroup = reader.GetString(0).Trim();
                                    string CounterName = reader.GetString(1).Trim();
                                    long CounterValue = reader.GetInt64(2);

                                    if (CounterGroup.EndsWith("Buffer Manager"))
                                    {
                                        if (CounterName == "Page life expectancy")
                                            metrics.PerfCounter_PLE = CounterValue;
                                    }
                                    else if (CounterGroup.EndsWith("Memory Manager"))
                                    {
                                        if (CounterName == "Total Server Memory (KB)")
                                            metrics.PerfCounter_TotalServerMemory = CounterValue;
                                        else if (CounterName == "Target Server Memory (KB)")
                                            metrics.PerfCounter_TargetServerMemory = CounterValue;
                                    }
                                    else if (CounterGroup.EndsWith("Databases"))
                                    {
                                        if (CounterName == "Data File(s) Size (KB)")
                                            metrics.PerfCounter_DataFileSize = CounterValue;
                                        else if (CounterName == "Log File(s) Size (KB)")
                                            metrics.PerfCounter_LogFileSize = CounterValue;
                                        else if (CounterName == "Log File(s) Used Size (KB)")
                                            metrics.PerfCounter_UsedLogFileSize = CounterValue;
                                    }
                                    else if (CounterGroup.EndsWith("Database Replica"))
                                    {
                                        if (CounterName == "Log Send Queue")
                                            metrics.PerfCounter_AlwaysOn_LogSendQueue = CounterValue;
                                    }
                                    else if (CounterGroup == "TOTAL-SERVER-MEMORY")
                                    {
                                        if (CounterName == "TOTAL-SERVER-MEMORY")
                                            metrics.ServerMemoryTotal = CounterValue;
                                    }

                                }
                            }
                        }
                    }

                    string queryText_5b = $@"                
                    SELECT 
                        [object_name], 
                        [counter_name], 
                        [cntr_value], 
                        ([cntr_value] - @Prev_BatchRequestsSec) / (DATEDIFF(second, @LastRefresh, GETDATE())), 
                        GETDATE()
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:SQL Statistics'
                          AND [counter_name] = 'Batch Requests/sec'
                          AND [instance_name] = ''
                    UNION ALL

                    SELECT 
                        [object_name], 
                        [counter_name], 
                        [cntr_value], 
                        ([cntr_value] - @Prev_SQLCompilationsSec) / (DATEDIFF(second, @LastRefresh, GETDATE())), 
                        GETDATE()
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:SQL Statistics'
                          AND [counter_name] = 'SQL Compilations/sec'
                          AND [instance_name] = ''

                    ";

                    using (SqlCommand command = new SqlCommand(queryText_5b, connection))
                    {
                        command.Parameters.AddWithValue("@LastRefresh", prev_metrics.PerfCounter_RefreshDateTime);
                        command.Parameters.AddWithValue("@Prev_BatchRequestsSec", prev_metrics.PerfCounter_BatchRequestsSec_Total);
                        command.Parameters.AddWithValue("@Prev_SQLCompilationsSec", prev_metrics.PerfCounter_SQLCompilationsSec_Total);

                        using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        {
                            if (reader.HasRows)
                            {
                                while (await reader.ReadAsync())
                                {

                                    string CounterGroup = reader.GetString(0).Trim();
                                    string CounterName = reader.GetString(1).Trim();
                                    long CounterValueTotal = reader.GetInt64(2);
                                    long CounterValue = reader.GetInt64(3);

                                    metrics.PerfCounter_RefreshDateTime = reader.GetDateTime(4);
                       
                                    if (CounterGroup.EndsWith("SQL Statistics"))
                                        if (CounterName == "Batch Requests/sec")
                                        {
                                            metrics.PerfCounter_BatchRequestsSec_Total = CounterValueTotal;
                                            metrics.PerfCounter_BatchRequestsSec = CounterValue;
                                        }
                                        else if (CounterName == "SQL Compilations/sec")
                                        {
                                            metrics.PerfCounter_SQLCompilationsSec_Total = CounterValueTotal;
                                            metrics.PerfCounter_SQLCompilationsSec = CounterValue;
                                        }

                                }
                            }
                        }
                    }

                    string queryText_5c = $@"
                    SELECT [counter_name], [cntr_value]
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:Buffer Manager'
                          AND [counter_name] IN ('Page reads/sec', 'Page writes/sec')
                    UNION ALL
                    SELECT [counter_name], [cntr_value]
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:Databases'
                          AND [counter_name] = 'Log Flushes/sec'
                          AND [instance_name] = '_Total'
                    UNION ALL
                    SELECT [counter_name], [cntr_value]
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:General Statistics'
                          AND [counter_name] = 'Transactions'
                    UNION ALL
                    SELECT [counter_name], [cntr_value]
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:Locks'
                          AND [counter_name] = 'Lock Waits/sec'
                          AND [instance_name] = '_Total'
                    UNION ALL
                    SELECT [counter_name], [cntr_value]
                    FROM sys.dm_os_performance_counters
                    WHERE [object_name] = '{perfCounterObjectName}:Memory Manager'
                          AND [counter_name] = 'Memory Grants Pending'
                    ";

                    using (SqlCommand command = new SqlCommand(queryText_5c, connection))
                    {
                        using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        {
                            if (reader.HasRows)
                            {
                                while (await reader.ReadAsync())
                                {
                                    string counterName = reader.GetString(0).Trim();
                                    long counterValue = reader.GetInt64(1);

                                    switch (counterName)
                                    {
                                        case "Page reads/sec":
                                            metrics.PerfCounter_PageReadsSec_Total = counterValue;
                                            break;
                                        case "Page writes/sec":
                                            metrics.PerfCounter_PageWritesSec_Total = counterValue;
                                            break;
                                        case "Log Flushes/sec":
                                            metrics.PerfCounter_LogFlushesSec_Total = counterValue;
                                            break;
                                        case "Transactions":
                                        case "Transactions/sec":
                                            metrics.PerfCounter_TransactionsSec_Total = counterValue;
                                            break;
                                        case "Lock Waits/sec":
                                            metrics.PerfCounter_LockWaitsSec_Total = counterValue;
                                            break;
                                        case "Memory Grants Pending":
                                            metrics.PerfCounter_MemoryGrantsPending = counterValue;
                                            break;
                                    }
                                }
                            }
                        }
                    }

                    double elapsedSeconds = (metrics.PerfCounter_RefreshDateTime - prev_metrics.PerfCounter_RefreshDateTime).TotalSeconds;
                    if (elapsedSeconds <= 0)
                    {
                        elapsedSeconds = 1;
                    }

                    metrics.PerfCounter_PageReadsSec = (long)((metrics.PerfCounter_PageReadsSec_Total - prev_metrics.PerfCounter_PageReadsSec_Total) / elapsedSeconds);
                    metrics.PerfCounter_PageWritesSec = (long)((metrics.PerfCounter_PageWritesSec_Total - prev_metrics.PerfCounter_PageWritesSec_Total) / elapsedSeconds);
                    metrics.PerfCounter_LogFlushesSec = (long)((metrics.PerfCounter_LogFlushesSec_Total - prev_metrics.PerfCounter_LogFlushesSec_Total) / elapsedSeconds);
                    metrics.PerfCounter_TransactionsSec = (long)((metrics.PerfCounter_TransactionsSec_Total - prev_metrics.PerfCounter_TransactionsSec_Total) / elapsedSeconds);
                    metrics.PerfCounter_LockWaitsSec = (long)((metrics.PerfCounter_LockWaitsSec_Total - prev_metrics.PerfCounter_LockWaitsSec_Total) / elapsedSeconds);

                    string queryText_6 = @"                
                    SELECT CAST(ISNULL(MIN(synchronization_health), 0) AS INT),
                           ISNULL(MAX(DATEDIFF(millisecond, last_commit_time, getdate())), 0) AS maxLatency,
                           ISNULL(SUM(drs.redo_queue_size), 0),
                           ISNULL(SUM(drs.log_send_queue_size), 0),
                           COUNT(*)
                    FROM sys.dm_hadr_database_replica_states AS drs;
                    ";

                    using (SqlCommand command = new SqlCommand(queryText_6, connection))
                    { using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        { if (reader.HasRows)
                            { while (await reader.ReadAsync())
                                {
                                    metrics.AlwaysOn_Health = reader.GetInt32(0);
                                    metrics.AlwaysOn_MaxLatency = reader.GetInt32(1);
                                    metrics.AlwaysOn_TotalRedoQueueSize = reader.GetInt64(2);
                                    metrics.AlwaysOn_TotalLogSentQueueSize = reader.GetInt64(3);

                                    metrics.AlwaysOn_Exists = (reader.GetInt32(4) > 0);
                                }
                            }
                        }
                    }

                    string queryText_7 = @"                
                    SELECT 
                            (SELECT isnull(count(*), 0) FROM sys.databases WHERE database_id > 4) AS CountUserDatabasesTotal,
                            (SELECT isnull(count(*), 0) FROM sys.databases WHERE database_id > 4
                                                                              AND user_access_desc = 'MULTI_USER'
                                                                              AND state_desc = 'ONLINE') AS CountUserDatabasesOkay;
                    ";

                    using (SqlCommand command = new SqlCommand(queryText_7, connection))
                    {
                        using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        {
                            if (reader.HasRows)
                            {
                                while (await reader.ReadAsync())
                                {
                                    metrics.CountUserDatabasesTotal = reader.GetInt32(0);
                                    metrics.CountUserDatabasesOkay = reader.GetInt32(1);
                                }
                            }
                        }
                    }


                    string queryText_71 = @"
                    SELECT TOP 20 
	                    wait_type,
	                    wait_time_ms / 1000.0 AS [WaitS]
	                    --ROW_NUMBER() OVER (ORDER BY wait_time_ms DESC) AS [RowNum]
                    FROM sys.dm_os_wait_stats WITH (NOLOCK)
                    WHERE [wait_type] NOT IN (N'BROKER_EVENTHANDLER', N'BROKER_RECEIVE_WAITFOR', N'BROKER_TASK_STOP', N'BROKER_TO_FLUSH', N'BROKER_TRANSMITTER',
                                            N'CHECKPOINT_QUEUE', N'CHKPT', N'CLR_AUTO_EVENT', N'CLR_MANUAL_EVENT', N'CLR_SEMAPHORE', N'CXCONSUMER', N'DBMIRROR_DBM_EVENT', 
                                            N'DBMIRROR_EVENTS_QUEUE', N'DBMIRROR_WORKER_QUEUE', N'DBMIRRORING_CMD', N'DIRTY_PAGE_POLL', N'DISPATCHER_QUEUE_SEMAPHORE', N'EXECSYNC',
                                            N'FSAGENT', N'FT_IFTS_SCHEDULER_IDLE_WAIT', N'FT_IFTSHC_MUTEX', N'HADR_CLUSAPI_CALL', N'HADR_FILESTREAM_IOMGR_IOCOMPLETION',
                                            N'HADR_LOGCAPTURE_WAIT', N'HADR_NOTIFICATION_DEQUEUE', N'HADR_TIMER_TASK', N'HADR_WORK_QUEUE', N'KSOURCE_WAKEUP', N'LAZYWRITER_SLEEP',
                                            N'LOGMGR_QUEUE', N'MEMORY_ALLOCATION_EXT', N'ONDEMAND_TASK_QUEUE', N'PARALLEL_REDO_DRAIN_WORKER', N'PARALLEL_REDO_LOG_CACHE',
                                            N'PARALLEL_REDO_TRAN_LIST', N'PARALLEL_REDO_WORKER_SYNC', N'PARALLEL_REDO_WORKER_WAIT_WORK', N'PREEMPTIVE_COM_GETDATA', N'PREEMPTIVE_COM_QUERYINTERFACE',
                                            N'PREEMPTIVE_HADR_LEASE_MECHANISM', N'PREEMPTIVE_SP_SERVER_DIAGNOSTICS', N'PREEMPTIVE_OS_LIBRARYOPS', N'PREEMPTIVE_OS_COMOPS', N'PREEMPTIVE_OS_CRYPTOPS',
                                            N'PREEMPTIVE_OS_PIPEOPS', N'PREEMPTIVE_OS_AUTHENTICATIONOPS', N'PREEMPTIVE_OS_GENERICOPS', N'PREEMPTIVE_OS_VERIFYTRUST', N'PREEMPTIVE_OS_FILEOPS',
                                            N'PREEMPTIVE_OS_DEVICEOPS', N'PREEMPTIVE_OS_QUERYREGISTRY', N'PREEMPTIVE_OS_WRITEFILE', N'PREEMPTIVE_OS_WRITEFILEGATHER', N'PREEMPTIVE_XE_CALLBACKEXECUTE',
                                            N'PREEMPTIVE_XE_DISPATCHER', N'PREEMPTIVE_XE_GETTARGETSTATE', N'PREEMPTIVE_XE_SESSIONCOMMIT', N'PREEMPTIVE_XE_TARGETINIT', N'PREEMPTIVE_XE_TARGETFINALIZE',
                                            N'PWAIT_ALL_COMPONENTS_INITIALIZED', N'PWAIT_DIRECTLOGCONSUMER_GETNEXT', N'PWAIT_EXTENSIBILITY_CLEANUP_TASK', N'QDS_PERSIST_TASK_MAIN_LOOP_SLEEP',
                                            N'QDS_ASYNC_QUEUE', N'QDS_CLEANUP_STALE_QUERIES_TASK_MAIN_LOOP_SLEEP', N'REQUEST_FOR_DEADLOCK_SEARCH', N'RESOURCE_QUEUE', N'SERVER_IDLE_CHECK',
                                            N'SLEEP_BPOOL_FLUSH', N'SLEEP_DBSTARTUP', N'SLEEP_DCOMSTARTUP', N'SLEEP_MASTERDBREADY', N'SLEEP_MASTERMDREADY', N'SLEEP_MASTERUPGRADED',
                                            N'SLEEP_MSDBSTARTUP', N'SLEEP_SYSTEMTASK', N'SLEEP_TASK', N'SLEEP_TEMPDBSTARTUP', N'SNI_HTTP_ACCEPT', N'SOS_WORK_DISPATCHER',
                                            N'SP_SERVER_DIAGNOSTICS_SLEEP', N'SOS_WORKER_MIGRATION', N'VDI_CLIENT_OTHER', N'SQLTRACE_BUFFER_FLUSH', N'SQLTRACE_INCREMENTAL_FLUSH_SLEEP',
                                            N'SQLTRACE_WAIT_ENTRIES', N'STARTUP_DEPENDENCY_MANAGER', N'WAIT_FOR_RESULTS', N'WAITFOR', N'WAITFOR_TASKSHUTDOWN', N'WAIT_XTP_HOST_WAIT',
                                            N'WAIT_XTP_OFFLINE_CKPT_NEW_LOG', N'WAIT_XTP_CKPT_CLOSE', N'WAIT_XTP_RECOVERY', N'XE_BUFFERMGR_ALLPROCESSED_EVENT', N'XE_DISPATCHER_JOIN',
                                            N'XE_DISPATCHER_WAIT', N'XE_LIVE_TARGET_TVF', N'XE_TIMER_EVENT')
	                    AND waiting_tasks_count > 0
	                    AND wait_time_ms > 0
                    ORDER BY wait_time_ms DESC;
                    ";

                    using (SqlCommand command = new SqlCommand(queryText_71, connection))
                    {
                        using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        {
                            if (reader.HasRows)
                            {
                                while (await reader.ReadAsync())
                                {
                                    metrics.AddWaitInfo(WaitName: reader.GetString(0), WaitSec: reader.GetDecimal(1));                                        
                                }
                            }
                        }
                    }

                    //---------------------------------------------------------------------------
                    // Slow metrics
                    int diffMinutes = 0;
                    if (metrics.SlowMetrics_RefreshDateTime != null)
                    {
                        TimeSpan SlowMetricDifference = DateTime.Now - metrics.SlowMetrics_RefreshDateTime;
                        diffMinutes = (int)SlowMetricDifference.TotalMinutes;
                    }

                    if (metrics.SlowMetrics_RefreshDateTime == null || diffMinutes >= 1)
                    {


                        // Why so complex? 
                        // Some databases have 100s of files, so the query can be slow, easier to find unique folders first
                        string queryText_8 = @"WITH AllCatalogs
                        AS (SELECT mf.database_id,
                                   file_id,
                                   REVERSE(SUBSTRING(REVERSE(physical_name), CHARINDEX('\', REVERSE(physical_name)) + 1, LEN(physical_name))) AS FolderName
                            FROM sys.master_files AS mf
                                 INNER JOIN sys.databases AS d
                                     ON mf.database_id = d.database_id
                                    AND d.state_desc = 'ONLINE'),
                         NumberedCatalogs AS (
						 	SELECT database_id,
                                   file_id,
                                   FolderName,
                                   ROW_NUMBER() OVER (PARTITION BY FolderName ORDER BY database_id, file_id) AS RN
                            FROM AllCatalogs)
                        SELECT DISTINCT vs.volume_mount_point,
                                        vs.logical_volume_name,
                                        vs.total_bytes,
                                        vs.available_bytes
                        FROM NumberedCatalogs AS f WITH (NOLOCK)
                             CROSS APPLY sys.dm_os_volume_stats(f.database_id, f.[file_id]) AS vs
                        WHERE f.RN = 1
                        ORDER BY vs.volume_mount_point
                        OPTION (RECOMPILE);
                        ";
                        using (SqlCommand command = new SqlCommand(queryText_8, connection))
                        {
                            using (SqlDataReader reader = await command.ExecuteReaderAsync())
                            {
                                if (reader.HasRows)
                                {
                                    while (await reader.ReadAsync())
                                    {
                                        metrics.AddDiskInfo(volumeMountPoint: reader.GetString(0), 
                                                volumeName: reader.GetString(1), 
                                                Capacity: reader.GetInt64(2), 
                                                FreeSpace: reader.GetInt64(3));
                                    }
                                }
                            }
                        }






                        metrics.SlowMetrics_RefreshDateTime = DateTime.Now;

                    } else
                    {
                        metrics.DisksInfo = prev_metrics.DisksInfo;
                    }

                }

                stopwatch.Stop();
                TimeSpan ts = stopwatch.Elapsed;

                metrics.ServerResponseTimeMs = (int)ts.TotalMilliseconds;
                metrics.Completed = true;

            }
            catch (Exception ex)
            {
                metrics.HasException = true;
                metrics.ExecutionException = ex.Message;
            }

            return metrics;
        }
        
    }
}
