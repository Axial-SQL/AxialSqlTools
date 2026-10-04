using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools.JobQuickView
{
    /// <summary>
    /// Uses the selected SSMS login. SQL Server Agent enforces authorization for every action.
    /// Each operation owns a fresh connection; nothing runs on the editor's query connection.
    /// </summary>
    internal sealed class JobQuickViewService
    {
        private const int CommandTimeoutSeconds = 15;
        private readonly Func<SqlConnection> _connectionFactory;

        public JobQuickViewService(Func<SqlConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        public JobQuickViewService(string connectionString)
        {
            var builder = new SqlConnectionStringBuilder(connectionString) { InitialCatalog = "msdb" };
            string msdbConnectionString = builder.ConnectionString;
            _connectionFactory = () => new SqlConnection(msdbConnectionString);
        }

        public Task<JobQuickViewSnapshot> LoadAsync(string jobName, CancellationToken cancellationToken)
        {
            if (string.IsNullOrEmpty(jobName))
                throw new ArgumentException("A job name is required.", nameof(jobName));
            return LoadCoreAsync(null, jobName, cancellationToken, executionOnly: false, includeHistory: true);
        }

        public Task<JobQuickViewSnapshot> LoadAsync(Guid jobId, CancellationToken cancellationToken)
        {
            ValidateJobId(jobId);
            return LoadCoreAsync(jobId, null, cancellationToken, executionOnly: false, includeHistory: false);
        }

        public Task<JobQuickViewSnapshot> LoadExecutionAsync(Guid jobId, CancellationToken cancellationToken)
        {
            ValidateJobId(jobId);
            return LoadCoreAsync(jobId, null, cancellationToken, executionOnly: true, includeHistory: true);
        }

        private async Task<JobQuickViewSnapshot> LoadCoreAsync(Guid? jobId, string jobName, CancellationToken token, bool executionOnly, bool includeHistory)
        {
            using (var connection = await OpenAsync(token).ConfigureAwait(false))
            {
                var snapshot = new JobQuickViewSnapshot();
                using (var command = Procedure(connection, "dbo.sp_help_job"))
                {
                    if (jobId.HasValue)
                        command.Parameters.Add("@job_id", SqlDbType.UniqueIdentifier).Value = jobId.Value;
                    else
                        command.Parameters.Add("@job_name", SqlDbType.NVarChar, 128).Value = jobName;
                    command.Parameters.Add("@job_aspect", SqlDbType.VarChar, 9).Value = "JOB";
                    using (var reader = await command.ExecuteReaderAsync(token).ConfigureAwait(false))
                    {
                        if (!await reader.ReadAsync(token).ConfigureAwait(false))
                            throw new InvalidOperationException("The job no longer exists or the current login cannot view it.");
                        snapshot.JobId = (Guid)reader["job_id"];
                        snapshot.Name = Text(reader, "name");
                        snapshot.Owner = Text(reader, "owner");
                        snapshot.Category = Text(reader, "category");
                        snapshot.Description = Text(reader, "description");
                        snapshot.CreatedAt = ServerDateTime(reader, "date_created");
                        snapshot.ModifiedAt = ServerDateTime(reader, "date_modified");
                        snapshot.IsEnabled = Number(reader, "enabled") != 0;
                        snapshot.StartStepId = Number(reader, "start_step_id");
                        int status = Number(reader, "current_execution_status");
                        snapshot.IsRunning = JobQuickViewValue.IsRunning(status);
                        snapshot.ExecutionStatus = JobQuickViewValue.ExecutionStatus(status);
                        snapshot.LastRunStarted = JobQuickViewValue.DateTimeFromAgent(Number(reader, "last_run_date"), Number(reader, "last_run_time"));
                        snapshot.LastRunOutcome = snapshot.LastRunStarted.HasValue
                            ? JobQuickViewValue.Outcome(Number(reader, "last_run_outcome")) : "No history";
                        snapshot.LastRunMessage = string.Empty;
                        snapshot.NextRun = snapshot.IsEnabled
                            ? JobQuickViewValue.DateTimeFromAgent(Number(reader, "next_run_date"), Number(reader, "next_run_time")) : null;
                    }
                }

                var warnings = new List<string>();
                if (!executionOnly)
                {
                    // Read nvarchar(max) directly: sp_help_job's documented step result can be truncated.
                    snapshot.Steps = await ReadStepsAsync(connection, snapshot.JobId, token).ConfigureAwait(false);
                    try
                    {
                        snapshot.Schedules = await ReadSchedulesAsync(connection, snapshot.JobId, token).ConfigureAwait(false);
                    }
                    catch (SqlException ex) when (!token.IsCancellationRequested)
                    {
                        snapshot.SchedulesLoaded = false;
                        warnings.Add("Schedules could not be loaded: " + ex.Message);
                    }
                }
                // History loads on first open and through its dedicated refresh only.
                if (includeHistory)
                {
                    try
                    {
                        await ReadHistoryAsync(connection, snapshot, token).ConfigureAwait(false);
                    }
                    catch (SqlException ex) when (!token.IsCancellationRequested)
                    {
                        snapshot.HistoryLoaded = false;
                        warnings.Add("Execution history could not be loaded: " + ex.Message);
                    }
                }
                try
                {
                    await ReadActivityAsync(connection, snapshot, token).ConfigureAwait(false);
                }
                catch (SqlException ex) when (!token.IsCancellationRequested)
                {
                    warnings.Add("Execution timing could not be loaded: " + ex.Message);
                }
                if (!snapshot.IsRunning.HasValue)
                    warnings.Add("The current execution state is unavailable. SQL Server Agent may be stopped.");
                snapshot.ActivityWarning = string.Join(Environment.NewLine, warnings);
                return snapshot;
            }
        }

        private static async Task<IReadOnlyList<JobQuickViewStep>> ReadStepsAsync(SqlConnection connection, Guid jobId, CancellationToken token)
        {
            const string sql = @"
SELECT s.step_id, s.step_uid, s.step_name, s.subsystem, s.database_name,
       s.database_user_name, s.proxy_id, s.command,
       s.on_success_action, s.on_success_step_id, s.on_fail_action, s.on_fail_step_id
FROM dbo.sysjobsteps AS s
INNER JOIN dbo.sysjobs_view AS j ON j.job_id = s.job_id
WHERE s.job_id = @job_id
ORDER BY s.step_id;";
            var steps = new List<JobQuickViewStep>();
            using (var command = Query(connection, sql, jobId))
            using (var reader = await command.ExecuteReaderAsync(token).ConfigureAwait(false))
            {
                while (await reader.ReadAsync(token).ConfigureAwait(false))
                {
                    steps.Add(new JobQuickViewStep
                    {
                        StepId = Number(reader, "step_id"),
                        StepUid = (Guid)reader["step_uid"],
                        Name = Text(reader, "step_name"),
                        Subsystem = Text(reader, "subsystem"),
                        DatabaseName = Text(reader, "database_name"),
                        DatabaseUserName = Text(reader, "database_user_name"),
                        ProxyId = Number(reader, "proxy_id"),
                        Command = Text(reader, "command"),
                        SuccessAction = JobQuickViewValue.StepAction(Number(reader, "on_success_action"), Number(reader, "on_success_step_id")),
                        FailureAction = JobQuickViewValue.StepAction(Number(reader, "on_fail_action"), Number(reader, "on_fail_step_id"))
                    });
                }
            }
            return steps;
        }

        private static async Task<IReadOnlyList<JobQuickViewSchedule>> ReadSchedulesAsync(SqlConnection connection, Guid jobId, CancellationToken token)
        {
            var schedules = new List<JobQuickViewSchedule>();
            using (var command = Procedure(connection, "dbo.sp_help_jobschedule", jobId))
            {
                command.Parameters.Add("@include_description", SqlDbType.Bit).Value = true;
                using (var reader = await command.ExecuteReaderAsync(token).ConfigureAwait(false))
                {
                    while (await reader.ReadAsync(token).ConfigureAwait(false))
                    {
                        bool enabled = Number(reader, "enabled") != 0;
                        schedules.Add(new JobQuickViewSchedule
                        {
                            ScheduleId = Number(reader, "schedule_id"),
                            Name = ReadScheduleName(reader),
                            IsEnabled = enabled,
                            Description = Text(reader, "schedule_description"),
                            NextRun = enabled ? JobQuickViewValue.DateTimeFromAgent(Number(reader, "next_run_date"), Number(reader, "next_run_time")) : null
                        });
                    }
                }
            }
            return schedules;
        }

        private static async Task ReadHistoryAsync(SqlConnection connection, JobQuickViewSnapshot snapshot, CancellationToken token)
        {
            var runs = new List<JobQuickViewHistoryRun>();
            using (var command = Query(connection, HistorySql, snapshot.JobId))
            {
                command.Parameters.Add("@summary_count", SqlDbType.Int).Value = JobQuickViewHistory.MaximumRuns;
                using (var reader = await command.ExecuteReaderAsync(token).ConfigureAwait(false))
                {
                    while (await reader.ReadAsync(token).ConfigureAwait(false))
                    {
                        int status = Number(reader, "run_status");
                        runs.Add(new JobQuickViewHistoryRun
                        {
                            InstanceId = Number(reader, "instance_id"),
                            Server = Text(reader, "server"),
                            RunStatus = status,
                            StartedAt = JobQuickViewValue.DateTimeFromAgent(Number(reader, "run_date"), Number(reader, "run_time")),
                            Duration = JobQuickViewValue.DurationFromAgent(Number(reader, "run_duration")),
                            Outcome = JobQuickViewValue.Outcome(status)
                        });
                    }
                }
            }
            snapshot.History = runs;
            if (runs.Count > 0)
            {
                snapshot.LastRunStarted = runs[0].StartedAt;
                snapshot.LastRunDuration = runs[0].Duration;
                snapshot.LastRunOutcome = runs[0].Outcome;
            }
        }

        // The list never reads messages or step rows. Details require an explicit selection.
        internal const string HistorySql = @"
SELECT TOP (@summary_count) h.instance_id, h.[server],
       h.run_date, h.run_time, h.run_duration, h.run_status
FROM dbo.sysjobhistory AS h
INNER JOIN dbo.sysjobs_view AS j ON j.job_id = h.job_id
WHERE h.job_id = @job_id AND h.step_id = 0 AND h.run_status IN (0, 1, 3)
ORDER BY h.instance_id DESC;";

        public async Task<JobQuickViewHistoryRun> LoadHistoryDetailsAsync(Guid jobId, int instanceId, CancellationToken token)
        {
            ValidateJobId(jobId);
            if (instanceId <= 0) throw new ArgumentOutOfRangeException(nameof(instanceId));
            using (var connection = await OpenAsync(token).ConfigureAwait(false))
            using (var command = Query(connection, HistoryDetailsSql, jobId))
            {
                command.Parameters.Add("@instance_id", SqlDbType.Int).Value = instanceId;
                using (var reader = await command.ExecuteReaderAsync(token).ConfigureAwait(false))
                {
                    if (!await reader.ReadAsync(token).ConfigureAwait(false))
                        throw new InvalidOperationException("This execution is no longer available. Its history may have been purged. Use Refresh history to update the list.");
                    var summary = ReadHistoryItem(reader);
                    var rows = new List<JobQuickViewHistoryItem>();
                    int previousId = Number(reader, "previous_id");
                    if (previousId > 0)
                        rows.Add(new JobQuickViewHistoryItem { InstanceId = previousId, Server = summary.Server, RunStatus = 1 });
                    rows.Add(summary);
                    if (await reader.NextResultAsync(token).ConfigureAwait(false))
                        while (await reader.ReadAsync(token).ConfigureAwait(false))
                            rows.Add(ReadHistoryItem(reader));
                    token.ThrowIfCancellationRequested();
                    return JobQuickViewHistory.Group(rows, 1)[0];
                }
            }
        }

        private static JobQuickViewHistoryItem ReadHistoryItem(SqlDataReader reader)
        {
            int status = Number(reader, "run_status");
            return new JobQuickViewHistoryItem
            {
                InstanceId = Number(reader, "instance_id"),
                StepId = Number(reader, "step_id"),
                StepName = Text(reader, "step_name"),
                Server = Text(reader, "server"),
                RunStatus = status,
                StartedAt = JobQuickViewValue.DateTimeFromAgent(Number(reader, "run_date"), Number(reader, "run_time")),
                Duration = JobQuickViewValue.DurationFromAgent(Number(reader, "run_duration")),
                Outcome = JobQuickViewValue.Outcome(status),
                Message = Text(reader, "message"),
                SqlSeverity = Number(reader, "sql_severity"),
                SqlMessageId = Number(reader, "sql_message_id")
            };
        }

        // Capture the selected summary and find its preceding boundary on the same target server.
        // A completed run is immutable; concurrent completions cannot move this bounded window.
        internal const string HistoryDetailsSql = @"
SET NOCOUNT ON;
DECLARE @run TABLE
(
    instance_id int, step_id int, step_name nvarchar(128), [server] nvarchar(128),
    run_date int, run_time int, run_duration int, run_status int,
    message nvarchar(4000), sql_severity int, sql_message_id int
);
INSERT INTO @run
SELECT h.instance_id, h.step_id, h.step_name, h.[server],
       h.run_date, h.run_time, h.run_duration, h.run_status,
       h.message, h.sql_severity, h.sql_message_id
FROM dbo.sysjobhistory AS h
INNER JOIN dbo.sysjobs_view AS j ON j.job_id = h.job_id
WHERE h.job_id = @job_id AND h.instance_id = @instance_id
  AND h.step_id = 0 AND h.run_status IN (0, 1, 3);

DECLARE @previous_id int;
SELECT @previous_id = MAX(h.instance_id)
FROM dbo.sysjobhistory AS h
CROSS JOIN @run AS r
WHERE h.job_id = @job_id AND h.step_id = 0 AND h.run_status IN (0, 1, 3)
  AND h.instance_id < r.instance_id
  AND ISNULL(h.[server], N'') = ISNULL(r.[server], N'');

SELECT r.*, ISNULL(@previous_id, 0) AS previous_id FROM @run AS r;
SELECT h.instance_id, h.step_id, h.step_name, h.[server],
       h.run_date, h.run_time, h.run_duration, h.run_status,
       h.message, h.sql_severity, h.sql_message_id
FROM dbo.sysjobhistory AS h
INNER JOIN dbo.sysjobs_view AS j ON j.job_id = h.job_id
CROSS JOIN @run AS r
WHERE h.job_id = @job_id AND h.step_id > 0
  AND h.instance_id > ISNULL(@previous_id, 0) AND h.instance_id < r.instance_id
  AND ISNULL(h.[server], N'') = ISNULL(r.[server], N'')
ORDER BY h.instance_id;";

        private static async Task ReadActivityAsync(SqlConnection connection, JobQuickViewSnapshot snapshot, CancellationToken token)
        {
            // Omitting @session_id asks Agent for its latest session. Old sessions can contain
            // unfinished rows after a crash; only the live sp_help_job state determines IsRunning.
            using (var command = Procedure(connection, "dbo.sp_help_jobactivity", snapshot.JobId))
            using (var reader = await command.ExecuteReaderAsync(token).ConfigureAwait(false))
            {
                if (await reader.ReadAsync(token).ConfigureAwait(false)
                    && snapshot.IsRunning == true
                    && reader["stop_execution_date"] == DBNull.Value
                    && reader["start_execution_date"] != DBNull.Value)
                {
                    snapshot.RunningSince = DateTime.SpecifyKind((DateTime)reader["start_execution_date"], DateTimeKind.Unspecified);
                }
            }
        }

        public Task SetEnabledAsync(Guid jobId, bool enabled, CancellationToken cancellationToken)
        {
            return ExecuteActionAsync("dbo.sp_update_job", jobId,
                command => command.Parameters.Add("@enabled", SqlDbType.TinyInt).Value = enabled ? 1 : 0,
                cancellationToken);
        }

        public Task StartAsync(Guid jobId, CancellationToken cancellationToken)
        {
            return ExecuteActionAsync("dbo.sp_start_job", jobId, null, cancellationToken);
        }

        public Task StopAsync(Guid jobId, CancellationToken cancellationToken)
        {
            return ExecuteActionAsync("dbo.sp_stop_job", jobId, null, cancellationToken);
        }

        private async Task ExecuteActionAsync(string procedure, Guid jobId, Action<SqlCommand> addParameters, CancellationToken token)
        {
            ValidateJobId(jobId);
            using (var connection = await OpenAsync(token).ConfigureAwait(false))
            using (var command = Procedure(connection, procedure, jobId))
            {
                addParameters?.Invoke(command);
                var result = command.Parameters.Add("@return_value", SqlDbType.Int);
                result.Direction = ParameterDirection.ReturnValue;
                await command.ExecuteNonQueryAsync(token).ConfigureAwait(false);
                if (result.Value != DBNull.Value && Convert.ToInt32(result.Value, CultureInfo.InvariantCulture) != 0)
                    throw new InvalidOperationException("SQL Server Agent did not accept the requested job action. Refresh the job and try again.");
            }
        }

        public async Task SaveStepCommandAsync(Guid jobId, JobQuickViewStep originalStep, string commandText, CancellationToken cancellationToken)
        {
            ValidateJobId(jobId);
            if (originalStep == null)
                throw new ArgumentNullException(nameof(originalStep));
            if (originalStep.StepId <= 0 || originalStep.StepUid == Guid.Empty)
                throw new ArgumentException("The original job step identity is required.", nameof(originalStep));
            if (commandText == null)
                throw new ArgumentNullException(nameof(commandText));

            using (var connection = await OpenAsync(cancellationToken).ConfigureAwait(false))
            using (var command = Query(connection, SaveCommandSql, jobId))
            {
                command.Parameters.Add("@step_id", SqlDbType.Int).Value = originalStep.StepId;
                command.Parameters.Add("@step_uid", SqlDbType.UniqueIdentifier).Value = originalStep.StepUid;
                command.Parameters.Add("@original_command", SqlDbType.NVarChar, -1).Value = originalStep.Command ?? string.Empty;
                command.Parameters.Add("@command", SqlDbType.NVarChar, -1).Value = commandText;
                command.Parameters.Add("@original_subsystem", SqlDbType.NVarChar, 40).Value = originalStep.Subsystem ?? string.Empty;
                command.Parameters.Add("@original_database", SqlDbType.NVarChar, 128).Value = originalStep.DatabaseName ?? string.Empty;
                command.Parameters.Add("@original_database_user", SqlDbType.NVarChar, 128).Value = originalStep.DatabaseUserName ?? string.Empty;
                command.Parameters.Add("@original_proxy", SqlDbType.Int).Value = originalStep.ProxyId;
                try
                {
                    await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
                }
                catch (SqlException ex) when (ex.Number == 50001)
                {
                    throw new JobCommandConflictException(
                        "This step was changed, moved, or deleted outside Quick Manage. Your edit was not saved. Copy your changes, then reload the step before trying again.", ex);
                }
            }
        }

        // The row lock spans comparison and the supported update procedure. Binary comparison
        // plus DATALENGTH detects case, whitespace, and trailing-space changes on any collation.
        // The stable UID prevents an inserted/reordered step at the same ordinal being overwritten.
        // Execution context is also checked; no other step or job settings are sent to the update.
        internal const string SaveCommandSql = @"
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    IF NOT EXISTS
    (
        SELECT 1
        FROM dbo.sysjobsteps AS s WITH (UPDLOCK, HOLDLOCK)
        INNER JOIN dbo.sysjobs_view AS j ON j.job_id = s.job_id
        WHERE s.job_id = @job_id AND s.step_id = @step_id AND s.step_uid = @step_uid
          AND DATALENGTH(ISNULL(s.command, N'')) = DATALENGTH(@original_command)
          AND CONVERT(varbinary(max), ISNULL(s.command, N'')) = CONVERT(varbinary(max), @original_command)
          AND CONVERT(varbinary(80), s.subsystem) = CONVERT(varbinary(80), @original_subsystem)
          AND CONVERT(varbinary(256), ISNULL(s.database_name, N'')) = CONVERT(varbinary(256), @original_database)
          AND CONVERT(varbinary(256), ISNULL(s.database_user_name, N'')) = CONVERT(varbinary(256), @original_database_user)
          AND ISNULL(s.proxy_id, 0) = @original_proxy
    )
        THROW 50001, 'The job step changed after it was loaded.', 1;

    DECLARE @result int;
    EXEC @result = dbo.sp_update_jobstep @job_id = @job_id, @step_id = @step_id, @command = @command;
    IF @result <> 0
        THROW 50002, 'SQL Server Agent did not save the step command.', 1;
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    THROW;
END CATCH;";

        private async Task<SqlConnection> OpenAsync(CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            // SSMS may acquire an authentication token while creating the connection object.
            // Keep that work off the WPF thread as well as the asynchronous SQL operations.
            var connection = await Task.Run(_connectionFactory, token).ConfigureAwait(false);
            if (connection == null)
                throw new InvalidOperationException("The selected SSMS connection could not be created.");
            try
            {
                await connection.OpenAsync(token).ConfigureAwait(false);
                if (!string.Equals(connection.Database, "msdb", StringComparison.OrdinalIgnoreCase))
                {
                    using (var command = new SqlCommand("USE [msdb];", connection) { CommandTimeout = CommandTimeoutSeconds })
                        await command.ExecuteNonQueryAsync(token).ConfigureAwait(false);
                }
                return connection;
            }
            catch
            {
                connection.Dispose();
                throw;
            }
        }

        private static SqlCommand Procedure(SqlConnection connection, string name, Guid? jobId = null)
        {
            var command = new SqlCommand(name, connection) { CommandType = CommandType.StoredProcedure, CommandTimeout = CommandTimeoutSeconds };
            if (jobId.HasValue)
                command.Parameters.Add("@job_id", SqlDbType.UniqueIdentifier).Value = jobId.Value;
            return command;
        }

        private static SqlCommand Query(SqlConnection connection, string sql, Guid jobId)
        {
            var command = new SqlCommand(sql, connection) { CommandTimeout = CommandTimeoutSeconds };
            command.Parameters.Add("@job_id", SqlDbType.UniqueIdentifier).Value = jobId;
            return command;
        }

        private static int Number(SqlDataReader reader, string name)
        {
            return reader[name] == DBNull.Value ? 0 : Convert.ToInt32(reader[name], CultureInfo.InvariantCulture);
        }

        private static string Text(SqlDataReader reader, string name)
        {
            return reader[name] == DBNull.Value ? string.Empty : Convert.ToString(reader[name], CultureInfo.InvariantCulture);
        }

        internal static string ReadScheduleName(IDataRecord reader)
        {
            // Handle the returned 'name' column as well as the documented 'schedule_name'.
            for (int index = 0; index < reader.FieldCount; index++)
            {
                if (string.Equals(reader.GetName(index), "name", StringComparison.OrdinalIgnoreCase))
                    return reader.IsDBNull(index) ? string.Empty : reader.GetString(index);
            }
            object value = reader["schedule_name"];
            return value == DBNull.Value ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static DateTime? ServerDateTime(SqlDataReader reader, string name)
        {
            return reader[name] == DBNull.Value
                ? (DateTime?)null
                : DateTime.SpecifyKind((DateTime)reader[name], DateTimeKind.Unspecified);
        }

        private static void ValidateJobId(Guid jobId)
        {
            if (jobId == Guid.Empty)
                throw new ArgumentException("A job ID is required.", nameof(jobId));
        }
    }
}
