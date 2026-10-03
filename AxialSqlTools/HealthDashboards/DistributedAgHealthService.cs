namespace AxialSqlTools
{
    using Microsoft.Data.SqlClient;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>Reads only the selected connection. Endpoint URLs are metadata, not SQL connection addresses.</summary>
    public static class DistributedAgHealthService
    {
        private const int QueryTimeoutSeconds = 20;
        private const string PreflightSql = @"
SELECT CONVERT(nvarchar(256), SERVERPROPERTY('ServerName')) AS ServerName,
       CONVERT(nvarchar(128), SERVERPROPERTY('ProductVersion')) AS ServerVersion,
       CONVERT(int, SERVERPROPERTY('ProductMajorVersion')) AS MajorVersion,
       CONVERT(int, SERVERPROPERTY('IsHadrEnabled')) AS IsHadrEnabled,
       HAS_PERMS_BY_NAME(NULL, NULL, 'VIEW ANY DEFINITION') AS CanViewDefinition,
       HAS_PERMS_BY_NAME(NULL, NULL, 'VIEW SERVER STATE') AS CanViewState,
       HAS_PERMS_BY_NAME(NULL, NULL, 'VIEW SERVER PERFORMANCE STATE') AS CanViewPerformanceState;";

        // The DAG replicas are underlying AG names. Do not join member AG database rows into this result.
        // LEFT JOIN preserves configured members when their state is not visible from the current replica.
        private const string HealthSql = @"
DECLARE @GroupId uniqueidentifier =
    (SELECT group_id FROM sys.availability_groups WHERE is_distributed = 1 AND name = @GroupName);
SELECT name FROM sys.availability_groups WHERE group_id = @GroupId;

SELECT ar.replica_server_name AS MemberName, ar.endpoint_url, ar.availability_mode_desc,
       ars.replica_id AS StateReplicaId, ars.is_local, ars.role, ars.connected_state,
       ars.connected_state_desc, ars.operational_state, ars.synchronization_health,
       ars.last_connect_error_number, ars.last_connect_error_description, ars.last_connect_error_timestamp
FROM sys.availability_replicas AS ar
LEFT JOIN sys.dm_hadr_availability_replica_states AS ars
  ON ars.group_id = ar.group_id AND ars.replica_id = ar.replica_id
WHERE ar.group_id = @GroupId
ORDER BY ar.replica_server_name;

SELECT COALESCE(adc.database_name, dbs.name, CONVERT(nvarchar(36), drs.group_database_id)) AS DatabaseName,
       ar.replica_server_name AS MemberName,
       drs.synchronization_state, drs.synchronization_state_desc, drs.synchronization_health,
       drs.database_state, drs.database_state_desc, drs.is_suspended, drs.suspend_reason_desc,
       drs.log_send_queue_size, drs.log_send_rate, drs.redo_queue_size, drs.redo_rate,
       CASE WHEN drs.is_suspended = 1 THEN NULL ELSE drs.secondary_lag_seconds END AS LagSeconds,
       drs.last_hardened_time
FROM sys.dm_hadr_database_replica_states AS drs
JOIN sys.availability_replicas AS ar
  ON ar.group_id = drs.group_id AND ar.replica_id = drs.replica_id
LEFT JOIN sys.availability_databases_cluster AS adc
  ON adc.group_id = drs.group_id AND adc.group_database_id = drs.group_database_id
LEFT JOIN sys.databases AS dbs ON dbs.database_id = drs.database_id
WHERE drs.group_id = @GroupId
ORDER BY DatabaseName, MemberName;";

        public static async Task<IReadOnlyList<string>> ListGroupsAsync(string connectionString, CancellationToken token)
        {
            using (var connection = new SqlConnection(connectionString))
            {
                await connection.OpenAsync(token).ConfigureAwait(false);
                await ReadServerAsync(connection, token).ConfigureAwait(false);
                using (var command = CreateCommand(connection,
                    "SELECT name FROM sys.availability_groups WHERE is_distributed = 1 ORDER BY name;"))
                using (var reader = await command.ExecuteReaderAsync(token).ConfigureAwait(false))
                {
                    var groups = new List<string>();
                    while (await reader.ReadAsync(token).ConfigureAwait(false)) groups.Add(reader.GetString(0));
                    return groups;
                }
            }
        }

        public static async Task<DistributedAgHealthSnapshot> LoadAsync(
            string connectionString, string groupName, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(groupName))
                throw new ArgumentException("Select a distributed availability group.", nameof(groupName));

            using (var connection = new SqlConnection(connectionString))
            {
                await connection.OpenAsync(token).ConfigureAwait(false);
                var snapshot = await ReadServerAsync(connection, token).ConfigureAwait(false);
                snapshot.GroupName = groupName;
                using (var command = CreateCommand(connection, HealthSql))
                {
                    command.Parameters.Add("@GroupName", SqlDbType.NVarChar, 128).Value = groupName;
                    using (var reader = await command.ExecuteReaderAsync(token).ConfigureAwait(false))
                    {
                        if (!await reader.ReadAsync(token).ConfigureAwait(false))
                            throw new InvalidOperationException("The selected distributed AG is not visible on this instance. " +
                                "Connect to the primary replica of one of its underlying availability groups and refresh. " +
                                "The group may also have been removed or renamed.");
                        snapshot.GroupName = reader.GetString(0);
                        await reader.NextResultAsync(token).ConfigureAwait(false);
                        while (await reader.ReadAsync(token).ConfigureAwait(false))
                            snapshot.Members.Add(ReadMember(reader));
                        await reader.NextResultAsync(token).ConfigureAwait(false);
                        while (await reader.ReadAsync(token).ConfigureAwait(false))
                            snapshot.Databases.Add(ReadDatabase(reader));
                    }
                }
                snapshot.CollectedAt = DateTime.Now;
                DistributedAgHealthEvaluator.Evaluate(snapshot);
                return snapshot;
            }
        }

        private static async Task<DistributedAgHealthSnapshot> ReadServerAsync(SqlConnection connection, CancellationToken token)
        {
            using (var command = CreateCommand(connection, PreflightSql))
            using (var reader = await command.ExecuteReaderAsync(token).ConfigureAwait(false))
            {
                await reader.ReadAsync(token).ConfigureAwait(false);
                int majorVersion = Number<int>(reader, "MajorVersion") ?? 0;
                if (majorVersion < 13)
                    throw new InvalidOperationException("Distributed availability groups require SQL Server 2016 or later.");
                if (Number<int>(reader, "IsHadrEnabled") != 1)
                    throw new InvalidOperationException("Always On availability groups are not enabled on this instance.");

                string statePermission = majorVersion >= 16 ? "VIEW SERVER PERFORMANCE STATE" : "VIEW SERVER STATE";
                string stateColumn = majorVersion >= 16 ? "CanViewPerformanceState" : "CanViewState";
                if (Number<int>(reader, stateColumn) != 1 || Number<int>(reader, "CanViewDefinition") != 1)
                    throw new InvalidOperationException("This dashboard requires " + statePermission +
                        " and VIEW ANY DEFINITION on the connected SQL Server instance. Ask a server administrator to grant the missing permissions.");

                return new DistributedAgHealthSnapshot
                {
                    ServerName = Text(reader, "ServerName"),
                    ServerVersion = Text(reader, "ServerVersion")
                };
            }
        }

        private static DistributedAgMemberHealth ReadMember(SqlDataReader reader)
        {
            int? role = Number<int>(reader, "role");
            int? health = Number<int>(reader, "synchronization_health");
            int? errorNumber = Number<int>(reader, "last_connect_error_number");
            DateTime? errorTime = Number<DateTime>(reader, "last_connect_error_timestamp");
            string error = errorNumber.HasValue && errorNumber.Value != 0
                ? "Last reported error " + errorNumber.Value.ToString(CultureInfo.CurrentCulture) +
                  (errorTime.HasValue ? " at " + errorTime.Value.ToString("g", CultureInfo.CurrentCulture) + " (server time)" : "") +
                  ": " + Text(reader, "last_connect_error_description")
                : "None reported";

            return new DistributedAgMemberHealth
            {
                Name = Text(reader, "MemberName"),
                Role = role == 1 ? "Global primary" : role == 2 ? "Forwarder" : role == 0 ? "Resolving" : "Not visible",
                AvailabilityMode = Text(reader, "availability_mode_desc"),
                ConnectionState = Text(reader, "connected_state_desc"),
                SynchronizationHealth = HealthText(health),
                EndpointUrl = Text(reader, "endpoint_url"),
                LastError = error,
                IsLocal = Number<bool>(reader, "is_local"),
                HasState = !reader.IsDBNull(reader.GetOrdinal("StateReplicaId")),
                RoleCode = role,
                ConnectedStateCode = Number<int>(reader, "connected_state"),
                OperationalStateCode = Number<int>(reader, "operational_state"),
                SynchronizationHealthCode = health
            };
        }

        private static DistributedAgDatabaseHealth ReadDatabase(SqlDataReader reader)
        {
            int? health = Number<int>(reader, "synchronization_health");
            return new DistributedAgDatabaseHealth
            {
                DatabaseName = Text(reader, "DatabaseName"),
                MemberName = Text(reader, "MemberName"),
                SynchronizationState = Text(reader, "synchronization_state_desc"),
                SynchronizationHealth = HealthText(health),
                DatabaseState = Text(reader, "database_state_desc"),
                IsSuspended = Number<bool>(reader, "is_suspended"),
                SuspendReason = Text(reader, "suspend_reason_desc", ""),
                SendQueueKb = Number<long>(reader, "log_send_queue_size"),
                SendRateKb = Number<long>(reader, "log_send_rate"),
                RedoQueueKb = Number<long>(reader, "redo_queue_size"),
                RedoRateKb = Number<long>(reader, "redo_rate"),
                LagSeconds = Number<long>(reader, "LagSeconds"),
                LastHardenedTime = Number<DateTime>(reader, "last_hardened_time"),
                SynchronizationStateCode = Number<int>(reader, "synchronization_state"),
                SynchronizationHealthCode = health,
                DatabaseStateCode = Number<int>(reader, "database_state")
            };
        }

        private static SqlCommand CreateCommand(SqlConnection connection, string sql)
        {
            return new SqlCommand(sql, connection) { CommandTimeout = QueryTimeoutSeconds };
        }

        private static string HealthText(int? health)
        {
            return health == 2 ? "Healthy" : health == 1 ? "Partially healthy" : health == 0 ? "Not healthy" : "Not visible";
        }

        private static string Text(SqlDataReader reader, string column, string fallback = "Not visible")
        {
            int ordinal = reader.GetOrdinal(column);
            return reader.IsDBNull(ordinal) ? fallback : reader.GetString(ordinal);
        }

        private static T? Number<T>(SqlDataReader reader, string column) where T : struct
        {
            int ordinal = reader.GetOrdinal(column);
            return reader.IsDBNull(ordinal) ? (T?)null : (T)Convert.ChangeType(reader.GetValue(ordinal), typeof(T), CultureInfo.InvariantCulture);
        }
    }
}
