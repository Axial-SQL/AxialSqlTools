namespace AxialSqlTools
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    public enum DistributedAgHealthLevel
    {
        Healthy,
        Warning,
        Critical,
        Unknown
    }

    public sealed class DistributedAgHealthSnapshot
    {
        public string ServerName { get; set; }
        public string GroupName { get; set; }
        public DateTime CollectedAt { get; set; }
        public string ServerVersion { get; set; }
        public List<DistributedAgMemberHealth> Members { get; } = new List<DistributedAgMemberHealth>();
        public List<DistributedAgDatabaseHealth> Databases { get; } = new List<DistributedAgDatabaseHealth>();
        public string StatusText { get; set; }
        public DistributedAgHealthLevel StatusLevel { get; set; } = DistributedAgHealthLevel.Unknown;
        public string StatusDetail { get; set; }
        public bool IsGlobalPrimaryObservation => Members.Any(member => member.IsLocal == true && member.RoleCode == 1);
    }

    public sealed class DistributedAgMemberHealth
    {
        public string Name { get; set; }
        public string Role { get; set; }
        public string AvailabilityMode { get; set; }
        public string ConnectionState { get; set; }
        public string SynchronizationHealth { get; set; }
        public string EndpointUrl { get; set; }
        public string LastError { get; set; }
        public bool? IsLocal { get; set; }
        public string DisplayHealth => StatusLevel == DistributedAgHealthLevel.Critical
            ? (ConnectedStateCode == 0 && RoleCode != 1 ? "Disconnected" : "Needs attention")
            : StatusLevel == DistributedAgHealthLevel.Warning ? "Partially healthy"
            : StatusLevel == DistributedAgHealthLevel.Healthy ? "Healthy" : "Not visible";
        public bool HasState { get; set; }
        public int? RoleCode { get; set; }
        public int? ConnectedStateCode { get; set; }
        public int? OperationalStateCode { get; set; }
        public int? SynchronizationHealthCode { get; set; }
        public DistributedAgHealthLevel StatusLevel { get; set; } = DistributedAgHealthLevel.Unknown;
    }

    public sealed class DistributedAgDatabaseHealth
    {
        public string DatabaseName { get; set; }
        public string MemberName { get; set; }
        public string SynchronizationState { get; set; }
        public string SynchronizationHealth { get; set; }
        public string DatabaseState { get; set; }
        public bool? IsSuspended { get; set; }
        public string SuspendReason { get; set; }
        public string DisplayHealth => StatusLevel == DistributedAgHealthLevel.Critical
            ? (IsSuspended == true ? "Suspended" : "Needs attention")
            : StatusLevel == DistributedAgHealthLevel.Warning ? "Partially healthy"
            : StatusLevel == DistributedAgHealthLevel.Healthy ? "Healthy" : "Not visible";
        public long? SendQueueKb { get; set; }
        public long? RedoQueueKb { get; set; }
        public long? SendRateKb { get; set; }
        public long? RedoRateKb { get; set; }
        public long? LagSeconds { get; set; }
        public DateTime? LastHardenedTime { get; set; }
        public int? SynchronizationStateCode { get; set; }
        public int? SynchronizationHealthCode { get; set; }
        public int? DatabaseStateCode { get; set; }
        public DistributedAgHealthLevel StatusLevel { get; set; } = DistributedAgHealthLevel.Unknown;
    }

    /// <summary>Evaluates reported SQL Server health without inventing queue or lag thresholds.</summary>
    public static class DistributedAgHealthEvaluator
    {
        public static void Evaluate(DistributedAgHealthSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));

            foreach (var member in snapshot.Members)
            {
                if (member.SynchronizationHealthCode == 0 ||
                    (member.ConnectedStateCode == 0 && member.RoleCode != 1) ||
                    member.OperationalStateCode >= 3)
                    member.StatusLevel = DistributedAgHealthLevel.Critical;
                else if (member.SynchronizationHealthCode == 1 || member.RoleCode == 0 ||
                         member.OperationalStateCode == 0 || member.OperationalStateCode == 1)
                    member.StatusLevel = DistributedAgHealthLevel.Warning;
                else if (!member.HasState || !member.RoleCode.HasValue || !member.SynchronizationHealthCode.HasValue ||
                         (member.RoleCode == 2 && !member.ConnectedStateCode.HasValue))
                    member.StatusLevel = DistributedAgHealthLevel.Unknown;
                else
                    member.StatusLevel = DistributedAgHealthLevel.Healthy;
            }

            foreach (var database in snapshot.Databases)
            {
                if (database.IsSuspended == true || database.SynchronizationHealthCode == 0 ||
                    database.SynchronizationStateCode == 0 || database.DatabaseStateCode >= 3)
                    database.StatusLevel = DistributedAgHealthLevel.Critical;
                else if (database.SynchronizationHealthCode == 1 || database.SynchronizationStateCode >= 3 ||
                         database.DatabaseStateCode == 1 || database.DatabaseStateCode == 2)
                    database.StatusLevel = DistributedAgHealthLevel.Warning;
                else if (!database.SynchronizationHealthCode.HasValue || !database.SynchronizationStateCode.HasValue)
                    database.StatusLevel = DistributedAgHealthLevel.Unknown;
                else
                    database.StatusLevel = DistributedAgHealthLevel.Healthy;
            }

            var levels = snapshot.Members.Select(member => member.StatusLevel)
                .Concat(snapshot.Databases.Select(database => database.StatusLevel)).ToList();
            bool limitedVisibility = snapshot.Members.Count < 2 || snapshot.Databases.Count == 0 ||
                levels.Contains(DistributedAgHealthLevel.Unknown);

            if (levels.Contains(DistributedAgHealthLevel.Critical))
            {
                snapshot.StatusLevel = DistributedAgHealthLevel.Critical;
                snapshot.StatusText = "Needs attention";
                snapshot.StatusDetail = "SQL Server reports an unhealthy, disconnected, suspended, or offline member/database.";
            }
            else if (levels.Contains(DistributedAgHealthLevel.Warning))
            {
                snapshot.StatusLevel = DistributedAgHealthLevel.Warning;
                snapshot.StatusText = "Partially healthy";
                snapshot.StatusDetail = "A member or database is recovering, resolving, or has not reached its target synchronization state.";
            }
            else if (limitedVisibility)
            {
                snapshot.StatusLevel = DistributedAgHealthLevel.Unknown;
                snapshot.StatusText = "Partial visibility";
                snapshot.StatusDetail = "Some distributed AG state is not visible from this instance. Open this dashboard on the global primary for the fullest view.";
            }
            else
            {
                snapshot.StatusLevel = DistributedAgHealthLevel.Healthy;
                snapshot.StatusText = "Healthy";
                snapshot.StatusDetail = "All reported distributed AG members and databases meet their configured synchronization target.";
            }

            if (limitedVisibility && snapshot.StatusLevel != DistributedAgHealthLevel.Unknown)
                snapshot.StatusDetail += " Some state is not visible from this instance.";
            snapshot.StatusDetail += " Remote values are the latest state reported to this server; underlying AG replicas are not queried.";
        }
    }
}
