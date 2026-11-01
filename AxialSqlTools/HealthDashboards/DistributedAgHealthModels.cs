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
        public string GroupId { get; set; }
        public bool IsDistributed { get; set; }
        public DateTime CollectedAt { get; set; }
        public string ServerVersion { get; set; }
        public List<DistributedAgMemberHealth> Members { get; } = new List<DistributedAgMemberHealth>();
        public List<DistributedAgDatabaseHealth> Databases { get; } = new List<DistributedAgDatabaseHealth>();
        public string StatusText { get; set; }
        public DistributedAgHealthLevel StatusLevel { get; set; } = DistributedAgHealthLevel.Unknown;
        public string StatusDetail { get; set; }
        public List<string> LocalGroupNames { get; } = new List<string>();
        public string LocalRole { get; set; }
        public List<DistributedAgDatabaseHealth> LocalDatabases { get; } = new List<DistributedAgDatabaseHealth>();
        public DistributedAgHealthLevel LocalStatusLevel { get; set; } = DistributedAgHealthLevel.Unknown;
        public string LocalStatusText { get; set; }
        public string LocalStatusDetail { get; set; }
        public int LocalDatabaseCount => LocalDatabases.Count;
        public int LocalHealthyCount => LocalDatabases.Count(database => database.StatusLevel == DistributedAgHealthLevel.Healthy);
        public int LocalWarningCount => LocalDatabases.Count(database => database.StatusLevel == DistributedAgHealthLevel.Warning);
        public int LocalCriticalCount => LocalDatabases.Count(database => database.StatusLevel == DistributedAgHealthLevel.Critical);
        public int LocalUnknownCount => LocalDatabases.Count(database => database.StatusLevel == DistributedAgHealthLevel.Unknown);
        public bool IsGlobalPrimaryObservation => IsDistributed && Members.Any(member =>
            member.IsDistributedGroup && member.GroupName == GroupName && member.IsLocal == true && member.RoleCode == 1);
    }

    public sealed class DistributedAgMemberHealth
    {
        public string Name { get; set; }
        public string GroupName { get; set; }
        public string GroupId { get; set; }
        public bool IsDistributedGroup { get; set; }
        public string ScopeLabel => IsDistributedGroup ? "Distributed AG link" : "Availability group";
        public string LocationLabel => IsLocal == true ? "This instance" : IsLocal == false ? "Remote replica" : "Not visible";
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
        public string GroupName { get; set; }
        public string GroupId { get; set; }
        public bool IsDistributedGroup { get; set; }
        public string ScopeLabel => IsDistributedGroup ? "Distributed AG link" : "Availability group";
        public bool? IsLocal { get; set; }
        public string LocationLabel => IsLocal == true ? "This instance" : IsLocal == false ? "Remote replica" : "Not visible";
        public string Role { get; set; }
        public int? RoleCode { get; set; }
        public bool HasState { get; set; }
        public int? DatabaseId { get; set; }
        public string GroupDatabaseId { get; set; }
        public string SynchronizationState { get; set; }
        public string SynchronizationHealth { get; set; }
        public string DatabaseState { get; set; }
        public bool? IsSuspended { get; set; }
        public string SuspendReason { get; set; }
        public string LocalIssue { get; set; }
        public string DisplayHealth => StatusLevel == DistributedAgHealthLevel.Critical
            ? (IsSuspended == true ? "Suspended" : !string.IsNullOrEmpty(LocalIssue) ? "Local DAG issue" : "Needs attention")
            : StatusLevel == DistributedAgHealthLevel.Warning ? "Partially healthy"
            : StatusLevel == DistributedAgHealthLevel.Healthy ? "Healthy" : "Not visible";
        public long? SendQueueKb { get; set; }
        public long? RedoQueueKb { get; set; }
        public long? SendRateKb { get; set; }
        public long? RedoRateKb { get; set; }
        public long? LagSeconds { get; set; }
        public DateTime? LastHardenedTime { get; set; }
        public DateTime? LastRedoneTime { get; set; }
        public DateTime? LastCommitTime { get; set; }
        public int? SynchronizationStateCode { get; set; }
        public int? SynchronizationHealthCode { get; set; }
        public int? DatabaseStateCode { get; set; }
        public DistributedAgHealthLevel StatusLevel { get; set; } = DistributedAgHealthLevel.Unknown;

        internal DistributedAgDatabaseHealth CopyForLocalView()
        {
            return (DistributedAgDatabaseHealth)MemberwiseClone();
        }
    }

    /// <summary>Evaluates reported SQL Server health without inventing queue or lag thresholds.</summary>
    public static class DistributedAgHealthEvaluator
    {
        public static void Evaluate(DistributedAgHealthSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));

            foreach (var member in snapshot.Members)
            {
                member.Role = RoleText(member.IsDistributedGroup, member.RoleCode);
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
                database.Role = RoleText(database.IsDistributedGroup, database.RoleCode);
                if (database.IsSuspended == true || database.SynchronizationHealthCode == 0 ||
                    database.SynchronizationStateCode == 0 || database.DatabaseStateCode >= 3)
                    database.StatusLevel = DistributedAgHealthLevel.Critical;
                else if (database.SynchronizationHealthCode == 1 || database.SynchronizationStateCode >= 3 ||
                         database.DatabaseStateCode == 1 || database.DatabaseStateCode == 2)
                    database.StatusLevel = DistributedAgHealthLevel.Warning;
                else if (!database.HasState || !database.SynchronizationHealthCode.HasValue || !database.SynchronizationStateCode.HasValue)
                    database.StatusLevel = DistributedAgHealthLevel.Unknown;
                else
                    database.StatusLevel = DistributedAgHealthLevel.Healthy;
            }

            EvaluateTopology(snapshot);
            EvaluateLocal(snapshot);
        }

        public static string RoleText(bool isDistributedGroup, int? role)
        {
            return role == 1 ? (isDistributedGroup ? "Global primary" : "Primary")
                : role == 2 ? (isDistributedGroup ? "Forwarder" : "Secondary")
                : role == 0 ? "Resolving" : "Not visible";
        }

        private static void EvaluateTopology(DistributedAgHealthSnapshot snapshot)
        {
            // Underlying AG health cannot establish the health of the distributed link.
            var members = snapshot.Members.Where(member => IsSelectedGroup(snapshot, member.GroupId, member.GroupName)).ToList();
            var databases = snapshot.Databases.Where(database => IsSelectedGroup(snapshot, database.GroupId, database.GroupName)).ToList();
            var levels = members.Select(member => member.StatusLevel)
                .Concat(databases.Select(database => database.StatusLevel)).ToList();
            bool limitedVisibility = members.Count == 0 || (snapshot.IsDistributed && members.Count != 2) ||
                !members.Any(member => member.RoleCode == 1) || databases.Count == 0 ||
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
                snapshot.StatusDetail = "Some selected " + (snapshot.IsDistributed ? "distributed AG" : "AG") +
                    " state is not visible from this instance. Connect to the " +
                    (snapshot.IsDistributed ? "global primary" : "primary replica") + " for the fullest topology view.";
            }
            else
            {
                snapshot.StatusLevel = DistributedAgHealthLevel.Healthy;
                snapshot.StatusText = "Healthy";
                snapshot.StatusDetail = "All reported members and databases of the selected " +
                    (snapshot.IsDistributed ? "distributed AG" : "AG") + " meet their configured synchronization target.";
            }

            if (limitedVisibility && snapshot.StatusLevel != DistributedAgHealthLevel.Unknown)
                snapshot.StatusDetail += " Some state is not visible from this instance.";
            snapshot.StatusDetail += " Remote values are the latest state reported to this server; other servers are not queried.";
        }

        private static void EvaluateLocal(DistributedAgHealthSnapshot snapshot)
        {
            snapshot.LocalDatabases.Clear();
            // One physical local database can have both an underlying AG row and a DAG row.
            // Prefer the underlying row, including an unknown catalog placeholder, so it is never counted twice.
            snapshot.LocalDatabases.AddRange(snapshot.Databases.Where(database => database.IsLocal == true)
                .GroupBy(database => database.DatabaseName, StringComparer.Ordinal)
                .Select(CreateLocalDatabase)
                .OrderBy(database => database.DatabaseName, StringComparer.CurrentCulture));

            var localMembers = snapshot.Members.Where(member => member.IsLocal == true).ToList();
            var physicalMembers = localMembers.Where(member => !member.IsDistributedGroup).ToList();
            if (physicalMembers.Count == 0) physicalMembers = localMembers;
            snapshot.LocalGroupNames.Clear();
            snapshot.LocalGroupNames.AddRange(physicalMembers.Select(member => member.GroupName)
                .Concat(snapshot.LocalDatabases.Select(database => database.GroupName))
                .Where(name => !string.IsNullOrEmpty(name)).Distinct(StringComparer.Ordinal)
                .OrderBy(name => name, StringComparer.CurrentCulture));

            var selectedLocalMember = localMembers.FirstOrDefault(member => IsSelectedGroup(snapshot, member.GroupId, member.GroupName));
            string physicalRole = string.Join(" / ", physicalMembers.Select(member => member.Role).Distinct());
            snapshot.LocalRole = selectedLocalMember != null ? selectedLocalMember.Role
                : !string.IsNullOrEmpty(physicalRole) ? physicalRole : "Not visible";
            if (snapshot.IsDistributed && physicalMembers.Any(member => !member.IsDistributedGroup))
                snapshot.LocalRole = (selectedLocalMember != null && selectedLocalMember.RoleCode.HasValue
                    ? selectedLocalMember.Role : "Distributed role not visible") +
                    " (underlying AG: " + physicalRole + ")";

            var levels = snapshot.LocalDatabases.Select(database => database.StatusLevel)
                .Concat(physicalMembers.Select(LocalMemberLevel))
                .Concat(localMembers.Where(member => member.IsDistributedGroup).Select(LocalMemberLevel)
                    .Where(level => level == DistributedAgHealthLevel.Critical || level == DistributedAgHealthLevel.Warning)).ToList();
            bool limitedVisibility = snapshot.LocalDatabases.Count == 0 || physicalMembers.Count == 0 ||
                levels.Contains(DistributedAgHealthLevel.Unknown);
            if (levels.Contains(DistributedAgHealthLevel.Critical))
            {
                snapshot.LocalStatusLevel = DistributedAgHealthLevel.Critical;
                snapshot.LocalStatusText = "Needs attention";
                snapshot.LocalStatusDetail = "A local database or replica has an active health issue, including any reported local distributed AG data movement issue.";
            }
            else if (levels.Contains(DistributedAgHealthLevel.Warning))
            {
                snapshot.LocalStatusLevel = DistributedAgHealthLevel.Warning;
                snapshot.LocalStatusText = "Partially healthy";
                snapshot.LocalStatusDetail = "A local database or replica is recovering, resolving, or catching up to its synchronization target.";
            }
            else if (limitedVisibility)
            {
                snapshot.LocalStatusLevel = DistributedAgHealthLevel.Unknown;
                snapshot.LocalStatusText = "Local state incomplete";
                snapshot.LocalStatusDetail = "Some configured local database or replica state is not available on this instance.";
            }
            else
            {
                snapshot.LocalStatusLevel = DistributedAgHealthLevel.Healthy;
                snapshot.LocalStatusText = "Healthy";
                snapshot.LocalStatusDetail = "The reported local databases and this instance's local replica meet their health targets.";
            }
            if (limitedVisibility && snapshot.LocalStatusLevel != DistributedAgHealthLevel.Unknown)
                snapshot.LocalStatusDetail += " Some local state is also unavailable.";
            snapshot.LocalStatusDetail += snapshot.IsDistributed
                ? " This describes the local copies, not the health or lag of the distributed AG link."
                : " Remote replica visibility is evaluated separately.";
        }

        private static DistributedAgHealthLevel LocalMemberLevel(DistributedAgMemberHealth member)
        {
            // A primary's synchronization-health rollup includes its remote replicas.
            // It must not make healthy local database copies look unhealthy.
            if ((member.ConnectedStateCode == 0 && member.RoleCode != 1) || member.OperationalStateCode >= 3)
                return DistributedAgHealthLevel.Critical;
            if (member.RoleCode == 0 || member.OperationalStateCode == 0 || member.OperationalStateCode == 1)
                return DistributedAgHealthLevel.Warning;
            if (!member.HasState || !member.RoleCode.HasValue ||
                (member.RoleCode == 2 && !member.ConnectedStateCode.HasValue))
                return DistributedAgHealthLevel.Unknown;
            return DistributedAgHealthLevel.Healthy;
        }

        private static DistributedAgDatabaseHealth CreateLocalDatabase(IGrouping<string, DistributedAgDatabaseHealth> rows)
        {
            var result = rows.OrderBy(database => database.IsDistributedGroup).First().CopyForLocalView();
            var adverseDagRow = rows.Where(database => database.IsDistributedGroup &&
                    (database.StatusLevel == DistributedAgHealthLevel.Critical || database.StatusLevel == DistributedAgHealthLevel.Warning))
                .OrderByDescending(database => database.StatusLevel == DistributedAgHealthLevel.Critical).FirstOrDefault();
            if (!result.IsDistributedGroup && adverseDagRow != null)
            {
                result.LocalIssue = adverseDagRow.IsSuspended == true
                    ? "Local distributed AG data movement is suspended. " + adverseDagRow.SuspendReason
                    : "A local distributed AG database row reports " + adverseDagRow.DisplayHealth.ToLowerInvariant() + ".";
                if (adverseDagRow.StatusLevel == DistributedAgHealthLevel.Critical || result.StatusLevel == DistributedAgHealthLevel.Healthy)
                    result.StatusLevel = adverseDagRow.StatusLevel;
            }
            return result;
        }

        private static bool IsSelectedGroup(DistributedAgHealthSnapshot snapshot, string groupId, string groupName)
        {
            return !string.IsNullOrEmpty(snapshot.GroupId) && !string.IsNullOrEmpty(groupId)
                ? string.Equals(snapshot.GroupId, groupId, StringComparison.OrdinalIgnoreCase)
                : string.Equals(snapshot.GroupName, groupName, StringComparison.Ordinal);
        }
    }
}
