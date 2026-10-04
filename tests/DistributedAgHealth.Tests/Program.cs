using System;
using System.Linq;
using AxialSqlTools;

public static class HealthSemanticsTests
{
    private static int passed;

    public static void Main()
    {
        Check("Async synchronizing is healthy", snapshot =>
        {
            snapshot.Databases[0].SynchronizationStateCode = 1;
            snapshot.Databases[0].SynchronizationHealthCode = 2;
        }, DistributedAgHealthLevel.Healthy);
        Check("Remote null operational state is expected", snapshot =>
            snapshot.Members[1].OperationalStateCode = null, DistributedAgHealthLevel.Healthy);
        Check("Historical connection errors do not change current health", snapshot =>
            snapshot.Members[1].LastError = "A previous connection timed out", DistributedAgHealthLevel.Healthy);
        Check("Suspension overrides healthy and zero lag", snapshot =>
        {
            snapshot.Databases[0].IsSuspended = true;
            snapshot.Databases[0].LagSeconds = 0;
        }, DistributedAgHealthLevel.Critical);
        Check("Missing remote state is unknown", snapshot =>
        {
            snapshot.Members[1].HasState = false;
            snapshot.Members[1].RoleCode = null;
            snapshot.Members[1].SynchronizationHealthCode = null;
        }, DistributedAgHealthLevel.Unknown);
        Check("Missing database rows are unknown", snapshot =>
            snapshot.Databases.Clear(), DistributedAgHealthLevel.Unknown);
        Check("Only one configured member is partial visibility", snapshot =>
            snapshot.Members.RemoveAt(1), DistributedAgHealthLevel.Unknown);
        Check("Disconnected remote with unknown role is critical", snapshot =>
        {
            snapshot.Members[1].RoleCode = null;
            snapshot.Members[1].ConnectedStateCode = 0;
        }, DistributedAgHealthLevel.Critical);
        Check("Sync replica catching up is warning", snapshot =>
            snapshot.Databases[0].SynchronizationHealthCode = 1, DistributedAgHealthLevel.Warning);
        Check("Recovery state is warning", snapshot =>
            snapshot.Databases[0].DatabaseStateCode = 2, DistributedAgHealthLevel.Warning);
        Check("Failure remains critical with partial coverage", snapshot =>
        {
            snapshot.Members[1].HasState = false;
            snapshot.Databases[0].IsSuspended = true;
        }, DistributedAgHealthLevel.Critical);
        Check("Null metrics do not invent failure", snapshot =>
        {
            snapshot.Databases[0].LagSeconds = null;
            snapshot.Databases[0].SendQueueKb = null;
            snapshot.Databases[0].RedoQueueKb = null;
        }, DistributedAgHealthLevel.Healthy);
        Check("Unknown forwarder connection is not healthy", snapshot =>
            snapshot.Members[1].ConnectedStateCode = null, DistributedAgHealthLevel.Unknown);
        Check("Unknown database synchronization is not healthy", snapshot =>
            snapshot.Databases[0].SynchronizationStateCode = null, DistributedAgHealthLevel.Unknown);
        CheckLocal("Forwarder exposes underlying local rows without certifying DAG link", Forwarder(),
            DistributedAgHealthLevel.Unknown, DistributedAgHealthLevel.Healthy, snapshot =>
            {
                Require(snapshot.LocalDatabaseCount == 1 && snapshot.Databases.Count == 2, "Local count includes remote rows");
                Require(snapshot.LocalGroupNames.SequenceEqual(new[] { "AGEYMP0_AWS" }), "Wrong underlying AG");
                Require(snapshot.LocalRole.StartsWith("Forwarder"), "Forwarder role lost");
                Require(!snapshot.IsGlobalPrimaryObservation, "Underlying primary became global primary");
                Require(snapshot.LocalDatabases[0].Role == "Primary", "Underlying row has distributed role label");
            });
        var missingDagRole = Forwarder();
        missingDagRole.Members[1].RoleCode = null;
        missingDagRole.Members[1].HasState = false;
        CheckLocal("Underlying primary cannot fill missing distributed role", missingDagRole,
            DistributedAgHealthLevel.Unknown, DistributedAgHealthLevel.Healthy, snapshot =>
            Require(!snapshot.IsGlobalPrimaryObservation && snapshot.LocalRole.StartsWith("Distributed role not visible"), "Dual roles conflated"));
        CheckLocal("Single-replica regular AG is complete", Regular(), DistributedAgHealthLevel.Healthy,
            DistributedAgHealthLevel.Healthy, snapshot => Require(snapshot.LocalRole == "Primary", "Wrong primary label"));
        CheckLocal("Regular secondary local healthy with unavailable primary", RegularSecondary(),
            DistributedAgHealthLevel.Unknown, DistributedAgHealthLevel.Healthy,
            snapshot => Require(snapshot.LocalRole == "Secondary", "Regular secondary mislabeled forwarder"));
        var missingDatabase = Regular();
        missingDatabase.Databases.Add(new DistributedAgDatabaseHealth
            { GroupName = "Regular", DatabaseName = "NotJoined", MemberName = "Node1", IsLocal = true, RoleCode = 1 });
        CheckLocal("Configured local database missing DMV row stays unknown", missingDatabase,
            DistributedAgHealthLevel.Unknown, DistributedAgHealthLevel.Unknown, snapshot =>
            Require(snapshot.LocalDatabaseCount == 2 && snapshot.LocalUnknownCount == 1, "Missing database hidden"));
        var localWarning = Regular();
        localWarning.Databases[0].SynchronizationHealthCode = 1;
        CheckLocal("Local partial synchronization is warning", localWarning, DistributedAgHealthLevel.Warning, DistributedAgHealthLevel.Warning);
        var localSuspension = Regular();
        localSuspension.Databases[0].IsSuspended = true;
        CheckLocal("Local suspension takes precedence", localSuspension, DistributedAgHealthLevel.Critical, DistributedAgHealthLevel.Critical,
            snapshot => Require(snapshot.LocalCriticalCount == 1 && snapshot.LocalHealthyCount == 0, "Wrong suspension counts"));
        var remoteFailure = Regular();
        remoteFailure.Members[0].SynchronizationHealthCode = 0;
        var remote = Member("Regular", "Node2", false, false, 2);
        remote.ConnectedStateCode = 0;
        remoteFailure.Members.Add(remote);
        CheckLocal("Remote failure in primary rollup does not repaint local copies", remoteFailure,
            DistributedAgHealthLevel.Critical, DistributedAgHealthLevel.Healthy);
        var localDisconnect = RegularSecondary();
        localDisconnect.Members[1].ConnectedStateCode = 0;
        CheckLocal("Local disconnection overrides healthy DB", localDisconnect, DistributedAgHealthLevel.Critical, DistributedAgHealthLevel.Critical);
        var localFailed = Regular();
        localFailed.Members[0].OperationalStateCode = 4;
        CheckLocal("Local operational failure overrides healthy DB", localFailed, DistributedAgHealthLevel.Critical, DistributedAgHealthLevel.Critical);
        var duplicate = Forwarder();
        duplicate.Databases.Add(Database("DAG", "AGEYMP0_AWS", true, true, 2));
        CheckLocal("Prefer physical row over duplicate DAG local row", duplicate, DistributedAgHealthLevel.Unknown,
            DistributedAgHealthLevel.Healthy, snapshot =>
            Require(snapshot.LocalDatabaseCount == 1 && !snapshot.LocalDatabases[0].IsDistributedGroup, "Local database counted twice"));
        var missingPhysical = Forwarder();
        missingPhysical.Databases[0].HasState = false;
        missingPhysical.Databases[0].SynchronizationHealthCode = null;
        missingPhysical.Databases[0].SynchronizationStateCode = null;
        missingPhysical.Databases.Add(Database("DAG", "AGEYMP0_AWS", true, true, 2));
        CheckLocal("Missing physical row cannot be concealed by healthy DAG row", missingPhysical,
            DistributedAgHealthLevel.Unknown, DistributedAgHealthLevel.Unknown);
        var suspendedDag = Forwarder();
        var suspendedDagRow = Database("DAG", "AGEYMP0_AWS", true, true, 2);
        suspendedDagRow.IsSuspended = true;
        suspendedDag.Databases.Add(suspendedDagRow);
        CheckLocal("Explicit DAG local suspension is retained with physical metrics", suspendedDag,
            DistributedAgHealthLevel.Critical, DistributedAgHealthLevel.Critical, snapshot =>
            {
                Require(snapshot.LocalCriticalCount == 1 && !string.IsNullOrEmpty(snapshot.LocalDatabases[0].LocalIssue), "DAG suspension discarded");
                Require(snapshot.Databases[0].StatusLevel == DistributedAgHealthLevel.Healthy, "Underlying source row was overwritten");
            });
        var disconnectedDag = Forwarder();
        disconnectedDag.Members[1].ConnectedStateCode = 0;
        CheckLocal("Explicit local DAG disconnection is retained", disconnectedDag, DistributedAgHealthLevel.Critical, DistributedAgHealthLevel.Critical);
        var localFallback = new DistributedAgHealthSnapshot { GroupName = "DAG", IsDistributed = true };
        localFallback.Members.Add(Member("DAG", "Source", true, true, 1));
        localFallback.Members.Add(Member("DAG", "Target", true, false, 2));
        localFallback.Databases.Add(Database("DAG", "Source", true, true, 1));
        CheckLocal("DAG local fallback when physical rows unavailable", localFallback, DistributedAgHealthLevel.Healthy,
            DistributedAgHealthLevel.Healthy, snapshot => Require(snapshot.LocalDatabases[0].IsDistributedGroup, "DAG fallback missing"));
        var mixed = Regular();
        var catchingUp = Database("Regular", "Node1", false, true, 1);
        catchingUp.DatabaseName = "CatchingUp";
        catchingUp.SynchronizationHealthCode = 1;
        mixed.Databases.Add(catchingUp);
        mixed.Databases.Add(new DistributedAgDatabaseHealth { GroupName = "Regular", DatabaseName = "Missing", IsLocal = true });
        CheckLocal("Mixed local health retains every count", mixed, DistributedAgHealthLevel.Warning, DistributedAgHealthLevel.Warning,
            snapshot => Require(snapshot.LocalHealthyCount == 1 && snapshot.LocalWarningCount == 1 && snapshot.LocalUnknownCount == 1, "Mixed counts lost data"));
        var idle = Regular();
        idle.Databases[0].LastHardenedTime = DateTime.Now.AddDays(-10);
        idle.Databases[0].LastRedoneTime = DateTime.Now.AddDays(-10);
        idle.Databases[0].LastCommitTime = DateTime.Now.AddDays(-10);
        CheckLocal("Idle timestamps do not invent lag", idle, DistributedAgHealthLevel.Healthy, DistributedAgHealthLevel.Healthy,
            snapshot => Require(snapshot.LocalDatabases[0].LagSeconds == null, "Idle age became lag"));
        var unrelated = Regular();
        var other = Database("Unrelated", "OtherNode", false, false, 2);
        other.IsSuspended = true;
        unrelated.Databases.Add(other);
        CheckLocal("Unrelated rows cannot affect selected topology", unrelated, DistributedAgHealthLevel.Healthy, DistributedAgHealthLevel.Healthy);
        Console.WriteLine(passed + " health scenarios passed.");
    }

    private static void Check(string name, Action<DistributedAgHealthSnapshot> arrange, DistributedAgHealthLevel expected)
    {
        var snapshot = new DistributedAgHealthSnapshot { GroupName = "DAG", IsDistributed = true };
        snapshot.Members.Add(new DistributedAgMemberHealth
        {
            GroupName = "DAG", IsDistributedGroup = true, Name = "Source", HasState = true, IsLocal = true, RoleCode = 1,
            SynchronizationHealthCode = 2, OperationalStateCode = 2
        });
        snapshot.Members.Add(new DistributedAgMemberHealth
        {
            GroupName = "DAG", IsDistributedGroup = true, Name = "Target", HasState = true, IsLocal = false, RoleCode = 2,
            SynchronizationHealthCode = 2, ConnectedStateCode = 1
        });
        snapshot.Databases.Add(new DistributedAgDatabaseHealth
        {
            GroupName = "DAG", IsDistributedGroup = true, HasState = true, IsLocal = false,
            DatabaseName = "Application", MemberName = "Target", IsSuspended = false,
            SynchronizationHealthCode = 2, SynchronizationStateCode = 1, DatabaseStateCode = 0
        });
        arrange(snapshot);
        DistributedAgHealthEvaluator.Evaluate(snapshot);
        if (snapshot.StatusLevel != expected)
            throw new Exception(name + ": expected " + expected + ", actual " + snapshot.StatusLevel);
        foreach (var member in snapshot.Members)
            if (member.StatusLevel != DistributedAgHealthLevel.Healthy && member.DisplayHealth == "Healthy")
                throw new Exception(name + ": unhealthy member displayed as Healthy");
        foreach (var database in snapshot.Databases)
            if (database.StatusLevel != DistributedAgHealthLevel.Healthy && database.DisplayHealth == "Healthy")
                throw new Exception(name + ": unhealthy database displayed as Healthy");
        passed++;
    }

    private static DistributedAgHealthSnapshot Forwarder()
    {
        var snapshot = new DistributedAgHealthSnapshot { GroupName = "DAG", IsDistributed = true };
        snapshot.Members.Add(new DistributedAgMemberHealth { GroupName = "DAG", Name = "OnPrem", IsDistributedGroup = true });
        snapshot.Members.Add(Member("DAG", "AGEYMP0_AWS", true, true, 2));
        snapshot.Members.Add(Member("AGEYMP0_AWS", "Node1", false, true, 1));
        snapshot.Members.Add(Member("AGEYMP0_AWS", "Node2", false, false, 2));
        snapshot.Databases.Add(Database("AGEYMP0_AWS", "Node1", false, true, 1));
        snapshot.Databases.Add(Database("AGEYMP0_AWS", "Node2", false, false, 2));
        return snapshot;
    }

    private static DistributedAgHealthSnapshot Regular()
    {
        var snapshot = new DistributedAgHealthSnapshot { GroupName = "Regular" };
        snapshot.Members.Add(Member("Regular", "Node1", false, true, 1));
        snapshot.Databases.Add(Database("Regular", "Node1", false, true, 1));
        return snapshot;
    }

    private static DistributedAgHealthSnapshot RegularSecondary()
    {
        var snapshot = new DistributedAgHealthSnapshot { GroupName = "Regular" };
        snapshot.Members.Add(new DistributedAgMemberHealth { GroupName = "Regular", Name = "Node1" });
        snapshot.Members.Add(Member("Regular", "Node2", false, true, 2));
        snapshot.Databases.Add(Database("Regular", "Node2", false, true, 2));
        return snapshot;
    }

    private static DistributedAgMemberHealth Member(string group, string name, bool distributed, bool local, int role)
    {
        return new DistributedAgMemberHealth
        {
            GroupName = group, Name = name, IsDistributedGroup = distributed, HasState = true, IsLocal = local,
            RoleCode = role, SynchronizationHealthCode = 2, OperationalStateCode = local ? (int?)2 : null,
            ConnectedStateCode = role == 2 ? (int?)1 : null
        };
    }

    private static DistributedAgDatabaseHealth Database(string group, string member, bool distributed, bool local, int role)
    {
        return new DistributedAgDatabaseHealth
        {
            GroupName = group, DatabaseName = "Application", MemberName = member, IsDistributedGroup = distributed,
            IsLocal = local, RoleCode = role, HasState = true, DatabaseId = local ? (int?)5 : null,
            IsSuspended = false, SynchronizationHealthCode = 2, SynchronizationStateCode = 1, DatabaseStateCode = 0
        };
    }

    private static void CheckLocal(string name, DistributedAgHealthSnapshot snapshot, DistributedAgHealthLevel expected,
        DistributedAgHealthLevel expectedLocal, Action<DistributedAgHealthSnapshot> verify = null)
    {
        DistributedAgHealthEvaluator.Evaluate(snapshot);
        Require(snapshot.StatusLevel == expected, name + ": topology expected " + expected + ", actual " + snapshot.StatusLevel);
        Require(snapshot.LocalStatusLevel == expectedLocal, name + ": local expected " + expectedLocal + ", actual " + snapshot.LocalStatusLevel);
        verify?.Invoke(snapshot);
        passed++;
    }

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }
}
