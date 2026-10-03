using System;
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
        Console.WriteLine(passed + " health scenarios passed.");
    }

    private static void Check(string name, Action<DistributedAgHealthSnapshot> arrange, DistributedAgHealthLevel expected)
    {
        var snapshot = new DistributedAgHealthSnapshot();
        snapshot.Members.Add(new DistributedAgMemberHealth
        {
            Name = "Source", HasState = true, IsLocal = true, RoleCode = 1,
            SynchronizationHealthCode = 2, OperationalStateCode = 2
        });
        snapshot.Members.Add(new DistributedAgMemberHealth
        {
            Name = "Target", HasState = true, IsLocal = false, RoleCode = 2,
            SynchronizationHealthCode = 2, ConnectedStateCode = 1
        });
        snapshot.Databases.Add(new DistributedAgDatabaseHealth
        {
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
}
