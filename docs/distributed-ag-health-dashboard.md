# Distributed AG Health Dashboard

View distributed availability group health inside SSMS 22 without opening queries or leaving Object Explorer.

## Open the dashboard

Right-click an availability group under **Always On High Availability > Availability Groups**, then choose **Distributed AG Health Dashboard**. The dashboard uses that Object Explorer connection and the selected AG name. It verifies that the selected group is a distributed AG.

You can also select a connected instance in Object Explorer and use **Axial SQL Tools toolbar > Tools > Distributed AG Health Dashboard**. Choose a distributed AG from the dropdown.

The dashboard contains:

- An observed-health summary and the connected SQL Server instance.
- A card for each participating AG with its role, connection state, synchronization health, commit mode, and endpoint. Hover over the connection state for the last reported connection error, if any.
- Database replication rows showing health, synchronization/database state, suspension, log send and redo queues in KB, rates in KB/s, secondary lag in seconds, and last hardened time in server time.
- Manual refresh and optional automatic refresh every 30 seconds. Hidden windows do not start new timed refreshes; closing a window cancels its query.

## Reading the results

This is a read-only view of the connected instance's distributed AG metadata and DMVs. It does not connect to endpoint URLs, change replication settings, initiate failover, or query every underlying AG replica.

Connect to the global primary for the fullest view. A forwarder or other replica can return only its local state, and remote state can be unavailable or last reported. Missing information is displayed as **Not visible**, **N/A**, or **Partial visibility**, never as a zero or a healthy result.

Asynchronous replicas can be healthy while synchronizing. Queue sizes are displayed without arbitrary failure thresholds. Lag is unavailable while data movement is suspended because SQL Server can otherwise report a misleading zero. Historical connection errors do not override current healthy connectivity.

If refresh fails, the dashboard labels any previous snapshot as stale. A healthy result reflects the rows SQL Server returned; it is not proof of failover readiness or zero data loss.

## Requirements

- SSMS 22 with Axial SQL Tools.
- SQL Server 2016 or later with Always On enabled.
- `VIEW ANY DEFINITION`, plus `VIEW SERVER STATE` on SQL Server 2016-2019 or `VIEW SERVER PERFORMANCE STATE` on SQL Server 2022 and later.

The command is available on availability group nodes without querying SQL Server while the context menu opens. Selecting a regular AG produces an explanatory message instead of displaying its state as a distributed AG.

## Validation

Run the health classification regression scenarios with .NET 8:

```shell
dotnet run --project tests/DistributedAgHealth.Tests
```

For integration validation, build the VSIX on Windows with SSMS 22 dependencies. Test on both the global primary and forwarder, and check healthy asynchronous replication, disconnection, suspension, missing permissions, refresh cancellation, and light/dark themes. Menu integration uses SSMS's Object Explorer tree and should be checked after SSMS updates.

Microsoft references: [Distributed availability groups](https://learn.microsoft.com/en-us/sql/database-engine/availability-groups/windows/distributed-availability-groups), [replica state DMV](https://learn.microsoft.com/en-us/sql/relational-databases/system-dynamic-management-objects/sys-dm-hadr-availability-replica-states-transact-sql), and [database replica state DMV](https://learn.microsoft.com/en-us/sql/relational-databases/system-dynamic-management-objects/sys-dm-hadr-database-replica-states-transact-sql).
