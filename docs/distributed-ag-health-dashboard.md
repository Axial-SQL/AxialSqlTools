# AG Health Dashboard

View regular and distributed availability group health inside SSMS 22. The dashboard starts with databases on the connected SQL Server instance and separates their local health from the wider availability group's visible state.

## Open the dashboard

Right-click an availability group under **Always On High Availability > Availability Groups**, then choose **AG Health Dashboard**. Regular AGs and distributed AGs use the same menu item. The dashboard captures the selected group's catalog name and Object Explorer connection before opening, independently of the active query window.

You can also select a connected instance or one of its nodes in Object Explorer and use **Axial SQL Tools toolbar > Tools > AG Health Dashboard**. Choose a regular or distributed AG from the dropdown. Selecting an AG node opens that group directly.

## Dashboard layout

Two summary cards separate **local database health** from **selected AG visibility**. A healthy local database result does not establish that every replica or distributed AG link is healthy.

The detail tabs are:

| Tab | Scope |
| --- | --- |
| **Databases on this node** | The default view. Database replication state on the connected SQL Server instance for the selected group, including its local underlying AG when a distributed AG is selected. |
| **Replica details** | The broader database replica rows returned by the connected instance, including remote rows when SQL Server makes them visible. |
| **Availability replicas** | Replica states for the selected regular AG, or the selected distributed AG's participating AGs plus its visible underlying AG replicas. Each card identifies its group and scope. |

Database rows include synchronization and database state, suspension, send/redo queues in KB, rates in KB/s, and timestamps where the DMV supplies them. The dashboard shows the connected instance and the scope of each result.

Use manual refresh or optional automatic refresh every 30 seconds. Hidden windows do not start new timed refreshes; closing a window cancels its query. If refresh fails, any previous snapshot is labeled stale.

## Regular AGs and distributed AG forwarders

For a **regular AG**, the local database view includes that group's databases on the connected instance. Local rows are available on both primary and secondary replicas. A secondary's limited view of other replicas does not make its own local database rows unavailable.

For a **distributed AG**, the dashboard queries the selected distributed group and identifies its underlying local regular AG by matching the distributed AG's participating AG names. This matters on a forwarder: its local database rows can belong to an underlying AG with `is_distributed = 0`. Filtering database rows only to the distributed AG's group ID would omit these local results.

Distributed link health remains separate from the underlying local AG's database health. An explicitly reported local distributed AG suspension or failure still needs attention, even when the underlying AG's physical database is healthy. Rows from an unrelated local AG are not included merely because it is on the same instance.

## Reading health and lag

This is a read-only snapshot of metadata and DMVs on the connected instance. It does not connect to endpoint URLs, query every underlying replica, change replication settings, or initiate failover.

A global primary usually exposes more remote state than a forwarder or secondary. Remote rows may be missing or represent the last state reported to this instance. Scope and visibility labels describe what was observed; they do not establish health across the whole distributed AG. Missing values are shown as **Not visible**, **N/A**, or **Partial visibility**, rather than being treated as healthy or zero.

The lag column uses SQL Server's `secondary_lag_seconds` DMV value when it is available and applicable. An unavailable value is **N/A**, not zero. Local primary rows show **N/A**: a forwarder's underlying AG is primary locally, which does not measure its inbound distributed AG lag. Lag is also unavailable while data movement is suspended, because SQL Server can otherwise report a misleading zero. A last-redo timestamp is an activity timestamp, not a lag measurement; time since last redo is not presented as replication lag.

Asynchronous replicas can be healthy while synchronizing. Queue sizes are displayed without arbitrary failure thresholds. Historical connection errors do not override current healthy connectivity. A healthy observed result is not proof of failover readiness or zero data loss.

## Requirements

- SSMS 22 with Axial SQL Tools.
- SQL Server 2016 or later with Always On enabled.
- `VIEW ANY DEFINITION`, plus `VIEW SERVER STATE` on SQL Server 2016-2019 or `VIEW SERVER PERFORMANCE STATE` on SQL Server 2022 and later.

The context menu is available on AG nodes without querying SQL Server while the menu opens. Metadata and health queries run after the dashboard opens.

## Validation

Run the health classification regression scenarios with .NET 8:

```shell
dotnet run --project tests/DistributedAgHealth.Tests
```

For integration validation, build the VSIX on Windows with SSMS 22 dependencies. Check:

- Regular AG primary and secondary: local database rows and selected group identity.
- Distributed AG global primary and forwarder: local underlying AG database rows and separate distributed link state.
- Multiple local AGs: no unrelated AG databases in the selected group's local view.
- Missing remote rows, unavailable DMV values, healthy asynchronous replication, disconnection, and suspension.
- Missing permissions, stale snapshots after failed refresh, refresh cancellation, and light/dark themes.
- Object Explorer targeting when the active query is connected to another server, repeated menu opening, and opening from the toolbar.

Menu integration uses SSMS's Object Explorer tree and should be checked after SSMS updates.

Microsoft references: [Distributed availability groups](https://learn.microsoft.com/en-us/sql/database-engine/availability-groups/windows/distributed-availability-groups), [replica state DMV](https://learn.microsoft.com/en-us/sql/relational-databases/system-dynamic-management-objects/sys-dm-hadr-availability-replica-states-transact-sql), and [database replica state DMV](https://learn.microsoft.com/en-us/sql/relational-databases/system-dynamic-management-objects/sys-dm-hadr-database-replica-states-transact-sql).
