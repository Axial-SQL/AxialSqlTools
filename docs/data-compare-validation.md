# Compare Table Data validation

Run the portable model, comparison engine and script checks with the .NET 8 SDK:

```console
dotnet run --project tests/data-compare-smoke/DataCompare.Smoke.csproj
```

The project links the production comparison sources and has no package dependencies. It checks mapping feedback, exact type compatibility, invalid configuration rejection, all four result categories, default selection, category selection across pages, SQL generation from selected rows, exact key matching, duplicate keys and cancellation. It does not connect to SQL Server or execute generated SQL.

The 10 regression checks passed during this change, including duplicate mapping feedback and its clearing. They were compiled with C# 7.3 and executed on .NET 8 using the same production source files linked by the project. This does not verify WPF behavior or an SSMS extension build.

Build the extension on Windows with Visual Studio and SSMS 22 installed. The extension targets .NET Framework 4.8 and uses SSMS assemblies and WPF; the portable checks do not replace that build or an SSMS UI check.

In SSMS, check the following interactions using disposable source and target tables:

1. Capture each database from Object Explorer. Confirm that Active query and New connection buttons are absent, endpoint labels identify the databases, table filtering works, and Refresh preserves a selection that still exists.
2. Pick tables with a primary key and one changed row, one source-only row, one target-only row and one identical row. Confirm that Compare opens the results and that selected operation counts distinguish inserts, updates and deletes.
3. Remove the key, map two source columns to one target, choose incompatible lengths, or enter an invalid timeout. Confirm that the setup explains the problem and Compare remains disabled until corrected.
4. Change mappings and options. Confirm that previously displayed results and scripts are invalidated, then compare again.
5. Compare more than 200 rows. Confirm that individual choices survive paging and category selection affects the whole category. Target-only rows must start unselected; identical rows must remain unavailable for synchronization.
6. Review full values and changed columns, then preview and save a script. Check that the target identity and selected operations match the results. Cancel a running comparison and confirm the window remains usable.
7. Review narrow and maximized window sizes, keyboard access, and light/dark themes. Confirm the setup and result controls remain legible and usable.

The smoke checks validate generated statements without applying them. A live database synchronization check, if needed, belongs on disposable tables with known contents.
