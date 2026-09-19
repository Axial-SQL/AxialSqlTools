# Table Data Compare

Open **Axial SQL Tools > Tools > Compare Table Data**. The command is in the same tools group as Data Transfer. Each invocation opens an independent comparison window.

This feature compares one pair of SQL Server user tables at a time. They can have different names and live in different databases or on different servers. It implements the table-level comparison workflow familiar from [Visual Studio Data Compare](https://learn.microsoft.com/en-us/sql/ssdt/how-to-compare-and-synchronize-the-data-of-two-databases): choose source and target, map columns, choose a key, inspect differences, and synchronize selected rows.

## Compare tables

1. For each side, choose **Object Explorer**, **Active query**, or **New connection**. When using Object Explorer, first select the desired database there. New connections support Windows, SQL Server, Microsoft Entra interactive, and Microsoft Entra default authentication. Credentials remain in memory for this window and are not saved.
2. Select a table on each side. Use **Refresh** to reload the table list after schema changes.
3. Review column mappings. Names are matched exactly first, with an unambiguous case-insensitive fallback. Different column names can be mapped manually. Unmapped, unsupported, computed, generated, and rowversion columns are excluded by default.
4. Choose a source primary/unique key or select a custom combination using the **Key** checkboxes. Every key column is automatically included. The complete combination must be unique on both sides. A duplicate, including a duplicate NULL combination, stops the comparison and discards incomplete results. A unique constraint on the target is recommended, but not required.
5. Configure optional non-key text comparisons, query timeout, and temporary disk limit. Select **Compare**. Database reads and comparison run outside the UI thread; **Cancel** requests SQL command cancellation and stops local sorting/merging.

Mapped base SQL types, lengths, precision, and scale must match. This intentionally avoids implicit comparison conversions. Keys always use exact, ordinal string matching, including trailing spaces, independent of either database's collation. The optional case and trailing-space settings apply only to non-key text. Trailing spaces means U+0020, not tabs or arbitrary whitespace. Binary values are compared byte for byte; NULL is distinct from empty text and empty binary. SQL decimals retain all 38 digits. datetimeoffset values are compared by their UTC instant. XML is compared as the SQL Server serialized text, without XML canonicalization.

## Inspect results

The result categories are **Different**, **Only in source**, **Only in target**, and **Identical**. The grid displays 200 rows per page, ordered by the selected keys. It reads the complete comparison; paging is not a row limit on the comparison itself. Select a row to see source and target values for every mapped column. Changed columns are highlighted. Double-click a column to inspect full values.

Different and source-only rows are selected for synchronization initially. Target-only rows are excluded initially. Selecting a target-only row means **delete it from the target**. Category selection applies to the entire category across all pages. Individual selections survive paging. Identical rows cannot be selected.

Changing connections, tables, mappings, keys, or options invalidates the previous result. Changing row selections invalidates the script preview.

## Synchronize

- **Preview script** validates selected actions and displays SQL. The preview is limited to 500,000 characters; a truncated preview is clearly marked and begins with a stopping error. **Save SQL** writes the complete script without the preview limit, replacing an existing file only after successful generation.
- **Apply selected to target** confirms the exact target and insert/update/delete counts, then executes the same guarded statements on a new target connection in one serializable transaction. Old results are discarded after an apply attempt. Run Compare again to verify the current state.
- Inserts copy mapped writable columns, including selected identity values through `IDENTITY_INSERT`. Excluded target columns use their defaults or NULL where allowed. Inserts are rejected when an unmapped required target column has neither a default nor an automatically generated value.
- Updates modify only changed, mapped writable non-key columns. Differences in identity, computed, generated, or rowversion columns cannot be applied. Exclude such columns from comparison or deselect those rows.
- Deletes remove selected target-only rows. Changes are ordered as deletes, updates, then inserts.

Scripts check the canonical target server and database, mapped target column metadata, original compared target values, and row counts. Target key checks use update/range locks within the transaction. Text concurrency checks include byte content and length, so case-insensitive collations or trailing-space padding do not hide changes. Post-write checks detect lossy conversions and triggers that suppress or alter the intended row. A detected conflict or SQL error rolls back the transaction. An exported script requires a fresh session without an existing transaction.

Constraints and triggers remain enabled. Foreign keys, unique-value swaps, self-referencing tables, typed XML constraints, or trigger behavior can prevent synchronization. Resolve those dependencies before retrying; this feature does not temporarily disable constraints or reorder multiple dependent tables. Changes to excluded columns are not part of the original-value checks. Triggers can affect other rows as part of their normal behavior. After a connection failure near COMMIT, compare again to establish the actual outcome before retrying.

Comparison reads never modify either table. Apply needs the corresponding INSERT, UPDATE, and DELETE permissions, and identity insertion can require additional ownership/ALTER permissions. Table listing and metadata require visibility of the selected objects; reading requires SELECT. Temporal history tables cannot be synchronization targets. The metadata queries target SQL Server 2016 or later and Azure SQL Database/Managed Instance equivalents.

## Export and storage

**Export comparison CSV** exports every category, with one record per compared row/column, full values, selection flags, and explicit row-existence and NULL flags. Cells beginning with spreadsheet formula markers receive a leading apostrophe. The CSV is a human-readable report, not a lossless database import format; the SQL export retains values for synchronization.

The engine performs an external merge sort in approximately 32 MiB chunks, merges at most 16 runs at a time, and stores the results and category offset indexes on disk. It does not assume the two SQL collations sort identically, and it does not approximate equality with checksums. Pages retain lightweight display strings; full data is loaded for the selected row. Very large values can increase transient memory during merge, display, and script generation.

Temporary data is placed below `%TEMP%\AxialSQL\DataCompare\<unique-session-id>`. The configurable default limit is 20 GiB across intermediate and result files. Exceeding the limit stops the entire comparison. Individual values are limited to 16 MiB and complete rows to 32 MiB. Exclude large columns if those limits are reached. Temporary files contain table data and use the current user's temporary directory permissions. They are deleted when results are replaced or the comparison window is disposed, and after handled failures/cancellation. An SSMS/process crash can leave files behind; remove abandoned session folders after closing SSMS.

The default read isolation is read committed. Optional snapshot isolation requires `ALLOW_SNAPSHOT_ISOLATION ON` on both databases and fails if it is unavailable. Each side is read at its own point in time, even with snapshot isolation; this is not a distributed point-in-time snapshot. Use stable tables or an appropriate database snapshot/maintenance window when concurrent writes could affect the comparison.

Unsupported columns include Always Encrypted values, `sql_variant`, CLR/spatial types, and newer types outside the explicit supported set. They can be excluded to compare the remaining columns. Database-wide object selection, schema synchronization, row filters, persisted comparison profiles, and automatic cross-table dependency scheduling are outside this table-pair feature.

## Build and verification

The extension keeps its existing .NET Framework 4.7.2 target, SSMS 22 assembly references, and shared WPF theme. No new production NuGet dependencies are added. Build `AxialSqlTools/AxialSqlTools.sln` in the repository's supported Windows/Visual Studio environment with SSMS 22 installed.

The dependency-free core regression executable links the production comparison source files and targets .NET 8 for cross-platform execution:

```powershell
dotnet run --project tests/DataCompare.Tests/DataCompare.Tests.csproj
```

It covers category classification, composite/NULL/duplicate keys, external-sort merge passes, exact binary/string comparison, full-precision values, paging, storage cleanup/limits, cancellation, mappings, SQL guards, identity handling, read-only columns, selection state, and CSV escaping.

Validation performed during implementation:

- 47 core regression cases passed using the C# 7.3 compiler.
- Comparison, SQL reader, and synchronization sources compiled against the real .NET Framework 4.7.2 and Microsoft.Data.SqlClient 5.2.2 reference assemblies.
- WPF control code compiled against real .NET Framework WPF references using generated field declarations and stand-ins for existing SSMS connection/theme helpers. This checks C# API compatibility, not XAML runtime behavior or the complete extension build.
- New XAML, project registration, and command-table XML were checked for well-formedness and file inclusion.
- Microsoft's TSql160Parser accepted generated insert/update/delete/identity scripts, the metadata queries, and the smoke-test SQL fixture.

Full VSIX compilation, WPF rendering, SSMS command loading, authentication, live SQL execution/rollback, and snapshot-isolation behavior still require Windows/SSMS and a SQL Server connection. They were not exercised in the Linux development environment.

Use `tests/sql/DataCompareFixture.sql` in a disposable database for the integration smoke test:

1. Compare `AxialDataCompareTest.SourceRows` to `TargetRows`, using Id as the key and the default column selections. Expect one row in each category.
2. Inspect Unicode, decimal(38,10), binary, datetime, datetime2, datetimeoffset, time, and XML values. Check the full-value viewer and the dark/light themes.
3. Preview/save a script. Apply the default selection: Id=2 updates, Id=3 inserts, and Id=4 remains.
4. Recompare, select the target-only row, and apply. Recompare again: three identical rows and no differences.
5. Reset the disposable fixture, compare, change target Id=2 from another query, then apply. It must stop on the original-value check without committing the other selected actions.
6. Choose DuplicateRows as source and SourceRows as target, select only Id/Label, and select Id as a custom key. The duplicate key must prevent results/synchronization.
7. Exercise cancellation on a larger fixture, table/column changes, denied permissions, disabled snapshot isolation, and closing multiple comparison windows. Verify temporary-file cleanup on normal disposal.
