**Open in Object Explorer**

Select an object name in a connected SQL query window, or put the cursor inside the name. Choose **Open in Object Explorer** from the SQL editor's right-click menu or from **Axial SQL Tools > Tools**, beside **Script Definition to New Window**. You can assign a shortcut to `AxialSqlTools.OpenInObjectExplorer` in SSMS keyboard settings.

The command expands Object Explorer, selects the object, and scrolls it into view. If the query's connection is missing in Object Explorer, it connects with a snapshot of that query window's SSMS connection settings. Changing tabs during lookup does not change the captured target.

Both navigation and scripting use the same object resolver and multiple-match picker. Names can be unqualified, schema-qualified, or database-qualified, including bracketed/double-quoted names and `database..object`. The lookup searches the named or active database, plus `master`, matching the existing scripting behavior. Selecting an object from the picker navigates to its actual database and schema.

Supported branches include tables, views, stored procedures, scalar/table functions, synonyms, table types, sequences, aggregates, indexes, keys, checks, and table/view triggers. A default constraint selects its owning column because Object Explorer has no separate node for that constraint; the status bar explains the selection. Synonyms select the synonym itself.

Caret lookup skips comments and single-quoted string literals. Aliases and CTE names are treated as object names, not resolved to their SQL definitions. Temp tables, table variables, linked-server names, invalid selections, unsupported object types, and unavailable objects produce an explanatory message. A lookup also requires metadata visibility in the connected database. Navigation times out after 30 seconds if Object Explorer cannot load the branch, and suggests checking its filters, permissions, and refresh state. Existing filters are not cleared automatically.

**Automated checks**

Run the portable checks with a .NET 8 SDK:

```sh
dotnet run --project tests/ObjectExplorer.Tests/ObjectExplorer.Tests.csproj
```

The checks cover multipart-name parsing, quoted/escaped names, caret boundaries, tabs and CRLF, comments/literals, unsupported input, object paths, view-owned children, default-column navigation, connection identity, delayed connection registration, lazy-loading retries, failure propagation, and cancellation. They exercise the production parser/path/navigation code. They do not simulate SSMS's actual tree control or connect to a SQL Server.

**SSMS 22 validation**

Build the extension on Windows with the repository's SSMS 22 and Visual Studio prerequisites. Load it in SSMS 22 and use an expendable test database to check:

1. Both command placements exist and invoke the same handler. With no query window, the command is disabled.
2. Select a table name and invoke the command with all Object Explorer folders collapsed. Repeat with only the caret inside the name. The correct node is selected and visible.
3. Repeat for a view, stored procedure, each function kind, synonym, table type, index, PK/FK/check, table trigger, view trigger, and default constraint. For a default, confirm the owning column and status message.
4. Test schema-qualified names, three-part names targeting a second database, spaces, escaped closing brackets, apostrophes, and non-ASCII names. Test two schemas with identically named objects and choose each picker result.
5. Test a matching object in `master`. Test a case-sensitive database with names differing only by case; the picker must not collapse distinct objects.
6. Disconnect the query's server from Object Explorer, leaving the query connected. Invoke the command and confirm automatic connection with the existing query's authentication and encryption options.
7. Use two servers, named instances, ports, and SQL logins. Keep the wrong connection selected in Object Explorer and confirm navigation never silently selects that connection's object. Multiple OE roots with the same server URN but different credentials may require closing the extra root if SSMS's FindNode API keeps returning it; failure must be explicit.
8. Start a lookup and switch query tabs. Confirm the original query's target remains authoritative.
9. Test no selection on whitespace, comments, strings, malformed identifiers, four-part names, temp objects, no query connection, insufficient metadata permissions, and a missing object. Confirm clear messages with no empty script windows.
10. Apply a filter excluding the target or use a slow/unavailable target. Confirm bounded navigation and an actionable message. Cancel the ambiguity picker with Escape and confirm that neither scripting nor navigation continues.

The SSMS boundary uses the documented `IObjectExplorerService` methods plus isolated access to SSMS 22's `Tree`/`TreeNode` wrappers. Host UI behavior therefore needs a real SSMS smoke test even when portable checks pass. API reference: [IObjectExplorerService](https://learn.microsoft.com/en-us/dotnet/api/microsoft.sqlserver.management.ui.vsintegration.objectexplorer.iobjectexplorerservice?view=sqlserver-2016).
