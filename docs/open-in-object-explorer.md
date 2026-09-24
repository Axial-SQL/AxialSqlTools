**Open in Object Explorer**

Select an object name in a connected SQL query window, or put the cursor inside the name. Choose **Open in Object Explorer** from the SQL editor's right-click menu or from **Axial SQL Tools > Tools**, beside **Script Definition to New Window**. You can assign a shortcut to `AxialSqlTools.OpenInObjectExplorer` in SSMS keyboard settings.

The command expands Object Explorer, selects the object, and scrolls it into view. If the query's connection is missing in Object Explorer, it connects with a snapshot of that query window's SSMS connection settings. Changing tabs during lookup does not change the captured target.

Both navigation and scripting use the same object resolver and multiple-match picker. Names can be unqualified, schema-qualified, or database-qualified, including bracketed/double-quoted names and `database..object`. The lookup searches the named or active database, plus `master`, matching the existing scripting behavior. Selecting an object from the picker navigates to its actual database and schema.

Supported branches include tables, views, stored procedures, scalar/table functions, synonyms, table types, sequences, aggregates, indexes, keys, checks, and table/view triggers. A default constraint selects its owning column because Object Explorer has no separate node for that constraint; the status bar explains the selection. Synonyms select the synonym itself.

Caret lookup skips comments and single-quoted string literals. Aliases and CTE names are treated as object names, not resolved to their SQL definitions. Temp tables, table variables, linked-server names, invalid selections, unsupported object types, and unavailable objects produce an explanatory message. A lookup also requires metadata visibility in the connected database. Missing or filtered nodes produce an error after their branch has been enumerated, with suggestions to check filters, permissions, and refresh state. A 30-second cancellation budget bounds connection polling and traversal between native calls. SSMS's synchronous child enumeration cannot be interrupted while it is running and may exceed that budget on a slow connection. Existing filters are not cleared automatically.

**Automated checks**

Run the portable checks with a .NET 8 SDK:

```sh
dotnet run --project tests/ObjectExplorer.Tests/ObjectExplorer.Tests.csproj
```

The checks exercise production parsing, hierarchy traversal, reflection, and the host adapter. Cases include cold branch loading, localized folder captions, schema grouping above and below the object-type folder, system databases, functions, table/view children, defaults, exact-case and dotted identifiers, filtered objects, cancellation, inherited private members, and duplicate server roots with different logins/authentication modes.

`HostContracts.cs` provides small SSMS/WinForms substitutes for the portable adapter tests. In particular, the fake explorer exposes no `FindNode` or `SynchronizeTree`: navigation must use its attached tree. These contracts are not a replacement for compilation against real SSMS assemblies or the Windows smoke test below.

**SSMS 22 validation**

Build the extension on Windows with the repository's SSMS 22 and Visual Studio prerequisites. Load it in SSMS 22 and use an expendable test database to check:

1. Both command placements exist and invoke the same handler. Open the Tools menu and the SQL editor context menu with a selected name, then with only the caret inside a name; the command must remain enabled in both cases. Also test an assigned keyboard shortcut. With no query window or connection, invoking the command shows an explanatory message. During navigation the command is disabled to prevent duplicate runs, and becomes available again after completion, cancellation, or failure.
2. Select a table name and invoke the command with all Object Explorer folders collapsed. Repeat with only the caret inside the name. The correct node is selected and visible. Repeat with schema grouping enabled and disabled, and with a localized Object Explorer UI.
3. Repeat for a view, stored procedure, each function kind, synonym, table type, index, PK/FK/check, table trigger, view trigger, and default constraint. For a default, confirm the owning column and status message.
4. Test schema-qualified names, three-part names targeting a second database, spaces, escaped closing brackets, apostrophes, and non-ASCII names. Test two schemas with identically named objects and choose each picker result.
5. Test a matching object in `master`. Test a case-sensitive database with names differing only by case; the picker must not collapse distinct objects.
6. Disconnect the query's server from Object Explorer, leaving the query connected. Invoke the command and confirm automatic connection with the existing query's authentication and encryption options.
7. Use two servers, named instances, ports, and SQL logins. Keep the wrong connection selected in Object Explorer and confirm navigation never silently selects that connection's object. Keep both connections open: navigation must select the attached node under the matching server/login/authentication root, even when both roots have identical server URNs.
8. Start a lookup and switch query tabs. Confirm the original query's target remains authoritative.
9. Test no selection on whitespace, comments, strings, malformed identifiers, four-part names, temp objects, no query connection, insufficient metadata permissions, and a missing object. Confirm clear messages with no empty script windows.
10. Apply a filter excluding the target or use a slow/unavailable target. Confirm bounded navigation and an actionable message. Cancel the ambiguity picker with Escape and confirm that neither scripting nor navigation continues.

**Navigation implementation**

The implementation was compared with `SsmsObjectExplorerSelector` in the user-supplied SQL Search add-in and its bundled Shell assembly. It follows the same host interaction sequence:

1. Obtain the actual connected root from the Object Explorer `Tree` control.
2. Call the node's `EnumerateChildren(false)` before `Expand()` and reading its children.
3. Identify folders through `containedItem.context["UniqueName"]`, including `UserTables`, `UserProgrammability`, `StoredProcedures`, and `UsrDbFunctions`. Caption fallbacks cover older/synthetic folders. Walk optional schema and system folders as needed.
4. Set the real tree's `SelectedNode`, scroll it into view, and invoke the explorer's `Show()` method or focus the tree.

The old implementation used generated collection URNs with `FindNode`, then assumed its result had a usable `TreeNode`. A detached `NodeContext` does not satisfy that assumption. Navigation now stays on the matched attached tree and does not call `FindNode` or `SynchronizeTree`.

Axial retains exact connection matching, cancellation checks between native calls, exact catalog identity matching, and default-constraint-to-column navigation. It also permits system databases. No Redgate libraries are required or distributed.

Native tree loading and selection must run on the SSMS UI thread. The traversal yields between branches, but native enumeration itself may block until SSMS completes it. Real SSMS behavior, authentication prompts, themes, and VSIX packaging still require the Windows checks above.
