# Fatal action warning

In **Axial SQL Tools > Settings > Query execution**, the single **Warn when running fatal actions** setting enables confirmation before SQL execution. It is enabled by default and saved per Windows user.

The warning covers `UPDATE` and `DELETE` without a statement-level `WHERE` clause, and `TRUNCATE TABLE`, including partition truncation. It checks the selected text when there is a selection, otherwise the complete query buffer. The dialog lists the actions, target names, and line numbers relative to that execution text, together with the starting server/database.

- **Cancel**, Escape, or closing the dialog cancels the SSMS execution command. Cancel has initial focus.
- **Run once** permits only the current execution.
- **Allow repeats this session** permits the exact same execution text in this query window and live connection. It does not suppress warnings in other windows or for another script.

An execution with different text, server, database or connection identity clears that window's approval. Closing the window or saving the setting also clears approvals. No approval is persisted across SSMS restarts. Query text is fingerprinted in memory for matching; the feature neither logs nor stores the SQL text. It does not open database connections or issue SQL probes. If the live session cannot be identified, the dialog allows Run once and explains why repeat approval is unavailable.

## SQL handling

The feature uses the existing SQL Server ScriptDom parser rather than searching for keywords. Comments, quoted identifiers, string literals and a `WHERE` inside another query do not hide a missing `WHERE` on the modifying statement. Joined updates/deletes, multiple statements, control flow and `GO` repeat counts are supported. `TOP` is not a substitute for `WHERE`.

Local/global temp tables and table variables are excluded, including qualified/bracketed temp names, ordinary target aliases and CTE chains whose sources are entirely temporary. A quoted permanent table named `[@something]` is not a table variable. Ambiguous aliases, case-only alias/CTE matches (the database collation is unknown), or CTEs mixing temporary and permanent sources are treated conservatively and can still prompt.

This is a check of the SQL submitted through the SSMS **Query.Execute** command. It does not expand stored procedure calls, dynamic SQL, SQLCMD includes/variables or trigger side effects. Defining/altering a procedure, function or trigger does not execute its body and does not trigger a body warning. `DROP`, `MERGE` and broad but explicit filters such as `WHERE 1 = 1` are outside the requested warning rules.

If ScriptDom cannot fully parse the selected script, the dialog explicitly reports an incomplete check and permits Cancel or Run once. Incomplete checks cannot receive repeat approval. If the confirmation UI fails, execution is cancelled.

## Validation

With .NET 8 installed:

```sh
dotnet run --project tests/QuerySafety.Tests
```

The tests compile the production analyzer, approval cache and guard. Real ScriptDom parses the SQL fixtures; lightweight host/dialog stand-ins exercise cancellation, selection handling, run-once and repeat approvals, database/connection changes, and error paths. The suite contains 76 checks and does not connect to SQL Server.

The changed components, execution event signature, settings methods and dialog code were also compiled against .NET Framework 4.7.2, WPF and DTE references using a host harness. A full VSIX build and live SSMS verification require Windows and have not been run in the Linux development environment.

Before merging, build/install the VSIX in SSMS 22 and verify on a disposable test database:

1. F5, Ctrl+E and the Execute toolbar button prompt on a permanent-table UPDATE/DELETE without WHERE and on TRUNCATE. Cancel, Escape and the close button leave the data unchanged.
2. Run once executes and prompts again next time. Allow repeats suppresses subsequent prompts for the same execution in the same tab/session.
3. Edited SQL, another tab, a different database and a disconnect/reconnect require new approval. Closing/reopening the tab and saving settings clear approval.
4. Selected SQL is checked independently of the rest of the buffer. Temp-table and table-variable operations, including aliases, run without prompting.
5. Turning the setting off and saving disables the warning; its value survives restart. Check light/dark themes and keyboard focus.
6. Statistics capture and query history continue to work for allowed queries and do not start a new capture for cancelled execution.
