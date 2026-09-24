# Quick Search navigation checks

Run with the .NET 8 SDK:

```sh
dotnet run --project tests/QuickSearchNavigation/QuickSearchNavigation.Tests.csproj -c Release
```

These checks link the production result mapping and tree traversal. They cover tables, views, procedures, scalar/inline/table-valued functions, owning objects for column/parameter matches, Agent jobs, localized folder captions, punctuation in identifiers, ambiguous captions, missing metadata, and cancellation. The tree is simulated; these checks do not validate SSMS hosting.

## SSMS 22 smoke test

Build/install the VSIX on Windows and verify:

1. Search across two databases containing identically named objects. Click **Open in Object Explorer** on table, view, procedure, and each function type. Confirm the database, schema, and object. Repeat with collapsed Object Explorer folders and names containing spaces, dots, apostrophes, and `]`.
2. Switch to a query on another server before navigating. The result must still select its original search server/login. Try two logins on the same server. A disconnected search server should show a reconnect message.
3. Check column and parameter matches select their owning table/procedure/function. A job step result should select its job under **SQL Server Agent > Jobs**. Missing or filtered objects should produce an explanation. Close Quick Search during navigation and confirm it stops without selecting another object.
4. In Query History, select different records and clear selection. Confirm the full SQL preview updates or clears, stays read-only, and supports selection and Copy. In Settings, check both Code Format examples and the Query History creation script. The source example should remain editable with Undo/Redo, Cut, Copy, and Paste.
5. Switch between light, dark, and high contrast themes with these windows open. Verify readable SQL, line numbers, selection, menus, and existing Quick Search match highlights. Theme changes must preserve SQL, selection, and the formatter source editor's undo history.
