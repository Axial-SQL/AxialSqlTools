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
4. In Query History, select different records and clear selection. Confirm the full SQL preview updates or clears, stays read-only, and supports selection and Copy. In Settings, check both Code Format examples and the Query History creation script. All SQL examples must remain read-only and support selection and Copy. Typing, Cut, and Paste must not change their text.
5. Switch between light, dark, and high contrast themes with these windows open. Verify readable SQL, line numbers, selection, menus, and existing Quick Search match highlights. Theme changes must preserve SQL and selection.

## SQL highlighting comparison

Compare the Settings source example with the same SQL in the SSMS editor. In the light theme, expect regular-weight blue keywords (including `GO`), red strings, green comments, magenta built-in functions such as `GETDATE()`, gray joins/operators/punctuation, and plain text for variables, numeric literals, and user object names such as `c.CustomerID` and `dbo.func(...)`. Dark themes adjust these colors for contrast; high contrast uses the system text color.

Also check escaped strings (`N'it''s text'`), bracketed names (`[name]]part]`), quoted names, nested block comments, and SQL immediately following each. These must not leak their color into the next statement. The preview uses lexical rules and common built-in names; it does not use SSMS's semantic classifier or custom Fonts and Colors settings.
