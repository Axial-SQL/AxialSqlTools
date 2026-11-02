# Object resolver regression tests

Run from the repository root with the .NET 10 SDK:

```sh
dotnet run --project tests/SqlObjectResolver/SqlObjectResolverTests.csproj
```

This dependency-free executable links the production resolver and selection model.
It substitutes a recording SQL client and minimal SSMS UI/name-parser stubs. The SQL
client completes operations asynchronously, honors cancellation, rejects `USE`, and
records connection catalogs, parameters, queries and disposal. Tests exercise the
resolver's connection/fallback behavior, not the SQL parser, real SQL metadata or
SSMS tree navigation.

Coverage includes database-scoped Azure lookups without master access, missing
objects, explicit database catalogs with punctuation, SQL Server/Managed Instance
master results and ambiguities, current-master lookups, parameter values, original
credentials, cancellation, error propagation and disposal.

## Manual SSMS validation

1. Connect a query window to an Azure SQL Database using a contained user with no
   access to virtual master. Select an existing `dbo.TableName` and run 'Open in
   Object Explorer'. Confirm the table is selected without a database-switch error.
2. Repeat with a nonexistent object. Confirm the object-not-found message appears,
   with no master-access or `USE` error.
3. Repeat with `[CurrentDatabase].[dbo].[TableName]`, and with a database name
   containing spaces or a closing bracket where available. Confirm the intended
   database is used. Test a permitted explicit other-database connection separately.
4. On SQL Server and Azure SQL Managed Instance, test a user object and an object
   present only in master. Confirm both still resolve. If the same name exists in
   both databases, confirm both candidates remain available in the picker.
5. Repeat a lookup with the query window already connected to master.
6. Verify the shared lookup in 'Script Selected Object' still locates the intended
   user-database and master objects. Its SMO scripting behavior needs a live server.
