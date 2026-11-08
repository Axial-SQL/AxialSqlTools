# Quick Manage execution history checks

Run with the .NET 10 SDK:

```sh
dotnet run --project tests/JobQuickView/JobQuickViewExecutionHistoryTests.csproj
```

These dependency-free checks exercise the production execution-list model: live
rows, stable selection, completion transitions, history lag, failed status reads,
Agent session changes, server aliases and elapsed-time updates. The production
files are compiled using C# 7.3 to match the extension's language compatibility.

The tests do not connect to SQL Server or render WPF. Live SQL Agent behavior and
SSMS layout still require Windows integration testing.
