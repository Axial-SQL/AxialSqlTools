# SQL Server Builds: load errors and recovery

The window opens even when the build workbook cannot be loaded. It shows progress, the displayed data source and load time, and a readable error or warning summary. Expand **Load and parsing details** to inspect or copy the underlying error and worksheet/row diagnostics.

- **Refresh / Retry download** loads the Microsoft source again. Network requests and response reads have a 30-second timeout.
- **Open .xlsx...** loads a readable local workbook with the same SQL Server version sheets and column headers as the Microsoft workbook.
- A failed refresh keeps the last successful data for this SSMS session and labels it as previously loaded. It does not persist a cache across restarts.
- Missing worksheets, malformed rows, unavailable dates and invalid links produce bounded diagnostics. Valid builds remain available. Unknown dates display as `Unknown` and export as SQL `NULL`.
- Copy as TSQL is disabled until data exists. KB links accept HTTP and HTTPS only.

An encrypted/protected Excel package or legacy `.xls` cannot be read by the Open XML parser. Obtain an unencrypted `.xlsx` from the workbook publisher or use another readable copy you are authorized to access. Renaming a file does not change its format. HTML responses, empty files and damaged ZIP packages also produce explicit errors.

The Microsoft download inspected on 2026-09-11 had an OLE compound-file header and an `EncryptedPackage` stream instead of a readable Open XML ZIP package. This fix reports that condition; it does not decrypt protected workbooks.

## Automated parser checks

With the .NET 8 SDK installed:

```sh
dotnet run --project tests/SqlServerBuilds.Tests
```

The executable runs 21 deterministic checks against generated workbook fixtures and malformed inputs, without SSMS or a live download. An optional path adds a regression check for the protected Microsoft response observed above:

```sh
dotnet run --project tests/SqlServerBuilds.Tests -- /path/to/protected-download.xlsx
```

That optional assertion specifically expects a protected/legacy-format error; it is not a permanent assertion about what Microsoft serves.

## Windows / SSMS verification

Build the VSIX with the repository's normal Visual Studio toolchain, then verify:

1. Open SQL Server Builds while its initial download is running, and after a download/parse failure. The window must remain usable and show progress or a clear error.
2. Retry with the network unavailable; inspect the error details and retry availability.
3. Open a valid local workbook, then an unreadable workbook. Confirm that the prior data, original source and load time remain visible with the failure message.
4. Open a partially malformed workbook. Confirm that readable builds remain visible, diagnostics identify skipped rows/sheets, and missing dates and URLs are handled without exceptions.
5. Check light/dark themes, resizing, Copy as TSQL, and closing/reopening the window during refresh.

The parser tests and .NET Framework 4.7.2/WPF reference compilation were checked in Linux. A full VSIX build and the interactive SSMS checks above require Windows and were not run there.
