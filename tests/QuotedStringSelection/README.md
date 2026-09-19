# Quoted-string selection checks

Run the SQL-boundary regression tests with the .NET 8 SDK:

```sh
dotnet run --project tests/QuotedStringSelection/QuotedStringSelection.Tests.csproj
```

The console runner links the production resolver and checks every character position in each fixture, including delimiters and EOF. Cases cover hyphens, spaces, Unicode prefixes, escaped quotes, multiline strings, comments, bracketed/quoted identifiers, multiple literals, empty strings, unfinished literals, and edits that move token offsets.

## SSMS 22 integration checks

Build and install the VSIX, restart SSMS, and open Settings > Editor. The feature is enabled by default and can be toggled without reopening query windows. Its setting is independent of snippets.

- Double-click within `SELECT 'abc-def-123', N'O''Brien-Smith';`. Only the complete contents of the clicked string should be selected. Include clicks directly on a hyphen or escaped quote.
- Check spaces, multiline strings, horizontal/vertical scrolling, zoom, wrapped lines, and split editor panes.
- Verify normal selection in comments, identifiers, empty/unfinished literals, and outside quotes. Verify Ctrl/Shift/Alt-click, triple-click, and dragging retain normal editor behavior.
- Disable the feature, save, and try the same open query window. Re-enable it and repeat. Test with snippets disabled.
- Edit text before the string, then double-click again to verify cached offsets are invalidated.
- Open a script larger than 1,048,576 characters. Normal SSMS selection is intentional to bound synchronous tokenization on the editor thread.
- For a registration problem, enable NLog debug output and look for the quoted-string selection attachment message with the actual editor content type. The provider targets the `SQL` and `SQL Server Tools` document content types through the VSIX MEF component asset.

These tests do not execute SQL or change query text. Automated resolver checks do not validate SSMS mouse-event ordering or VSIX/MEF discovery; those require Windows/SSMS.
