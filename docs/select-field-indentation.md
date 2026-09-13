# SELECT field indentation after DISTINCT and TOP

Enable **Use standard field indentation after DISTINCT/TOP** alongside the other advanced formatting options in Settings. It is also available in the **Shift+Format** options dialog and participates in that dialog's Check all / Uncheck all actions.

The option is disabled by default, saved with the other formatting settings, and reflected immediately in the settings preview. The existing serialized key `breakSelectFieldsAfterTopAndUnindent` is retained for compatibility.

For SELECT statements containing DISTINCT, TOP, or both, every field starts on a new line with one normal indentation level relative to SELECT. The formatter's normal indentation size is used, currently four spaces:

```sql
SELECT DISTINCT TOP (100)
    CustomerId,
    CustomerName,
    CreatedAt
FROM dbo.Customers;
```

This applies inside CTEs, derived tables, scalar subqueries, unions and procedural blocks. TOP expressions, PERCENT and WITH TIES remain part of the SELECT header. Plain SELECT statements retain their existing layout. When the option is disabled, the existing formatter layout is used.

Multiline field expressions move with their continuation lines; nested SELECTs then receive their own indentation. Only whitespace changes. String literals, identifiers and expression contents remain intact. With Preserve comments enabled, comment contents are retained, and standalone comments between fields align with the field list. The pass runs after the other advanced formatting options so it uses the actual positions after CASE and BEGIN/END formatting.

## Validation

Run the regression executable with the .NET 8 SDK:

```sh
dotnet run --project tests/SelectFieldIndentation.Tests
```

The 41 checks exercise the production formatter, comment interleaver and settings model. They verify output parsing, unchanged SQL/comment tokens, field columns, repeated formatting, old/default settings and JSON round trips. Fixtures cover nested queries, TOP variants, comments, multiline literals and combinations with the other advanced options.

The production formatter and settings code, FormatOptionsDialog code, and settings load/save/preview methods also compile against .NET Framework 4.7.2 and WPF reference assemblies with a host harness. A full Windows VSIX build and interactive SSMS checks remain to be run:

1. Toggle the option in Settings and confirm the preview changes immediately; save and reopen Settings to verify persistence.
2. Hold Shift while invoking Format, toggle the option, and verify Check all / Uncheck all include it.
3. Format DISTINCT-only, TOP-only and combined queries, including nested queries and CTEs. Repeating Format should preserve the layout.
4. Disable the option and confirm the previous SELECT layout is restored.
