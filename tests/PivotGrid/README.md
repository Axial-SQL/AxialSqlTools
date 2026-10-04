# Pivot Grid regression checks

Run from the repository root with the .NET 8 SDK:

```powershell
dotnet run --project tests/PivotGrid/PivotGrid.Regression.csproj --configuration Release
```

This console harness exits with a nonzero code on failure. It links the production
`PivotEngine`, `PivotPresentation` and `PivotLayoutStore` files directly, enforces
C# 7.3, and uses the actual NReco.PivotData 1.5.1 and Newtonsoft.Json 13.0.3 packages.
Only the settings file boundary is replaced with an in-memory stub. Tests never
read or write the user's extension settings or run GitHub workflows.

The 20 test groups cover:

- All seven aggregations together, weighted averages, distinct grand totals and
  drill-down identity for every measure, column group and total.
- Empty inputs, missing intersections, zero/one/two axes and explicit axis order.
- All filter operators, combined filters, null handling and exact value lists.
- Applied-request isolation, invalid requests, cancellation and measure-aware
  output limits.
- Decimal precision loss and underflow rejection, equivalent numeric formats,
  and preserving raw grouping keys when the grouped field is also aggregated.
- Header sorting, stable ties, restoring original order, pinned total exclusion,
  drill-down identity, decimal38 labels, scientific notation, dates and booleans.
- Nested pending-change comparison, layout round-trips, schema isolation,
  case-insensitive preset replacement, corruption preservation and save failures.

Validation during implementation: **20/20 test groups passed** against the real
packages. The engine, presentation, layout store and all four Pivot Grid WPF
code-behind files also compiled with Roslyn C# 7.3 against the actual .NET Framework
4.8 reference assemblies, including WPF, and the packages' .NET Framework DLLs.
That focused compile generated named fields and event assignments from XAML;
it stubbed only SSMS host interfaces, package scheduling and theme integration.
It was a source/API check, not a complete VSIX build or a WPF rendering test.

For a full build, use the existing solution and required SSMS/Visual Studio
dependencies on Windows. Workflows remain manually triggered.

Before release, verify these interactions inside SSMS:

1. Search fields; add, drag, reorder and remove row/column fields; repeat with
   keyboard controls. Add Count, Sum and Average for the same result set.
2. Change a field or filter and confirm the pending indicator appears while the
   prior result remains usable. Apply, cancel and provoke a validation failure.
3. Sort grouping labels and individual measure columns in both directions; reset
   sort. Confirm the grand total remains pinned and drilled rows match the cell.
4. Combine checklist, numeric and NULL filters. Search the checklist, clear only
   visible selections, remove a filter chip and clear all filters.
5. Collapse setup, swap axes, reset, save/replace/load/delete a named layout and
   reopen a result with the same schema to verify the remembered layout.
6. Check light/dark themes, keyboard focus, narrow document widths and scrolling.
   Open and close drill-down/filter windows, including during a long calculation.
