# Pivot Grid

Activate a completed query's window, then choose **Tools > Export Grid to Pivot Table** on the **Axial SQL Tools** toolbar. A new SSMS tab captures all rows of that grid, independently of the current cell selection. Each invocation gets its own snapshot. The tab includes the source session ID when available, for example **Pivot Grid (97)**. This ID is captured before opening the pivot; no additional query or connection is used. Changing or closing the source query after capture does not change the pivot.

Click fields under **Rows** and **Columns** to select or deselect them, choose an aggregation and a **Values** field, then **Apply**. Multiple grouping fields follow source column order. Leave Columns empty for a grouped summary, or clear both axes for a single total. The optional filter performs a case-insensitive “contains” search on one source column. Select output cells and use Ctrl+C to copy with headers.

Available calculations:

- Count rows, including rows with null values.
- Count non-null values and distinct non-null values.
- Sum, average, minimum and maximum of a numeric column.
- Row, column and grand totals. Averages and distinct totals are calculated from aggregate state, not by adding or averaging the visible summary values.

## Drill down into source rows

Double-click a value or total cell to open a large, resizable modal window centered over SSMS. You can also select the cell and press Enter, click **Show underlying rows**, or use the same command in its context menu. Grouping labels do not open details.

The window shows all original columns and a one-based **Source row** number from the captured grid. Its heading identifies the selected group and applied filter. Row totals include all column groups; column totals include all row groups; the grand total includes every source row allowed by the applied filter. Details include duplicate records and rows with null measures, even when the aggregation ignores null values or counts distinct values.

Matching runs in the background and can be cancelled. Pages contain up to 200 rows (fewer for very wide grids), with Previous/Next controls and an exact matching-row count. No matching rows are silently dropped. Ctrl+C copies selected cells from the current page. Numeric columns and the source row number are right-aligned with thousands separators. Resize or maximize the window for wide results; choose **Close** or press Escape to return to the pivot. The main pivot keeps its full display area.

Details use the configuration that produced the displayed pivot. Editing a selector or filter does not change them until **Apply** rebuilds the pivot. All rows come from the captured snapshot; drill-down does not execute SQL.

## Displayed values and limits

This feature uses SSMS's existing `GetCellDataAsString` grid API. It does not re-run SQL or make a database connection. Display truncation therefore applies. SQL NULL and the literal text `NULL` cannot be distinguished in text columns through this API. **Treat displayed NULL as null** is enabled by default; clear it to include those values as text. Numeric NULLs remain null in either mode. Non-numeric dimensions are grouped by their displayed text, with case-sensitive keys.

Numeric statistics use invariant decimal parsing and NReco's decimal arithmetic. Values outside the supported decimal range and overflowing totals produce an error instead of a partial result. SQL `decimal(38, ...)` and extreme floating-point values may need an explicit cast or rounding in the query first. Numeric-looking text columns must be cast to a numeric SQL type to use numeric statistics.

Capture stays on the UI thread and yields regularly for cancellation. Aggregation runs on a background thread; the output grid virtualizes rows and columns. Closing the tab cancels pending work and releases the snapshot. Capture rejects grids exceeding 1,000,000 rows, 5,000,000 cells, or 64 million displayed characters. Pivots allow up to six grouping fields, 100,000 groups, 200 output columns, and 500,000 output cells. These are memory limits, not sampling: reduce or filter the data if a limit is reached.

The MIT-licensed **NReco.PivotData 1.5.1** core supplies aggregation and pivot tables. The WPF interface belongs to Axial SQL Tools. The commercial NReco UI extensions are not used. The NReco license is included in the VSIX under `ThirdPartyNotices`.

## Verification

Run the standalone calculation checks with the .NET 8 SDK:

```sh
dotnet run --project tests/PivotGrid/PivotGrid.Tests.csproj -c Release
```

The production extension still targets .NET Framework 4.8. The standalone tests link the same engine source and package version.

For an SSMS smoke test, build the VSIX on Windows, verify that it contains `NReco.PivotData.dll` and its license, install it, restart SSMS, and run:

```sql
SELECT Region, Period, Amount, Name
FROM (VALUES
    (N'East', N'A', CAST(10 AS decimal(18,2)), N'alpha'),
    (N'East', N'A', 30, N'alpha'),
    (N'East', N'B', 60, N'beta'),
    (N'West', N'A', 100, N'beta'),
    (N'West', N'B', NULL, NULL)
) AS d(Region, Period, Amount, Name);

-- A second grid verifies that the command captures only the clicked grid.
SELECT N'Other grid' AS Region, 999 AS Amount;
```

1. Open Pivot Grid on the first grid. Group by Region in Rows and Period in Columns, with Sum of Amount. East/A should be 40; the grand total should be 200.
2. Change to Average. The grand total should be 50, with East's row total 33.333..., rather than an average of the displayed group averages.
3. Choose Distinct count of Name. The grand total should be 2, excluding SQL NULL. Filter Region by `east`; the source match count should become 3.
4. Open a separate pivot on the second grid. Its total should be 999. The first pivot should retain its original data after its query window closes.
5. Check an empty grid, duplicate/unnamed column headers, SQL NULL versus literal `NULL`, and a wide/high-cardinality result. Errors should be readable and should not show partial statistics.
6. Double-click East/A and verify the two source rows (10 and 30). Drill into East's total (3 rows), the A total (3 rows), and the grand total (5 rows, including the null measure). Try Enter and the context menu, reorder pivot columns, and verify that the selected cell still opens the correct group.
7. Close details, change a filter without applying it, and reopen details to verify that they still match the displayed pivot. Close the modal, apply the new filter, and reopen details to confirm they use it. Test Previous/Next on a group with more than 200 source rows, an empty intersection, and resizing/maximizing the modal, right-aligned numeric columns, paging, Escape/Close, and SSMS ownership (the main window should be disabled while details are open).
8. Cancel during a large capture and during aggregation; close a busy pivot tab. SSMS should remain usable. Verify light/dark themes, horizontal scrolling, and Ctrl+C from the output.

SSMS hosting, COM menu events, actual grid extraction, theme rendering, XAML compilation, and VSIX packaging require this Windows smoke test; pure engine tests cannot validate them.

## Value sorting

Choose **None**, **Ascending**, or **Descending** in **Sort values**. Sorting applies immediately to the displayed results and remains selected after Apply. Grouped summaries sort by **Value**; cross-tab pivots sort by the row **Grand total** using the selected aggregation. Values sort numerically, with nulls first in ascending order and last in descending order; equal values retain their original order. The grand-total row is pinned below the scrollable results and remains visible while scrolling vertically. Its columns follow horizontal scrolling, resizing, and reordering in the results grid. Double-click a total to drill into its underlying rows. **None** restores the original pivot order. Drill-down continues to use the selected group after sorting.

## Source selection

The result grid does not need keyboard focus: you can invoke the command while editing SQL. If a result grid has focus, that grid is used; otherwise the first populated result grid in the active query window is used. Focus a specific grid when a query returns several result sets and you want a different one. With no active query window or no populated grid, an informational message explains what to do.
