# Pivot Grid

Open **Tools > Export Grid to Pivot Table** from a query with completed grid results, or use the same command on the result grid's context menu. Pivot Grid opens in an SSMS document tab and uses a captured snapshot. Applying a layout never re-runs SQL.

## Build a layout

- Search **Available fields**, select a field, then choose **Rows**, **Columns**, or **Values**. Fields can also be dragged into these areas. Double-clicking a catalog field adds it to Rows.
- Reorder grouping fields with the arrow buttons, Alt+Up/Down, or drag and drop. Moving a field to the other grouping axis removes it from its previous axis. A field can still be both a grouping field and a value.
- Use **Add measure** to compare several calculations together, such as Sum(Duration), Average(Duration), and Count rows. Each measure has its own aggregation and source field. Numeric calculations offer numeric fields only.
- Use **Add count** to add a row count. Remove individual measures with their remove button. At least one measure is always retained.
- **Swap rows/columns** exchanges the grouping axes. **Reset** restores an ungrouped count and clears filters. Both update setup; use Apply to update results.
- Collapse **Fields and filters** to give the result grid more space.

Use **Apply** or **Ctrl+Enter** to calculate. The summary describes the applied layout. A **Changes pending** message means the controls contain unapplied changes and the grid still shows the previous result. Validation errors, cancelled calculations, and failed calculations retain the last successful result.

The first opening starts with an ungrouped count, avoiding an expensive automatic grouping on a unique identifier. Later openings restore the last successfully applied layout for exactly matching column names, order, and types.

## Filters

Filters apply to source rows before aggregation. Every filter must match (AND).

- **Selected values** provides a searchable checklist. Search does not clear hidden selections; **Select all shown** and **Clear shown** affect only values matching the search. At least one value must be selected.
- **Contains**, **Equals**, and **Does not equal** compare text without case sensitivity. Selected-value checklists match exact displayed values, including case.
- Numeric fields support equality and greater/less comparisons. Use a period for decimals; scientific notation is supported. Numeric comparisons reject values that cannot be represented exactly as a decimal.
- **Is NULL** and **Is not NULL** handle missing values explicitly. Ordinary non-null comparisons exclude NULL rows.
- Click a filter chip to edit it, remove it individually, or use **Clear all**.

The **Advanced** setting controls whether displayed `NULL` text counts as null. SSMS display strings cannot distinguish SQL NULL from literal text `NULL` in text columns. Numeric NULLs always count as null. Empty strings remain separate values.

## Sort and inspect

Click any result column header to cycle through ascending, descending, and original order. Numeric measures and numeric grouping labels sort numerically. Date/time and boolean grouping fields use their type where their displayed text can be parsed. The sort arrow indicates the active column.

Grouping labels remain frozen and the grand-total row stays pinned at the bottom. Every measure has its own row, column, and grand totals. Average totals are calculated from underlying rows, and distinct-count totals are recomputed over the relevant set rather than added together.

Double-click a value or total, press Enter, or use **Show underlying rows**. Drill-down always uses the applied filters and the selected cell's original group, even after sorting or column reordering. Existing page navigation and Ctrl+C copying remain available.

## Save layouts

Enter a name in **Layout** and choose **Save**. Choose a saved name and **Load** to restore it into setup, then Apply. Saving an existing name asks before replacing it; Delete removes a named layout.

Named layouts and the last applied layout are matched to the exact result schema, including duplicate column positions and data types. Only configuration and filter criteria are saved in the existing per-user settings file. Captured source rows are never persisted. A layout saved before Apply can differ from the currently displayed result.

## Limits

Existing source snapshot limits remain: up to 1,000,000 rows, 5,000,000 cells, and 64,000,000 characters. Pivot output is bounded to 200 columns, 500,000 cells, and 100,000 aggregation groups across measures, with up to six grouping fields. More measures consume more of these limits. Oversized results fail with guidance instead of silently sampling rows.

SSMS display-text truncation still applies. Values outside decimal range or precision require casting or rounding in the source SQL before numeric calculation. Numeric grouping labels retain their full displayed precision when sorting.

See [regression checks and SSMS validation](../tests/PivotGrid/README.md).
