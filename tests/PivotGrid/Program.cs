using AxialSqlTools.PivotGrid;
using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;

internal static class Program
{
    private static int checks;
    private static readonly PivotField[] Fields = {
        new PivotField(0, "Region", typeof(string)), new PivotField(1, "Period", typeof(string)),
        new PivotField(2, "Amount", typeof(decimal)), new PivotField(3, "Name", typeof(string)) };
    private static readonly PivotSnapshot Source = new PivotSnapshot(Fields, new[] {
        new[] { "East", "A", "10", "alpha" }, new[] { "East", "A", "30", "alpha" },
        new[] { "East", "B", "60", "beta" }, new[] { "West", "A", "100", "beta" },
        new[] { "West", "B", "NULL", "NULL" } });

    private static PivotRequest Request(PivotAggregation aggregation) => new PivotRequest {
        Rows = new[] { 0 }, Columns = new[] { 1 }, Value = 2, Aggregation = aggregation };
    private static DataTable Build(PivotRequest request, PivotSnapshot source = null) =>
        PivotEngine.Build(source ?? Source, request, CancellationToken.None).Table;
    private static object Total(DataTable table) => table.Rows[table.Rows.Count - 1][table.Columns.Count - 1];
    private static void Equal(object expected, object actual, string label)
    {
        checks++;
        if (!Equals(expected, actual)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }
    private static void Reject<T>(Action action, string label) where T : Exception
    {
        checks++;
        try { action(); } catch (T) { return; }
        throw new Exception(label + ": expected " + typeof(T).Name);
    }

    private static void DrillDownChecks()
    {
        var request = Request(PivotAggregation.Sum);
        var result = PivotEngine.Build(Source, request, CancellationToken.None);
        Equal(2, result.GetUnderlyingRows(0, 1, CancellationToken.None).Count, "East/A underlying rows");
        Equal(3, result.GetUnderlyingRows(0, 3, CancellationToken.None).Count, "Row total includes every period");
        Equal(3, result.GetUnderlyingRows(2, 1, CancellationToken.None).Count, "Column total includes every region");
        Equal(5, result.GetUnderlyingRows(2, 3, CancellationToken.None).Count, "Grand total includes null measure rows");
        Equal(1, result.GetUnderlyingRows(1, 2, CancellationToken.None).Count, "Null-valued group retains its source row");
        Equal("NULL", result.GetUnderlyingRows(1, 2, CancellationToken.None).GetPage(0).Rows[0][3], "Original source text retained");
        Equal(false, result.CanDrillDown(0, 0), "Grouping labels are not value cells");
        Reject<ArgumentOutOfRangeException>(() => result.GetUnderlyingRows(0, 0, CancellationToken.None), "Reject grouping cell");
        Reject<OperationCanceledException>(() => result.GetUnderlyingRows(0, 1, new CancellationToken(true)), "Drill cancellation");
        request.Rows[0] = 3; request.FilterField = 0; request.FilterText = "West";
        Equal(2, result.GetUnderlyingRows(0, 1, CancellationToken.None).Count, "Result retains applied request and axes");
        var filtered = Request(PivotAggregation.DistinctCount); filtered.Value = 3; filtered.FilterField = 0; filtered.FilterText = "east";
        var filteredResult = PivotEngine.Build(Source, filtered, CancellationToken.None);
        Equal(3, filteredResult.GetUnderlyingRows(1, 3, CancellationToken.None).Count, "Distinct drill shows duplicates within applied filter");
        var emptyIntersection = new PivotSnapshot(Fields, new[] { Source.Rows[0], Source.Rows[4] });
        var sparse = PivotEngine.Build(emptyIntersection, Request(PivotAggregation.CountRows), CancellationToken.None);
        Equal(0, sparse.GetUnderlyingRows(0, 2, CancellationToken.None).Count, "Empty intersection has no source rows");
        var numericGroups = new PivotSnapshot(Fields, new[] { new[] { "E", "A", "10.0", "x" }, new[] { "E", "A", "10.00", "x" } });
        var byMeasure = new PivotRequest { Rows = new[] { 2 }, Value = 2, Aggregation = PivotAggregation.Sum };
        Equal(2, PivotEngine.Build(numericGroups, byMeasure, CancellationToken.None).GetUnderlyingRows(0, 1, CancellationToken.None).Count,
            "Drill uses the same numeric key normalization as the cube");
        var nullGroups = new PivotSnapshot(Fields, new[] { new[] { (string)null, "A", "1", "x" }, new[] { "NULL", "A", "1", "x" } });
        var nullRequest = new PivotRequest { Rows = new[] { 0 } };
        Equal(2, PivotEngine.Build(nullGroups, nullRequest, CancellationToken.None).GetUnderlyingRows(0, 1, CancellationToken.None).Count,
            "Null keys use applied null handling");
        nullRequest.NullTextIsNull = false;
        Equal(1, PivotEngine.Build(nullGroups, nullRequest, CancellationToken.None).GetUnderlyingRows(0, 1, CancellationToken.None).Count,
            "Literal NULL and actual null remain separate when requested");
        var multi = Request(PivotAggregation.CountRows); multi.Rows = new[] { 0, 1 }; multi.Columns = new[] { 3 };
        var multiResult = PivotEngine.Build(Source, multi, CancellationToken.None);
        Equal(2, multiResult.GetUnderlyingRows(0, multiResult.Table.Columns.Count - 1, CancellationToken.None).Count,
            "Multiple row dimensions restrict detail rows");
        var scalar = PivotEngine.Build(Source, new PivotRequest(), CancellationToken.None);
        Equal(5, scalar.GetUnderlyingRows(0, 1, CancellationToken.None).Count, "Scalar total drills into every source row");
        var many = new PivotSnapshot(Fields, Enumerable.Range(0, 450).Select(i => new[] { "E", "A", i.ToString(), "x" }).ToArray());
        var all = PivotEngine.Build(many, new PivotRequest(), CancellationToken.None).GetUnderlyingRows(0, 1, CancellationToken.None);
        Equal(450, all.Count, "Details are not truncated to the first page");
        Equal(3, all.PageCount, "Page count");
        Equal(200, all.GetPage(0).Rows.Count, "Bounded first page");
        Equal(201, all.GetPage(1).Rows[0][0], "Original source row number on next page");
        Equal(50, all.GetPage(2).Rows.Count, "Last page");
        Reject<ArgumentOutOfRangeException>(() => all.GetPage(3), "Reject invalid detail page");
    }

    private static void Main()
    {
        var sum = Build(Request(PivotAggregation.Sum));
        Equal(200m, Total(sum), "Sum grand total");
        Equal(40m, sum.Rows[0][1], "East/A sum");
        Equal(DBNull.Value, sum.Rows[1][2], "All-null sum remains empty");
        Equal(50m, Total(Build(Request(PivotAggregation.Average))), "Weighted average total");
        Equal(100m / 3m, Build(Request(PivotAggregation.Average)).Rows[0][3], "Weighted row average");
        Equal(10m, Total(Build(Request(PivotAggregation.Minimum))), "Minimum");
        Equal(100m, Total(Build(Request(PivotAggregation.Maximum))), "Maximum");
        Equal(5UL, Total(Build(Request(PivotAggregation.CountRows))), "Count rows includes nulls");
        Equal(4UL, Total(Build(Request(PivotAggregation.CountValues))), "Count values excludes nulls");
        var distinct = Request(PivotAggregation.DistinctCount); distinct.Value = 3;
        Equal(2, Total(Build(distinct)), "Distinct total is not sum of distinct subtotals");
        distinct.NullTextIsNull = false;
        Equal(3, Total(Build(distinct)), "Literal NULL text can be included");
        var numericNulls = Request(PivotAggregation.Sum); numericNulls.NullTextIsNull = false;
        Equal(200m, Total(Build(numericNulls)), "Numeric NULL remains null");
        var filter = Request(PivotAggregation.Sum); filter.FilterField = 0; filter.FilterText = "EAST";
        Equal(100m, Total(Build(filter)), "Case-insensitive source filter");
        filter.FilterText = "missing";
        Equal(DBNull.Value, Total(Build(filter)), "Empty sum");
        filter.Aggregation = PivotAggregation.CountRows;
        Equal(0UL, Total(Build(filter)), "Empty count");
        Equal(0, PivotEngine.Build(Source, filter, CancellationToken.None).MatchedRows, "Filtered source row count");
        var scalar = new PivotRequest();
        Equal(5UL, Total(Build(scalar)), "No axes returns scalar total");
        var onlyColumns = Request(PivotAggregation.Sum); onlyColumns.Rows = new int[0];
        Equal(140m, Build(onlyColumns).Rows[0][1], "Column-only pivot");
        var onlyRows = Request(PivotAggregation.Sum); onlyRows.Columns = new int[0];
        Equal(100m, Build(onlyRows).Rows[0][1], "Row-only pivot");
        var multi = Request(PivotAggregation.CountRows); multi.Rows = new[] { 0, 1 }; multi.Columns = new int[0];
        Equal(5, Build(multi).Rows.Count, "Multiple row dimensions and total");
        var duplicate = new PivotSnapshot(new[] { new PivotField(0, "x].[y", typeof(string)),
            new PivotField(1, "x].[y", typeof(decimal)) }, new[] { new[] { "A", "12" } });
        Equal(12m, Total(Build(new PivotRequest { Rows = new[] { 0 }, Value = 1, Aggregation = PivotAggregation.Sum }, duplicate)),
            "Duplicate and binding-special headers remain independent");
        var overlap = Request(PivotAggregation.CountRows); overlap.Columns = new[] { 0 };
        Reject<InvalidOperationException>(() => Build(overlap), "Overlapping axes");
        var nonNumeric = Request(PivotAggregation.Sum); nonNumeric.Value = 3;
        Reject<InvalidOperationException>(() => Build(nonNumeric), "Non-numeric sums");
        var unparseable = new PivotSnapshot(Fields, new[] { new[] { "East", "A", "1e100", "name" } });
        Reject<InvalidOperationException>(() => Build(Request(PivotAggregation.Sum), unparseable), "Out-of-range numeric input");
        var sentinel = new PivotSnapshot(Fields, new[] { new[] { "East", "A", decimal.MinValue.ToString(CultureInfo.InvariantCulture), "name" } });
        Reject<InvalidOperationException>(() => Build(Request(PivotAggregation.Sum), sentinel), "NReco decimal sentinel not silently skipped");
        var overflow = new PivotSnapshot(Fields, new[] { new[] { "East", "A", decimal.MaxValue.ToString(), "name" },
            new[] { "East", "A", "1", "name" } });
        Reject<OverflowException>(() => Build(Request(PivotAggregation.Sum), overflow), "Overflow is reported");
        var wide = new PivotSnapshot(Fields, Enumerable.Range(0, 205).Select(i => new[] { "East", i.ToString(), "1", "n" }).ToArray());
        Reject<InvalidOperationException>(() => Build(Request(PivotAggregation.Sum), wide), "Wide output rejected");
        var dense = new PivotSnapshot(Fields, Enumerable.Range(0, 3000).Select(i => new[] { i.ToString(), (i % 180).ToString(), "1", "n" }).ToArray());
        Reject<InvalidOperationException>(() => Build(Request(PivotAggregation.Sum), dense), "Dense output rejected before rendering");
        Reject<OperationCanceledException>(() => PivotEngine.Build(Source, scalar, new CancellationToken(true)), "Cancellation");
        var culture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            var fractional = new PivotSnapshot(Fields, new[] { new[] { "East", "A", "1.25", "n" } });
            Equal(1.25m, Total(Build(Request(PivotAggregation.Sum), fractional)), "Invariant numeric parsing");
        }
        finally { CultureInfo.CurrentCulture = culture; }
        DrillDownChecks();
        var many = new PivotSnapshot(Fields, Enumerable.Range(0, 100000).Select(i =>
            new[] { (i % 100).ToString(), ((i / 100) % 20).ToString(), "1", "n" }).ToArray());
        var timer = Stopwatch.StartNew();
        Equal(100000m, Total(Build(Request(PivotAggregation.Sum), many)), "100,000-row aggregation");
        Console.WriteLine(checks + " checks passed. 100,000-row engine smoke test: " + timer.ElapsedMilliseconds + " ms.");
    }
}
