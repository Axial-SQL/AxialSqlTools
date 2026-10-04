using AxialSqlTools;
using AxialSqlTools.PivotGrid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;

internal static class Program
{
    private static int passed;
    private static int failed;
    private static readonly CancellationToken None = CancellationToken.None;

    private static int Main()
    {
        CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
        Run("Multiple measures, weighted totals, distinct totals and drill-down", MeasuresAndDrillDown);
        Run("Empty, ungrouped, row-only and column-only pivots", EmptyAndUngrouped);
        Run("Requested axis order is preserved", AxisOrder);
        Run("All filter operators, combined filters and null semantics", Filters);
        Run("Applied request is isolated from subsequent edits", RequestIsolation);
        Run("Invalid configuration is rejected before processing", Validation);
        Run("Numeric aggregation and filtering reject precision loss", NumericPrecision);
        Run("Grouping and summing the same field preserves raw group keys", RawNumericKeys);
        Run("Missing cross-tab intersections retain correct empty values", MissingIntersections);
        Run("Cancellation stops build and drill-down", Cancellation);
        Run("Column limits include every measure and total", MultiMeasureLimits);
        Run("Header sorting excludes totals and preserves drill-down identity", Sorting);
        Run("Numeric grouping sort preserves decimal38 precision", NumericSorting);
        Run("Date and boolean grouping sorts use field types", TypedSorting);
        Run("Pending-change comparison checks nested configuration", RequestComparison);
        Run("Layouts round-trip and isolate returned objects", LayoutRoundTrip);
        Run("Layout schema includes ordinal, name and type", LayoutSchema);
        Run("Named layouts replace case-insensitively and delete", LayoutNames);
        Run("Corrupt or unsupported layout data is preserved", LayoutCorruption);
        Run("Layout save failures are reported", LayoutSaveFailure);
        Console.WriteLine("\n" + passed + " passed; " + failed + " failed.");
        return failed == 0 ? 0 : 1;
    }

    private static void Run(string name, Action action)
    {
        try { action(); passed++; Console.WriteLine("PASS " + name); }
        catch (Exception exception) { failed++; Console.WriteLine("FAIL " + name + "\n" + exception); }
    }

    private static PivotSnapshot Source() => new PivotSnapshot(new[]
    {
        new PivotField(0, "Region", typeof(string)), new PivotField(1, "Quarter", typeof(string)),
        new PivotField(2, "Amount", typeof(decimal)), new PivotField(3, "Customer", typeof(string))
    }, new[]
    {
        new[] { "East", "Q1", "10", "A" }, new[] { "East", "Q1", "30", "B" },
        new[] { "East", "Q2", "20", "A" }, new[] { "West", "Q1", "100", "A" },
        new[] { "West", "Q2", "NULL", null }, new[] { "West", "Q2", "50", "NULL" },
        new[] { "East", "Q2", "-10", "B" }
    });

    private static PivotMeasure Measure(PivotAggregation aggregation, int value = -1) =>
        new PivotMeasure { Aggregation = aggregation, Value = value };

    private static PivotRequest Request() => new PivotRequest
    {
        Rows = new[] { 0 }, Columns = new[] { 1 },
        Measures = new[] { Measure(PivotAggregation.CountRows), Measure(PivotAggregation.Sum, 2),
            Measure(PivotAggregation.Average, 2), Measure(PivotAggregation.DistinctCount, 3),
            Measure(PivotAggregation.CountValues, 3), Measure(PivotAggregation.Minimum, 2),
            Measure(PivotAggregation.Maximum, 2) }
    };

    private static void MeasuresAndDrillDown()
    {
        var source = Source(); var request = Request(); var result = PivotEngine.Build(source, request, None);
        Equal(7, result.MatchedRows); Equal(1, result.RowColumnCount); Equal(3, result.Table.Rows.Count);
        Equal(22, result.Table.Columns.Count);
        for (int row = 0; row < result.Table.Rows.Count; row++)
        {
            for (int column = result.RowColumnCount; column < result.Table.Columns.Count; column++)
            {
                var measure = request.Measures[(column - result.RowColumnCount) % request.Measures.Length];
                var ids = UnderlyingIds(result, row, column);
                var expected = Aggregate(source, ids, measure);
                Equal(expected, result.Table.Rows[row][column]);
                if ((column - result.RowColumnCount) % request.Measures.Length != 0)
                    Sequence(UnderlyingIds(result, row, column - 1), ids);
                Equal(column >= result.Table.Columns.Count - request.Measures.Length, result.IsTotalColumn(column));
            }
        }
        int total = result.Table.Rows.Count - 1; int start = result.Table.Columns.Count - request.Measures.Length;
        Equal(200m, result.Table.Rows[total][start + 1]);
        Equal(200m / 6m, result.Table.Rows[total][start + 2]);
        Equal(2m, result.Table.Rows[total][start + 3]);
        True(result.Table.Columns.Cast<DataColumn>().Skip(1).All(c => !string.IsNullOrWhiteSpace(c.Caption) && c.Caption != "Value"));
        False(result.CanDrillDown(0, 0));
    }

    private static object Aggregate(PivotSnapshot source, int[] ids, PivotMeasure measure)
    {
        if (measure.Aggregation == PivotAggregation.CountRows) return (decimal)ids.Length;
        var texts = ids.Select(i => source.Rows[i - 1][measure.Value]).Where(s => s != null && s != "NULL").ToArray();
        if (measure.Aggregation == PivotAggregation.CountValues) return (decimal)texts.Length;
        if (measure.Aggregation == PivotAggregation.DistinctCount) return (decimal)texts.Distinct().Count();
        if (texts.Length == 0) return DBNull.Value;
        var numbers = texts.Select(s => decimal.Parse(s, CultureInfo.InvariantCulture)).ToArray();
        switch (measure.Aggregation)
        {
            case PivotAggregation.Sum: return numbers.Sum();
            case PivotAggregation.Average: return numbers.Average();
            case PivotAggregation.Minimum: return numbers.Min();
            case PivotAggregation.Maximum: return numbers.Max();
            default: throw new InvalidOperationException();
        }
    }

    private static void EmptyAndUngrouped()
    {
        var source = Source();
        foreach (var axes in new[] { new[] { false, false }, new[] { true, false }, new[] { false, true }, new[] { true, true } })
        {
            var request = Request(); request.Rows = axes[0] ? new[] { 0 } : new int[0]; request.Columns = axes[1] ? new[] { 1 } : new int[0];
            var result = PivotEngine.Build(source, request, None);
            int last = result.Table.Rows.Count - 1; int start = result.Table.Columns.Count - request.Measures.Length;
            Equal(7m, result.Table.Rows[last][start]); Equal(200m, result.Table.Rows[last][start + 1]);
            Sequence(Enumerable.Range(1, 7), UnderlyingIds(result, last, start + 2));
            var empty = PivotEngine.Build(new PivotSnapshot(source.Fields, new string[0][]), request, None);
            Equal(0, empty.MatchedRows); Equal(1, empty.Table.Rows.Count);
            Equal(0m, empty.Table.Rows[0][empty.Table.Columns.Count - request.Measures.Length]);
            Equal(DBNull.Value, empty.Table.Rows[0][empty.Table.Columns.Count - request.Measures.Length + 1]);
        }
    }

    private static void AxisOrder()
    {
        var request = Request(); request.Rows = new[] { 1, 0 }; request.Columns = new int[0];
        var result = PivotEngine.Build(Source(), request, None);
        True(result.Table.Columns[0].Caption.Contains("Quarter")); True(result.Table.Columns[1].Caption.Contains("Region"));
        Equal(2, result.RowColumnCount);
        for (int row = 0; row < result.Table.Rows.Count - 1; row++)
        {
            foreach (int id in UnderlyingIds(result, row, result.RowColumnCount))
            { Equal(Source().Rows[id - 1][1], result.Table.Rows[row][0]); Equal(Source().Rows[id - 1][0], result.Table.Rows[row][1]); }
        }
    }

    private static void Filters()
    {
        CheckFilter(0, PivotFilterOperator.Contains, "as", null, new[] { 1, 2, 3, 7 });
        CheckFilter(0, PivotFilterOperator.Equals, "east", null, new[] { 1, 2, 3, 7 });
        CheckFilter(0, PivotFilterOperator.NotEquals, "East", null, new[] { 4, 5, 6 });
        CheckFilter(3, PivotFilterOperator.In, null, new[] { "A" }, new[] { 1, 3, 4 });
        CheckFilter(3, PivotFilterOperator.IsNull, null, null, new[] { 5, 6 });
        CheckFilter(3, PivotFilterOperator.IsNotNull, null, null, new[] { 1, 2, 3, 4, 7 });
        CheckFilter(2, PivotFilterOperator.GreaterThan, "20", null, new[] { 2, 4, 6 });
        CheckFilter(2, PivotFilterOperator.GreaterThanOrEqual, "20", null, new[] { 2, 3, 4, 6 });
        CheckFilter(2, PivotFilterOperator.LessThan, "20", null, new[] { 1, 7 });
        CheckFilter(2, PivotFilterOperator.LessThanOrEqual, "20", null, new[] { 1, 3, 7 });
        CheckFilter(3, PivotFilterOperator.In, null, new string[] { null, "B" }, new[] { 2, 5, 6, 7 });
        var request = Request(); request.Filters = new[]
        {
            new PivotFilter { Field = 0, Operator = PivotFilterOperator.Equals, Text = "East" },
            new PivotFilter { Field = 2, Operator = PivotFilterOperator.GreaterThan, Text = "15" }
        };
        var result = PivotEngine.Build(Source(), request, None);
        Equal(2, result.MatchedRows); Sequence(new[] { 2, 3 }, UnderlyingIds(result, result.Table.Rows.Count - 1, result.Table.Columns.Count - 1));
        request.Filters = new[] { new PivotFilter { Field = 3, Operator = PivotFilterOperator.IsNull } };
        request.NullTextIsNull = false; result = PivotEngine.Build(Source(), request, None); Equal(1, result.MatchedRows);
        request.Filters = new[] { new PivotFilter { Field = 2, Operator = PivotFilterOperator.IsNull } };
        result = PivotEngine.Build(Source(), request, None); Equal(1, result.MatchedRows);
    }

    private static void CheckFilter(int field, PivotFilterOperator op, string text, string[] values, int[] expected)
    {
        var request = Request(); request.Filters = new[] { new PivotFilter { Field = field, Operator = op, Text = text, Values = values } };
        var result = PivotEngine.Build(Source(), request, None); Equal(expected.Length, result.MatchedRows);
        Sequence(expected, UnderlyingIds(result, result.Table.Rows.Count - 1, result.Table.Columns.Count - 1));
    }

    private static void RequestIsolation()
    {
        var request = Request(); request.Filters = new[] { new PivotFilter { Field = 3, Operator = PivotFilterOperator.In, Values = new[] { "A" } } };
        var result = PivotEngine.Build(Source(), request, None);
        request.Rows[0] = 3; request.Columns[0] = 2; request.Measures[0].Value = 3;
        request.Filters[0].Values[0] = "B"; request.Filters[0].Field = 0;
        Sequence(new[] { 1, 3, 4 }, UnderlyingIds(result, result.Table.Rows.Count - 1, result.Table.Columns.Count - 1));
    }

    private static void Validation()
    {
        var source = Source();
        Action<PivotRequest> bad = candidate => Throws<InvalidOperationException>(() => PivotEngine.Validate(source, candidate));
        var request = Request(); request.Rows = new[] { 0, 0 }; bad(request);
        request = Request(); request.Columns = new[] { 0 }; bad(request);
        request = Request(); request.Rows = new[] { 40 }; bad(request);
        request = Request(); request.Measures = new PivotMeasure[0]; bad(request);
        request = Request(); request.Measures = new[] { Measure(PivotAggregation.Sum, 0) }; bad(request);
        request = Request(); request.Measures = new[] { Measure(PivotAggregation.Sum, -1) }; bad(request);
        request = Request(); request.Measures = new[] { Measure((PivotAggregation)100, 2) }; bad(request);
        request = Request(); request.Filters = new[] { new PivotFilter { Field = 100, Operator = PivotFilterOperator.Contains, Text = "a" } }; bad(request);
        request = Request(); request.Filters = new[] { new PivotFilter { Field = 0, Operator = PivotFilterOperator.GreaterThan, Text = "4" } }; bad(request);
        request = Request(); request.Filters = new[] { new PivotFilter { Field = 2, Operator = PivotFilterOperator.GreaterThan, Text = "bad" } }; bad(request);
        request = Request(); request.Filters = new[] { new PivotFilter { Field = 2, Operator = (PivotFilterOperator)100 } }; bad(request);
        request = Request(); request.Filters = new[] { new PivotFilter { Field = 3, Operator = PivotFilterOperator.In, Values = new string[0] } }; bad(request);
        request = Request(); request.Filters = new[] { new PivotFilter { Field = 3, Operator = PivotFilterOperator.In, Values = null } }; bad(request);
        var invalid = Source(); invalid.Rows[0][2] = "not numeric";
        Throws<InvalidOperationException>(() => PivotEngine.Build(invalid, Request(), None));
    }

    private static void Cancellation()
    {
        var token = new CancellationToken(true);
        Throws<OperationCanceledException>(() => PivotEngine.Build(Source(), Request(), token));
        var result = PivotEngine.Build(Source(), Request(), None);
        Throws<OperationCanceledException>(() => result.GetUnderlyingRows(0, 1, token));
    }

    private static void NumericPrecision()
    {
        foreach (string number in new[] { "0.1234567890123456789012345678901", "0.1234567890123456789012345678902", "1e-100" })
        {
            var source = Source(); source.Rows[0][2] = number;
            Throws<InvalidOperationException>(() => PivotEngine.Build(source, Request(), None));
            var request = Request(); request.Measures = new[] { Measure(PivotAggregation.CountRows) };
            request.Filters = new[] { new PivotFilter { Field = 2, Operator = PivotFilterOperator.Equals, Text = number } };
            Throws<InvalidOperationException>(() => PivotEngine.Validate(Source(), request));
            request.Filters[0].Text = "0";
            Throws<InvalidOperationException>(() => PivotEngine.Build(source, request, None));
        }
        foreach (string number in new[] { "1.000", "1e0", "100e-2" })
        {
            var source = Source(); source.Rows[0][2] = number;
            var request = Request(); request.Filters = new[] { new PivotFilter { Field = 2, Operator = PivotFilterOperator.Equals, Text = "1" } };
            var result = PivotEngine.Build(source, request, None); Equal(1, result.MatchedRows);
        }
    }

    private static void RawNumericKeys()
    {
        var source = new PivotSnapshot(new[] { new PivotField(0, "Code", typeof(int)) }, new[] { new[] { "01" }, new[] { "1" }, new[] { "01" } });
        var request = new PivotRequest { Rows = new[] { 0 }, Measures = new[] { Measure(PivotAggregation.CountRows), Measure(PivotAggregation.Sum, 0) } };
        var result = PivotEngine.Build(source, request, None); Equal(3, result.Table.Rows.Count);
        int first = result.Table.Rows.IndexOf(result.Table.Rows.Cast<DataRow>().Single(row => Equals(row[0], "01")));
        int second = result.Table.Rows.IndexOf(result.Table.Rows.Cast<DataRow>().Single(row => Equals(row[0], "1")));
        Equal(2m, result.Table.Rows[first][2]); Equal(1m, result.Table.Rows[second][2]);
        Sequence(new[] { 1, 3 }, UnderlyingIds(result, first, 2)); Sequence(new[] { 2 }, UnderlyingIds(result, second, 2));
    }

    private static void MissingIntersections()
    {
        var source = Source(); source = new PivotSnapshot(source.Fields, source.Rows.Where(row => row[0] != "West" || row[1] != "Q2").ToArray());
        var request = Request(); var result = PivotEngine.Build(source, request, None);
        int rowIndex = result.Table.Rows.IndexOf(result.Table.Rows.Cast<DataRow>().Single(row => Equals(row[0], "West")));
        var columns = result.Table.Columns.Cast<DataColumn>().Where(column => column.Caption.Contains("Q2")).ToArray();
        Equal(request.Measures.Length, columns.Length);
        foreach (var column in columns)
        {
            Equal(0, result.GetUnderlyingRows(rowIndex, column.Ordinal, None).Count);
            int index = (column.Ordinal - result.RowColumnCount) % request.Measures.Length;
            Equal(Aggregate(source, new int[0], request.Measures[index]), result.Table.Rows[rowIndex][column]);
        }
    }

    private static void MultiMeasureLimits()
    {
        var source = new PivotSnapshot(Source().Fields, Enumerable.Range(0, 110).Select(i => new[] { "East", i.ToString(), "1", "A" }).ToArray());
        var request = Request(); request.Measures = new[] { Measure(PivotAggregation.CountRows) };
        var single = PivotEngine.Build(source, request, None); True(single.Table.Columns.Count <= PivotEngine.MaxColumns);
        request.Measures = new[] { Measure(PivotAggregation.CountRows), Measure(PivotAggregation.Sum, 2) };
        Throws<InvalidOperationException>(() => PivotEngine.Build(source, request, None));
    }

    private static void Sorting()
    {
        var source = Source(); var request = Request(); var result = PivotEngine.Build(source, request, None);
        int sum = result.Table.Columns.Count - request.Measures.Length + 1;
        var rows = PivotPresentation.SortRows(result, source, request, sum, ListSortDirection.Descending);
        Equal(2, rows.Length); Equal("West", rows[0][0]);
        Sequence(new[] { 4, 5, 6 }, UnderlyingIds(result, result.Table.Rows.IndexOf(rows[0].Row), sum));
        True(ReferenceEquals(rows[0].Row, result.Table.Rows.Cast<DataRow>().Single(row => Equals(row[0], "West"))));
        var reset = PivotPresentation.SortRows(result, source, request, sum, null);
        Sequence(result.Table.Rows.Cast<DataRow>().Take(result.Table.Rows.Count - 1), reset.Select(view => view.Row));
        source.Rows[0][0] = "Grand total"; result = PivotEngine.Build(source, request, None);
        rows = PivotPresentation.SortRows(result, source, request, 0, ListSortDirection.Ascending);
        True(rows.Any(view => Equals(view[0], "Grand total"))); Equal(result.Table.Rows.Count - 1, rows.Length);
    }

    private static void NumericSorting()
    {
        string[] labels = { "10", "2", "-20", "-3", "1e2", "0", "10000000000000000000000000000000000001", "10000000000000000000000000000000000000", "bad", null, "2.00" };
        var source = new PivotSnapshot(new[] { new PivotField(0, "Number", typeof(decimal)) }, labels.Select(s => new[] { s }).ToArray());
        var request = new PivotRequest { Rows = new[] { 0 }, Measures = new[] { Measure(PivotAggregation.CountRows) } };
        var result = PivotEngine.Build(source, request, None);
        var sorted = PivotPresentation.SortRows(result, source, request, 0, ListSortDirection.Ascending);
        var expected = new object[] { DBNull.Value, "-20", "-3", "0", "2", "2.00", "10", "1e2", "10000000000000000000000000000000000000", "10000000000000000000000000000000000001", "bad" };
        var ties = result.Table.Rows.Cast<DataRow>().Take(result.Table.Rows.Count - 1).Select(r => r[0]).Where(v => Equals(v, "2") || Equals(v, "2.00")).ToArray();
        expected[4] = ties[0]; expected[5] = ties[1]; Sequence(expected, sorted.Select(r => r[0]));
        var descending = PivotPresentation.SortRows(result, source, request, 0, ListSortDirection.Descending);
        Equal("bad", descending[0][0]); Equal(DBNull.Value, descending[descending.Length - 1][0]);
        Sequence(ties, descending.Where(r => Equals(r[0], "2") || Equals(r[0], "2.00")).Select(r => r[0]));
    }

    private static void TypedSorting()
    {
        foreach (var item in new[]
        {
            Tuple.Create(typeof(DateTime), new[] { "12/1/2020", "2/1/2020", "1/1/2021" }, new[] { "2/1/2020", "12/1/2020", "1/1/2021" }),
            Tuple.Create(typeof(bool), new[] { "True", "False" }, new[] { "False", "True" })
        })
        {
            var source = new PivotSnapshot(new[] { new PivotField(0, "Typed", item.Item1) }, item.Item2.Select(v => new[] { v }).ToArray());
            var request = new PivotRequest { Rows = new[] { 0 }, Measures = new[] { Measure(PivotAggregation.CountRows) } };
            var result = PivotEngine.Build(source, request, None);
            Sequence(item.Item3, PivotPresentation.SortRows(result, source, request, 0, ListSortDirection.Ascending).Select(r => (string)r[0]));
        }
    }

    private static void RequestComparison()
    {
        var request = Request(); request.Filters = new[] { new PivotFilter { Field = 3, Operator = PivotFilterOperator.In, Values = new string[] { "A", null } } };
        var copy = request.Copy(); True(PivotPresentation.RequestsEqual(request, copy));
        copy.Filters[0].Values[0] = "B"; False(PivotPresentation.RequestsEqual(request, copy));
        copy = request.Copy(); copy.Measures[1].Aggregation = PivotAggregation.Average; False(PivotPresentation.RequestsEqual(request, copy));
        copy = request.Copy(); copy.NullTextIsNull = false; False(PivotPresentation.RequestsEqual(request, copy));
        copy = request.Copy(); copy.Rows[0] = 3; False(PivotPresentation.RequestsEqual(request, copy));
    }

    private static void LayoutRoundTrip()
    {
        SettingsFileStore.Reset(); var source = Source(); var request = Request();
        True(PivotLayoutStore.SaveLast(source, request)); True(PivotPresentation.RequestsEqual(request, PivotLayoutStore.GetLast(source)));
        var copy = PivotLayoutStore.GetLast(source); copy.Measures[0].Value = 99;
        True(PivotPresentation.RequestsEqual(request, PivotLayoutStore.GetLast(source)));
        True(PivotLayoutStore.SaveNamed(source, "Analysis", request));
        False(SettingsFileStore.Values["PivotGridLayouts"].Contains("East"));
        False(SettingsFileStore.Values["PivotGridLayouts"].Contains("West"));
        Equal(1, PivotLayoutStore.GetLayouts(source).Count);
    }

    private static void LayoutSchema()
    {
        SettingsFileStore.Reset(); var source = Source(); True(PivotLayoutStore.SaveLast(source, Request()));
        var renamed = new PivotSnapshot(new[] { new PivotField(0, "Another", typeof(string)) }.Concat(source.Fields.Skip(1)).ToArray(), source.Rows);
        Equal(null, PivotLayoutStore.GetLast(renamed));
        var retyped = new PivotSnapshot(source.Fields.Take(2).Concat(new[] { new PivotField(2, "Amount", typeof(double)), source.Fields[3] }).ToArray(), source.Rows);
        Equal(null, PivotLayoutStore.GetLast(retyped));
        var reordered = new PivotSnapshot(new[] { new PivotField(0, "Quarter", typeof(string)), new PivotField(1, "Region", typeof(string)), source.Fields[2], source.Fields[3] }, source.Rows);
        Equal(null, PivotLayoutStore.GetLast(reordered));
        True(PivotPresentation.RequestsEqual(Request(), PivotLayoutStore.GetLast(Source())));
    }

    private static void LayoutNames()
    {
        SettingsFileStore.Reset(); var source = Source(); var request = Request();
        True(PivotLayoutStore.SaveNamed(source, "  Sales  ", request));
        request.Columns = new int[0]; True(PivotLayoutStore.SaveNamed(source, "sales", request));
        Equal(1, PivotLayoutStore.GetLayouts(source).Count);
        False(PivotLayoutStore.SaveNamed(source, " ", request));
        False(PivotLayoutStore.SaveNamed(source, new string('x', 81), request));
        True(PivotLayoutStore.DeleteNamed(source, "SALES")); Equal(0, PivotLayoutStore.GetLayouts(source).Count);
    }

    private static void LayoutCorruption()
    {
        foreach (string invalid in new[] { "not JSON", "{}", "{\"Version\":99,\"Schemas\":[]}", "{\"Version\":1}" })
        {
            SettingsFileStore.Reset(); SettingsFileStore.Values["PivotGridLayouts"] = invalid;
            Equal(null, PivotLayoutStore.GetLast(Source())); False(string.IsNullOrWhiteSpace(PivotLayoutStore.LastError));
            False(PivotLayoutStore.SaveLast(Source(), Request())); Equal(invalid, SettingsFileStore.Values["PivotGridLayouts"]);
        }
    }

    private static void LayoutSaveFailure()
    {
        SettingsFileStore.Reset(); SettingsFileStore.FailSave = true;
        False(PivotLayoutStore.SaveLast(Source(), Request())); False(string.IsNullOrWhiteSpace(PivotLayoutStore.LastError));
        Equal(0, SettingsFileStore.Values.Count); SettingsFileStore.Reset();
    }

    private static int[] UnderlyingIds(PivotResult result, int row, int column)
    {
        var details = result.GetUnderlyingRows(row, column, None); var ids = new List<int>();
        for (int page = 0; page < details.PageCount; page++)
            ids.AddRange(details.GetPage(page).Rows.Cast<DataRow>().Select(r => (int)r[0]));
        Equal(details.Count, ids.Count); return ids.ToArray();
    }

    private static void True(bool value) { if (!value) throw new Exception("Expected true."); }
    private static void False(bool value) { if (value) throw new Exception("Expected false."); }
    private static void Equal(object expected, object actual)
    {
        if (expected == null || expected == DBNull.Value) { if (actual == null || actual == DBNull.Value) return; }
        if (expected is decimal && actual != null && actual != DBNull.Value && Convert.ToDecimal(actual, CultureInfo.InvariantCulture) == (decimal)expected) return;
        if (!Equals(expected, actual)) throw new Exception("Expected " + (expected ?? "<null>") + "; actual " + (actual ?? "<null>") + ".");
    }
    private static void Sequence<T>(IEnumerable<T> expected, IEnumerable<T> actual)
    {
        if (!expected.SequenceEqual(actual)) throw new Exception("Sequences differ: expected [" + string.Join(", ", expected) + "]; actual [" + string.Join(", ", actual) + "].");
    }
    private static void Throws<T>(Action action) where T : Exception
    {
        try { action(); } catch (T) { return; }
        throw new Exception("Expected " + typeof(T).Name + ".");
    }
}
