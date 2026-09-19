using AxialSqlTools.DataCompare;
using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

internal static class Program
{
    private static int passed;
    private static string scriptOutput;
    private static int Main(string[] args)
    {
        if (args.Length == 1) { scriptOutput = args[0]; Directory.CreateDirectory(scriptOutput); }
        Run("all four categories and per-column differences", Categories);
        Run("empty source, target and both", EmptyTables);
        Run("duplicate source key rejects entire result", DuplicateSource);
        Run("duplicate target key rejects entire result", DuplicateTarget);
        Run("composite keys with NULL and delimiter characters", CompositeKeys);
        Run("case and spaces remain exact in keys", ExactKeys);
        Run("case and trailing space options affect values only", TextOptions);
        Run("external sort multiple merge passes", MultipleRuns);
        Run("large paged results preserve group ordering", Paging);
        Run("NULL differs from empty text and empty binary", Nulls);
        Run("binary equality compares all bytes", BinaryEquality);
        Run("SQL decimal preserves full 38-digit precision", DecimalPrecision);
        Run("all supported CLR values round-trip", ValueRoundTrip);
        Run("UTF-16 surrogate code units round-trip", Surrogates);
        Run("date/time ticks and offsets round-trip", DateTimeRoundTrip);
        Run("date offset equality uses the same instant", OffsetEquality);
        Run("duplicate key is detected across sorted runs", DuplicateAcrossRuns);
        Run("cancellation during input returns no results", Cancellation);
        Run("cancellation during merge returns no results", MergeCancellation);
        Run("storage budget stops comparison and cleans temporary files", StorageLimit);
        Run("same table comparison is rejected", SameTable);
        Run("large aggregate row is rejected", LargeRow);
        Run("cell limit stops comparison and cleans temporary files", CellLimit);
        Run("completed and disposed results clean temporary files", DisposeCleanup);
        Run("automatic mapping uses complete composite primary key", AutoMapping);
        Run("case-sensitive column names and ambiguous mappings", AmbiguousMapping);
        Run("missing keys and duplicate mappings are rejected", InvalidMappings);
        Run("incompatible type lengths and scales are rejected", IncompatibleTypes);
        Run("unsupported types and encrypted keys are rejected", UnsupportedTypes);
        Run("selection defaults exclude deletes and identical rows", SelectionDefaults);
        Run("category selection and row exceptions", CategorySelection);
        Run("SQL script is transactional and destination guarded", ScriptTransaction);
        Run("UPDATE protects original values and exact text", UpdateGuards);
        Run("source-only rows generate guarded INSERTs", InsertGuards);
        Run("deletes require explicit selection", DeleteSelection);
        Run("identity INSERTs enable and reset identity insertion", IdentityInsert);
        Run("required unmapped target columns reject inserts", RequiredTargetColumns);
        Run("computed-only differences cannot be applied", ReadOnlyDifference);
        Run("identity updates cannot be applied", IdentityDifference);
        Run("temporal history target cannot be synchronized", HistoryTarget);
        Run("quoted identifiers cannot inject SQL", Identifiers);
        Run("Unicode values are hex encoded in scripts", StringLiterals);
        Run("SQL literals are culture invariant", CultureIndependent);
        Run("datetime precision and XML style are explicit", TypedLiterals);
        Run("ignored text differences are preserved during updates", IgnoredPostcondition);
        Run("CSV preserves null metadata and escapes formulas", CsvExport);
        Run("script cancellation produces no complete script", ScriptCancellation);
        Console.WriteLine(passed + " regression tests passed.");
        return 0;
    }

    private static void Run(string name, Action test)
    {
        try { test(); passed++; Console.WriteLine("PASS " + name); }
        catch (Exception ex) { Console.Error.WriteLine("FAIL " + name + ": " + ex); Environment.Exit(1); }
    }
    private static void Check(bool condition, string message = "Assertion failed") { if (!condition) throw new Exception(message); }
    private static void Reject<T>(Action action) where T : Exception
    {
        try { action(); } catch (T) { return; }
        throw new Exception("Expected " + typeof(T).Name);
    }
    private static TableColumn Column(string name, string type = "nvarchar") => new TableColumn { Name = name, Type = type, Length = type == "int" ? (short)4 : (short)100, Nullable = true };
    private static ComparisonPlan Plan()
    {
        var source = new TableSchema { Server = "source", Database = "SourceDb", Schema = "dbo", Name = "Source" };
        var target = new TableSchema { Server = "target", Database = "TargetDb", Schema = "dbo", Name = "Target" };
        source.Columns.AddRange(new[] { Column("Id", "int"), Column("Value") });
        target.Columns.AddRange(new[] { Column("Id", "int"), Column("Value") });
        source.Keys.Add(new TableKey { Name = "PK_Source", Primary = true }); source.Keys[0].Columns.Add("Id");
        return new ComparisonPlan { Source = source, Target = target, Columns = ComparisonPlan.AutoMap(source, target),
            Options = new ComparisonOptions { SortMemoryBytes = 2048 } };
    }
    private static object[] Row(int key, object value) => new object[] { key, value };
    private static ComparisonResult Compare(ComparisonPlan plan, IEnumerable<object[]> source, IEnumerable<object[]> target,
        CancellationToken token = default(CancellationToken), IProgress<ComparisonProgress> progress = null) => ComparisonEngine.Compare(plan, source, target, token, progress);
    private static string Script(ComparisonResult result)
    {
        using (var writer = new StringWriter(CultureInfo.InvariantCulture))
        {
            SynchronizationScript.Write(result, writer, CancellationToken.None);
            string sql = writer.ToString();
            if (scriptOutput != null) File.WriteAllText(Path.Combine(scriptOutput, "generated-" + passed + ".sql"), sql);
            return sql;
        }
    }
    private static void Categories()
    {
        using (var r = Compare(Plan(), new[] { Row(4, "same"), Row(1, "new"), Row(2, "source") }, new[] { Row(3, "target"), Row(1, "old"), Row(4, "same") }))
        {
            foreach (DifferenceKind k in Enum.GetValues(typeof(DifferenceKind))) Check(r.Count(k) == 1);
            var different = r.Read(DifferenceKind.Different).Single();
            Check(!different.Changed[0] && different.Changed[1]);
            Check((string)different.Source[1] == "new" && (string)different.Target[1] == "old");
        }
    }
    private static void EmptyTables()
    {
        using (var r = Compare(Plan(), new object[0][], new object[0][])) Check(r.SelectedCount == 0);
        using (var r = Compare(Plan(), new[] { Row(1, "a") }, new object[0][])) Check(r.Count(DifferenceKind.OnlySource) == 1);
        using (var r = Compare(Plan(), new object[0][], new[] { Row(1, "a") })) Check(r.Count(DifferenceKind.OnlyTarget) == 1);
    }
    private static void DuplicateSource() => Reject<InvalidOperationException>(() => Compare(Plan(), new[] { Row(1, "a"), Row(1, "b") }, new object[0][]));
    private static void DuplicateTarget() => Reject<InvalidOperationException>(() => Compare(Plan(), new object[0][], new[] { Row(1, "a"), Row(1, "b") }));
    private static void CompositeKeys()
    {
        var p = Plan(); p.Columns[1].IsKey = true;
        var rows = new[] { Row(1, "a|b"), Row(1, null), Row(1, ""), Row(2, "a|b") };
        using (var r = Compare(p, rows, rows.Reverse())) Check(r.Count(DifferenceKind.Identical) == 4);
    }
    private static void ExactKeys()
    {
        var p = Plan(); p.Columns[0].IsKey = false; p.Columns[1].IsKey = true;
        p.Options.IgnoreCase = p.Options.IgnoreTrailingSpaces = true;
        using (var r = Compare(p, new[] { Row(1, "key"), Row(2, "key ") }, new[] { Row(1, "KEY"), Row(2, "key ") }))
        { Check(r.Count(DifferenceKind.OnlySource) == 1 && r.Count(DifferenceKind.OnlyTarget) == 1 && r.Count(DifferenceKind.Identical) == 1); }
    }
    private static void TextOptions()
    {
        var p = Plan();
        using (var r = Compare(p, new[] { Row(1, "Hello ") }, new[] { Row(1, "HELLO") })) Check(r.Count(DifferenceKind.Different) == 1);
        p.Options.IgnoreCase = p.Options.IgnoreTrailingSpaces = true;
        using (var r = Compare(p, new[] { Row(1, "Hello ") }, new[] { Row(1, "HELLO") })) Check(r.Count(DifferenceKind.Identical) == 1);
        using (var r = Compare(p, new[] { Row(1, "Hello\t") }, new[] { Row(1, "HELLO") })) Check(r.Count(DifferenceKind.Different) == 1);
    }
    private static void MultipleRuns()
    {
        var random = new Random(72);
        var input = Enumerable.Range(0, 5000).OrderBy(_ => random.Next()).Select(i => Row(i, "v" + i)).ToArray();
        using (var r = Compare(Plan(), input, input.Reverse()))
        {
            Check(r.Count(DifferenceKind.Identical) == input.Length);
            Check(r.Read(DifferenceKind.Identical).Select(x => (int)x.Source[0]).SequenceEqual(Enumerable.Range(0, input.Length)));
        }
    }
    private static void Paging()
    {
        using (var r = Compare(Plan(), Enumerable.Range(0, 555).Select(i => Row(i, "s")), Enumerable.Range(0, 555).Select(i => Row(i, "t"))))
        {
            var page = r.Read(DifferenceKind.Different, 200, 200).ToList();
            Check(page.Count == 200 && page[0].Index == 200 && (int)page[199].Source[0] == 399);
            Check(r.Read(DifferenceKind.Different, 500, 200).Count() == 55);
            Check(!r.Read(DifferenceKind.Different, 600, 200).Any());
        }
    }
    private static void Nulls()
    {
        Check(ValueCodec.Compare(null, DBNull.Value) == 0);
        Check(ValueCodec.Compare(null, "") != 0 && ValueCodec.Compare(null, new byte[0]) != 0);
        using (var r = Compare(Plan(), new[] { Row(1, null), Row(2, "") }, new[] { Row(1, ""), Row(2, null) })) Check(r.Count(DifferenceKind.Different) == 2);
    }
    private static void BinaryEquality()
    {
        Check(ValueCodec.Compare(new byte[] { 1, 0 }, new byte[] { 1 }) != 0);
        Check(ValueCodec.Compare(new byte[] { 1, 2 }, new byte[] { 1, 2 }) == 0);
        using (var r = Compare(Plan(), new[] { Row(1, new byte[] { 0, 255 }) }, new[] { Row(1, new byte[] { 0, 254 }) })) Check(r.Count(DifferenceKind.Different) == 1);
    }
    private static void DecimalPrecision()
    {
        var big = SqlDecimal.Parse("99999999999999999999999999999999999999");
        var less = SqlDecimal.Parse("99999999999999999999999999999999999998");
        using (var r = Compare(Plan(), new[] { Row(1, big) }, new[] { Row(1, less) }))
        {
            Check(r.Count(DifferenceKind.Different) == 1);
            Check(((SqlDecimal)r.Read(DifferenceKind.Different).Single().Source[1]).ToString() == big.ToString());
            Check(SqlText.Literal(big) == "99999999999999999999999999999999999999");
        }
    }
    private static object[] RoundTrip(object[] values)
    {
        using (var stream = new MemoryStream())
        {
            using (var writer = new BinaryWriter(stream, System.Text.Encoding.UTF8, true)) ValueCodec.WriteRow(writer, values);
            stream.Position = 0;
            using (var reader = new BinaryReader(stream)) return ValueCodec.ReadRow(reader);
        }
    }
    private static void ValueRoundTrip()
    {
        var values = new object[] { null, DBNull.Value, "", "a\0z", new byte[] { 0, 255 }, long.MinValue, int.MaxValue, short.MinValue, byte.MaxValue,
            true, 123.450m, SqlDecimal.Parse("-0.12345678901234567890123456789012345678"), 1.234567890123456e-120, 1.2345f,
            new DateTime(2026, 1, 1).AddTicks(1234567), new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.FromHours(5.5)), TimeSpan.FromTicks(123456789), Guid.NewGuid() };
        var copy = RoundTrip(values);
        for (int i = 0; i < values.Length; i++) Check(ValueCodec.Compare(values[i], copy[i]) == 0, "Round-trip column " + i);
    }
    private static void Surrogates() { var value = "\ud800x\udfff"; Check((string)RoundTrip(new object[] { value })[0] == value); }
    private static void DateTimeRoundTrip()
    {
        var date = new DateTime(2026, 8, 12, 0, 0, 0).AddTicks(1234567);
        var offset = new DateTimeOffset(date, TimeSpan.FromHours(-7));
        var row = RoundTrip(new object[] { date, offset, TimeSpan.FromTicks(99) });
        Check(((DateTime)row[0]).Ticks == date.Ticks && ((DateTimeOffset)row[1]).Offset == offset.Offset && ((TimeSpan)row[2]).Ticks == 99);
    }
    private static void OffsetEquality()
    {
        var now = new DateTimeOffset(2026, 1, 1, 12, 0, 0, TimeSpan.Zero);
        Check(ValueCodec.Compare(now, now.ToOffset(TimeSpan.FromHours(-7))) == 0);
    }
    private static void DuplicateAcrossRuns() => Reject<InvalidOperationException>(() => Compare(Plan(), Enumerable.Range(0, 1000).Select(i => Row(i, "x")).Concat(new[] { Row(1, "duplicate") }), new object[0][]));
    private static void Cancellation()
    {
        using (var c = new CancellationTokenSource())
        {
            var input = Enumerable.Range(0, 200).Select(i => { if (i == 50) c.Cancel(); return Row(i, "x"); });
            Reject<OperationCanceledException>(() => Compare(Plan(), input, new object[0][], c.Token));
        }
    }
    private sealed class ImmediateProgress : IProgress<ComparisonProgress>
    {
        public Action<ComparisonProgress> ReportAction;
        public void Report(ComparisonProgress value) => ReportAction(value);
    }
    private static void MergeCancellation()
    {
        using (var c = new CancellationTokenSource())
        {
            var progress = new ImmediateProgress { ReportAction = p => { if (p.Phase == "Comparing") c.Cancel(); } };
            Reject<OperationCanceledException>(() => Compare(Plan(), Enumerable.Range(0, 4000).Select(i => Row(i, "x")), new object[0][], c.Token, progress));
        }
    }
    private static string TempRoot => Path.Combine(Path.GetTempPath(), "AxialSQL", "DataCompare");
    private static int TempDirectories() => Directory.Exists(TempRoot) ? Directory.GetDirectories(TempRoot).Length : 0;
    private static void StorageLimit()
    {
        int before = TempDirectories(); var p = Plan(); p.Options.MaxTemporaryBytes = 1024;
        Reject<IOException>(() => Compare(p, Enumerable.Range(0, 100).Select(i => Row(i, new string('x', 100))), new object[0][]));
        Check(TempDirectories() == before);
    }
    private static void SameTable()
    {
        var p = Plan(); p.Target.Server = p.Source.Server; p.Target.Database = p.Source.Database; p.Target.Name = p.Source.Name;
        Reject<InvalidOperationException>(p.Validate);
    }
    private static void LargeRow()
    {
        var p = Plan(); p.Options.MaxCellBytes = int.MaxValue;
        Reject<InvalidDataException>(() => Compare(p, new[] { Row(1, new string('x', 17 * 1024 * 1024)) }, new object[0][]));
    }
    private static void CellLimit()
    {
        int before = TempDirectories(); var p = Plan(); p.Options.MaxCellBytes = 10;
        Reject<InvalidDataException>(() => Compare(p, new[] { Row(1, "123456") }, new object[0][])); Check(TempDirectories() == before);
    }
    private static void DisposeCleanup()
    {
        int before = TempDirectories(); var r = Compare(Plan(), new[] { Row(1, "x") }, new object[0][]);
        Check(TempDirectories() == before + 1); r.Dispose(); Check(TempDirectories() == before);
        Reject<InvalidOperationException>(() => r.Read(DifferenceKind.OnlySource).ToList());
    }
    private static void AutoMapping()
    {
        var p = Plan(); p.Source.Keys[0].Columns.Add("Value");
        var mappings = ComparisonPlan.AutoMap(p.Source, p.Target); Check(mappings.All(m => m.IsKey));
        p.Target.Columns[1].Computed = true;
        Check(!ComparisonPlan.AutoMap(p.Source, p.Target)[1].Include);
    }
    private static void AmbiguousMapping()
    {
        var p = Plan(); p.Source.Columns[1].Name = "VALUE"; p.Target.Columns.Add(Column("value"));
        Check(ComparisonPlan.AutoMap(p.Source, p.Target)[1].Target == null);
        p.Source.Columns[1].Name = "Value"; Check(ComparisonPlan.AutoMap(p.Source, p.Target)[1].Target.Name == "Value");
    }
    private static void InvalidMappings()
    {
        var p = Plan(); p.Columns[0].IsKey = false; Reject<InvalidOperationException>(p.Validate);
        p = Plan(); p.Columns[1].Target = p.Columns[0].Target; Reject<InvalidOperationException>(p.Validate);
        p = Plan(); p.Columns[1].Target = null; Reject<InvalidOperationException>(p.Validate);
    }
    private static void IncompatibleTypes()
    {
        var p = Plan(); p.Target.Columns[1].Length = 200; Reject<InvalidOperationException>(p.Validate);
        p = Plan(); p.Target.Columns[1].Scale = 3; Reject<InvalidOperationException>(p.Validate);
    }
    private static void UnsupportedTypes()
    {
        var p = Plan(); p.Source.Columns[0].Encrypted = true; Reject<InvalidOperationException>(p.Validate);
        p = Plan(); p.Source.Columns[1].Type = "geography"; Reject<InvalidOperationException>(p.Validate);
    }
    private static void SelectionDefaults()
    {
        using (var r = Compare(Plan(), new[] { Row(1, "x"), Row(2, "s"), Row(4, "same") }, new[] { Row(2, "t"), Row(3, "delete"), Row(4, "same") }))
        { Check(r.SelectedCount == 2); Check(!r.Selection.IsSelected(DifferenceKind.OnlyTarget, 0)); Check(!r.Selection.IsSelected(DifferenceKind.Identical, 0)); }
    }
    private static void CategorySelection()
    {
        var s = new ResultSelection(); s.SetAll(DifferenceKind.Different, false); s.Set(DifferenceKind.Different, 5, true);
        Check(s.Count(DifferenceKind.Different, 10) == 1 && s.IsSelected(DifferenceKind.Different, 5));
        s.SetAll(DifferenceKind.Different, true); s.Set(DifferenceKind.Different, 2, false); Check(s.Count(DifferenceKind.Different, 10) == 9);
        s.SetAll(DifferenceKind.Identical, true); Check(!s.IsSelected(DifferenceKind.Identical, 0));
    }
    private static void ScriptTransaction()
    {
        using (var r = Compare(Plan(), new[] { Row(1, "new") }, new[] { Row(1, "old") }))
        {
            string sql = Script(r);
            Check(sql.Contains("BEGIN TRANSACTION") && sql.Contains("ROLLBACK TRANSACTION") && sql.Contains("SET XACT_ABORT ON"));
            Check(sql.Contains("SERVERPROPERTY('ServerName')") && sql.Contains("N'TargetDb'") && sql.Contains("Target column metadata changed"));
            Check(!sql.Contains("DISABLE TRIGGER") && !sql.Contains("NOCHECK"));
        }
    }
    private static void UpdateGuards()
    {
        using (var r = Compare(Plan(), new[] { Row(1, "new") }, new[] { Row(1, "OLD ") }))
        {
            string sql = Script(r); Check(sql.Contains("UPDLOCK, HOLDLOCK") && sql.Contains("COUNT_BIG(*)"));
            Check(sql.Contains("Target data changed") && sql.Contains("DATALENGTH") && sql.Contains("4F004C0044002000"));
            Check(sql.Contains("UPDATE [dbo].[Target] SET [Value]") && !sql.Contains("SET [Id]"));
        }
    }
    private static void InsertGuards()
    {
        using (var r = Compare(Plan(), new[] { Row(1, "new") }, new object[0][]))
        {
            string sql = Script(r); Check(sql.Contains("IF EXISTS") && sql.Contains("INSERT INTO [dbo].[Target] ([Id], [Value])"));
            Check(sql.Contains("Target did not preserve"));
        }
    }
    private static void DeleteSelection()
    {
        using (var r = Compare(Plan(), new object[0][], new[] { Row(1, "delete") }))
        {
            Reject<InvalidOperationException>(() => Script(r)); r.Selection.SetAll(DifferenceKind.OnlyTarget, true);
            string sql = Script(r); Check(sql.Contains("DELETE FROM") && sql.Contains("Target row remains"));
        }
    }
    private static void IdentityInsert()
    {
        var p = Plan(); p.Target.Columns[0].Identity = true;
        using (var r = Compare(p, new[] { Row(1, "new") }, new object[0][]))
        { string sql = Script(r); Check(sql.Contains("IDENTITY_INSERT [dbo].[Target] ON") && sql.Contains("IDENTITY_INSERT [dbo].[Target] OFF")); }
    }
    private static void RequiredTargetColumns()
    {
        var p = Plan(); p.Target.Columns.Add(new TableColumn { Name = "Required", Type = "int", Nullable = false });
        using (var r = Compare(p, new[] { Row(1, "x") }, new object[0][])) Reject<InvalidOperationException>(() => Script(r));
        p.Target.Columns[2].HasDefault = true;
        using (var r = Compare(p, new[] { Row(1, "x") }, new object[0][])) Check(Script(r).Contains("INSERT INTO"));
    }
    private static void ReadOnlyDifference()
    {
        var p = Plan(); p.Target.Columns[1].Computed = true;
        using (var r = Compare(p, new[] { Row(1, "x") }, new[] { Row(1, "y") })) Reject<InvalidOperationException>(() => Script(r));
    }
    private static void IdentityDifference()
    {
        var p = Plan(); p.Target.Columns[1].Identity = true;
        using (var r = Compare(p, new[] { Row(1, "x") }, new[] { Row(1, "y") })) Reject<InvalidOperationException>(() => Script(r));
    }
    private static void HistoryTarget()
    {
        var p = Plan(); p.Target.HistoryTable = true;
        using (var r = Compare(p, new[] { Row(1, "x") }, new object[0][])) Reject<InvalidOperationException>(() => Script(r));
    }
    private static void Identifiers() { Check(SqlText.Quote("a]; DROP TABLE x;--") == "[a]]; DROP TABLE x;--]"); Check(SqlText.String("o'clock") == "N'o''clock'"); }
    private static void StringLiterals()
    {
        string literal = SqlText.Literal("'; DROP TABLE x;--\nGO\n🙂");
        Check(literal.StartsWith("CONVERT(nvarchar(max), 0x") && !literal.Contains("DROP TABLE"));
        Check(SqlText.Literal("\ud800") == "CONVERT(nvarchar(max), 0x00D8)");
    }
    private static void CultureIndependent()
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            Check(SqlText.Literal(123.45m) == "123.45"); Check(SqlText.Literal(SqlDecimal.Parse("123.45")) == "123.45");
            Check(SqlText.Literal(1.23456789e-70).Contains("1.23456789"));
        }
        finally { CultureInfo.CurrentCulture = original; }
    }
    private static void TypedLiterals()
    {
        Check(SqlText.Literal(new DateTime(2026, 1, 1).AddMilliseconds(3), Column("d", "datetime")).Contains(".003', 126)"));
        var c = Column("d", "datetime2"); c.Scale = 7;
        Check(SqlText.Literal(new DateTime(2026, 1, 1).AddTicks(123), c).Contains("datetime2(7)"));
        Check(SqlText.Literal("<x> </x>", Column("xml", "xml")).EndsWith(", 1)"));
    }
    private static void IgnoredPostcondition()
    {
        var p = Plan(); p.Options.IgnoreCase = true;
        var sc = Column("Other"); var tc = Column("Other"); p.Source.Columns.Add(sc); p.Target.Columns.Add(tc);
        p.Columns.Add(new ColumnMapping { Source = sc, Target = tc, Include = true });
        using (var r = Compare(p, new[] { new object[] { 1, "hello", "new" } }, new[] { new object[] { 1, "HELLO", "old" } }))
        {
            string sql = Script(r); Check(sql.Contains("SET [Other]") && !sql.Contains("SET [Value]"));
            Check(sql.Contains("480045004C004C004F00"));
        }
    }
    private static void CsvExport()
    {
        using (var r = Compare(Plan(), new[] { Row(1, "=1+1"), Row(2, null), Row(3, "a,\"b\"\nc") }, new object[0][]))
        using (var writer = new StringWriter())
        {
            ComparisonCsv.Write(r, writer, CancellationToken.None);
            string csv = writer.ToString(); Check(csv.Contains("SourceIsNull") && csv.Contains("'=1+1") && csv.Contains("a,\"\"b\"\"\nc"));
        }
    }
    private static void ScriptCancellation()
    {
        using (var r = Compare(Plan(), new[] { Row(1, "x") }, new object[0][]))
        using (var c = new CancellationTokenSource())
        using (var writer = new StringWriter())
        {
            c.Cancel(); Reject<OperationCanceledException>(() => SynchronizationScript.Write(r, writer, c.Token));
            Check(!writer.ToString().Contains("COMMIT TRANSACTION"));
        }
    }
}
