using AxialSqlTools.DataCompare;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

internal static class Program
{
    private static int Main()
    {
        var checks = new Action[]
        {
            DisplayExposesTypeDifferences,
            MappingFeedbackFollowsEdits,
            DuplicateMappingFeedbackClears,
            ReadOnlyAndUnsupportedMappings,
            InvalidMappingsStopBeforeReadingRows,
            CategoriesAndSynchronizationSelection,
            CategorySelectionSpansPages,
            KeyMatchingRemainsExact,
            DuplicateKeysRejectBothSides,
            CancellationProducesNoResult
        };
        try
        {
            foreach (var check in checks)
            {
                check();
                Console.WriteLine("PASS " + check.Method.Name);
            }
            Console.WriteLine(checks.Length + " data comparison regression checks passed.");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 1;
        }
    }

    private static void DisplayExposesTypeDifferences()
    {
        Equal("nvarchar(50)", Column("Name", "nvarchar", 100).TypeDisplay);
        Equal("nvarchar(max)", Column("Name", "nvarchar", -1).TypeDisplay);
        Equal("varchar(100)", Column("Name", "varchar", 100).TypeDisplay);
        Equal("varbinary(max)", Column("Payload", "varbinary", -1).TypeDisplay);
        Equal("decimal(18,4)", Column("Amount", "decimal", 9, 18, 4).TypeDisplay);
        Equal("datetime2(3)", Column("Created", "datetime2", 7, 23, 3).TypeDisplay);
        Equal("float(53)", Column("Value", "float", 8, 53).TypeDisplay);
        var left = Column("Source", "nvarchar", 100);
        var right = Column("Target", "nvarchar", 102);
        Check(!left.HasMatchingType(right), "Different character lengths must remain incompatible.");
        right.Length = 100;
        Check(left.HasMatchingType(right), "Mapping names need not match when types match.");
        var number = Column("Number", "decimal", 9, 18, 4);
        Check(!number.HasMatchingType(Column("Number", "decimal", 9, 17, 4)), "Precision must match.");
        Check(!number.HasMatchingType(Column("Number", "decimal", 9, 18, 3)), "Scale must match.");
    }

    private static void MappingFeedbackFollowsEdits()
    {
        var mapping = new ColumnMapping { Source = Column("Id", "int", 4, 10) };
        var notifications = new HashSet<string>();
        mapping.PropertyChanged += (sender, args) => notifications.Add(args.PropertyName);
        Check(!mapping.HasIssue, "Excluded columns must not block comparison.");
        mapping.Include = true;
        Check(mapping.HasIssue && mapping.Status.Contains("choose target"), "An included unmapped column needs a target.");
        Check(notifications.Contains("Status") && notifications.Contains("HasIssue"), "Feedback must update after Include changes.");
        notifications.Clear();
        mapping.Target = Column("DestinationId", "int", 4, 10);
        Check(!mapping.HasIssue && mapping.Status == "Included", "Valid remapping must clear the issue.");
        Check(notifications.Contains("Status") && notifications.Contains("HasIssue"), "Feedback must update after Target changes.");
        notifications.Clear();
        mapping.IsKey = true;
        Check(mapping.Include && mapping.Status == "Key", "A key must be included.");
        Check(notifications.Contains("Status") && notifications.Contains("HasIssue"), "Feedback must update after key changes.");
        mapping.Include = false;
        Check(!mapping.IsKey && !mapping.HasIssue && mapping.Status == "Excluded", "Excluding a key must clear it.");
        mapping.Target = Column("Id", "bigint", 8, 19);
        Check(!mapping.HasIssue && mapping.Status.Contains("type mismatch"), "Excluded mismatch should explain itself without blocking.");
        mapping.Include = true;
        Check(mapping.HasIssue, "An included mismatch must be marked invalid.");
    }

    private static void ReadOnlyAndUnsupportedMappings()
    {
        var mapping = new ColumnMapping { Source = Column("Value", "int", 4, 10), Target = Column("Value", "int", 4, 10), Include = true };
        mapping.Target.Computed = true;
        Check(!mapping.HasIssue && mapping.Status == "read only", "Read-only non-key comparison remains supported.");
        mapping.IsKey = true;
        Check(mapping.HasIssue && mapping.Status.Contains("Invalid key"), "Read-only columns cannot form a key.");
        mapping.Include = false;
        Check(!mapping.HasIssue, "Excluded read-only columns must not block comparison.");
        mapping.Target.Computed = false;
        mapping.Source.Encrypted = true;
        mapping.Include = true;
        Check(mapping.HasIssue && mapping.Status == "unsupported", "Encrypted columns remain unsupported.");
    }

    private static void DuplicateMappingFeedbackClears()
    {
        var mapping = new ColumnMapping { Source = Column("Source", "int", 4, 10), Target = Column("Target", "int", 4, 10), Include = true };
        var notifications = new List<string>();
        mapping.PropertyChanged += (sender, args) => notifications.Add(args.PropertyName);
        mapping.HasDuplicateTarget = true;
        Check(mapping.HasIssue && mapping.Status == "Target mapped more than once", "Duplicate target mapping must appear in mapping feedback.");
        Check(notifications.Contains("Status") && notifications.Contains("HasIssue"), "Duplicate mapping changes must refresh dependent feedback.");
        notifications.Clear();
        mapping.HasDuplicateTarget = true;
        Equal(0, notifications.Count);
        mapping.Include = false;
        Check(!mapping.HasIssue && mapping.Status == "Excluded", "An excluded duplicate must not block comparison.");
        mapping.Include = true;
        Check(!mapping.Copy().HasDuplicateTarget, "Plan copies must not retain mapping-list context.");
        notifications.Clear();
        mapping.HasDuplicateTarget = false;
        Check(!mapping.HasIssue && mapping.Status == "Included", "Resolving duplicate mapping must clear its issue.");
        Check(notifications.Contains("Status") && notifications.Contains("HasIssue"), "Resolving a duplicate must update feedback.");
    }

    private static void InvalidMappingsStopBeforeReadingRows()
    {
        var plan = Plan();
        plan.Columns[1].Target.Length = 80;
        Throws<InvalidOperationException>(() => ComparisonEngine.Compare(plan, UnexpectedRead(), UnexpectedRead(), CancellationToken.None), "types, lengths");
        plan = Plan();
        plan.Columns[0].IsKey = false;
        Throws<InvalidOperationException>(() => plan.Validate(), "key columns");
        plan = Plan();
        plan.Columns[1].Target = plan.Columns[0].Target;
        Throws<InvalidOperationException>(() => plan.Validate(), "mapped once");
    }

    private static IEnumerable<object[]> UnexpectedRead()
    {
        throw new Exception("Invalid setup must be rejected before reading rows.");
#pragma warning disable CS0162
        yield break;
#pragma warning restore CS0162
    }

    private static void CategoriesAndSynchronizationSelection()
    {
        var source = new[] { Row(3, "insert"), Row(2, "new"), Row(1, "same") };
        var target = new[] { Row(4, "delete"), Row(1, "same"), Row(2, "old") };
        using (var result = ComparisonEngine.Compare(Plan(), source, target, CancellationToken.None))
        {
            foreach (DifferenceKind kind in Enum.GetValues(typeof(DifferenceKind))) Equal(1L, result.Count(kind));
            Equal(2L, result.SelectedCount);
            Check(result.Selection.IsSelected(DifferenceKind.Different, 0), "Updates are selected initially.");
            Check(result.Selection.IsSelected(DifferenceKind.OnlySource, 0), "Inserts are selected initially.");
            Check(!result.Selection.IsSelected(DifferenceKind.OnlyTarget, 0), "Deletes must require selection.");
            var changed = result.Read(DifferenceKind.Different).Single();
            Check(!changed.Changed[0] && changed.Changed[1], "Details must identify only the changed value.");
            var script = Script(result);
            Check(script.Contains("INSERT INTO [dbo].[Target]") && script.Contains("UPDATE [dbo].[Target]"), "Selected inserts and updates must be scripted.");
            Check(!script.Contains("DELETE FROM [dbo].[Target]"), "Unselected deletes must not be scripted.");
            result.Selection.SetAll(DifferenceKind.OnlyTarget, true);
            Equal(3L, result.SelectedCount);
            Check(Script(result).Contains("DELETE FROM [dbo].[Target]"), "Explicitly selected deletes must be scripted.");
            result.Selection.SetAll(DifferenceKind.Identical, true);
            Equal(3L, result.SelectedCount);
            Check(!result.Selection.IsSelected(DifferenceKind.Identical, 0), "Identical rows cannot be synchronized.");
        }
    }

    private static void CategorySelectionSpansPages()
    {
        var source = Enumerable.Range(1, 450).Reverse().Select(i => Row(i, "row " + i));
        using (var result = ComparisonEngine.Compare(Plan(), source, Array.Empty<object[]>(), CancellationToken.None))
        {
            Equal(450L, result.SelectedCount);
            Equal(201, (int)result.Read(DifferenceKind.OnlySource, 200, 200).First().Source[0]);
            result.Selection.Set(DifferenceKind.OnlySource, 210, false);
            Equal(449L, result.SelectedCount);
            Check(!result.Selection.IsSelected(DifferenceKind.OnlySource, 210), "Row selection must survive paging.");
            result.Selection.SetAll(DifferenceKind.OnlySource, false);
            Equal(0L, result.SelectedCount);
            result.Selection.SetAll(DifferenceKind.OnlySource, true);
            Equal(450L, result.SelectedCount);
            Equal(450, result.SelectedRows(CancellationToken.None).Count());
        }
    }

    private static void KeyMatchingRemainsExact()
    {
        var plan = Plan();
        plan.Columns[0].Source = plan.Columns[0].Target = Column("Id", "varchar", 50);
        plan.Options.IgnoreCase = true;
        plan.Options.IgnoreTrailingSpaces = true;
        var source = new[] { new object[] { "case", "same" }, new object[] { "space ", "same" }, new object[] { "exact", "Value " } };
        var target = new[] { new object[] { "CASE", "same" }, new object[] { "space", "same" }, new object[] { "exact", "value" } };
        using (var result = ComparisonEngine.Compare(plan, source, target, CancellationToken.None))
        {
            Equal(2L, result.Count(DifferenceKind.OnlySource));
            Equal(2L, result.Count(DifferenceKind.OnlyTarget));
            Equal(1L, result.Count(DifferenceKind.Identical));
            Equal(0L, result.Count(DifferenceKind.Different));
        }
    }

    private static void DuplicateKeysRejectBothSides()
    {
        var rows = new[] { Row(1, "one"), Row(1, "two") };
        Throws<InvalidOperationException>(() => ComparisonEngine.Compare(Plan(), rows, Array.Empty<object[]>(), CancellationToken.None), "source table");
        Throws<InvalidOperationException>(() => ComparisonEngine.Compare(Plan(), Array.Empty<object[]>(), rows, CancellationToken.None), "target table");
    }

    private static void CancellationProducesNoResult()
    {
        using (var cancellation = new CancellationTokenSource())
        {
            cancellation.Cancel();
            Throws<OperationCanceledException>(() => ComparisonEngine.Compare(Plan(), new[] { Row(1, "one") }, Array.Empty<object[]>(), cancellation.Token));
        }
    }

    private static ComparisonPlan Plan()
    {
        var source = new TableSchema { Server = "Server", Database = "Database", Schema = "dbo", Name = "Source" };
        var target = new TableSchema { Server = "Server", Database = "Database", Schema = "dbo", Name = "Target" };
        source.Columns.AddRange(new[] { Column("Id", "int", 4, 10), Column("Value", "nvarchar", 100) });
        target.Columns.AddRange(new[] { Column("Id", "int", 4, 10), Column("Value", "nvarchar", 100) });
        var key = new TableKey { Name = "PK_Source", Primary = true };
        key.Columns.Add("Id");
        source.Keys.Add(key);
        return new ComparisonPlan { Source = source, Target = target, Columns = ComparisonPlan.AutoMap(source, target) };
    }

    private static TableColumn Column(string name, string type, short length, byte precision = 0, byte scale = 0) =>
        new TableColumn { Name = name, Type = type, Length = length, Precision = precision, Scale = scale };
    private static object[] Row(int id, string value) => new object[] { id, value };
    private static string Script(ComparisonResult result)
    {
        using (var writer = new StringWriter())
        {
            SynchronizationScript.Write(result, writer, CancellationToken.None);
            return writer.ToString();
        }
    }
    private static void Equal<T>(T expected, T actual) => Check(EqualityComparer<T>.Default.Equals(expected, actual), "Expected " + expected + "; received " + actual + ".");
    private static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
    private static void Throws<T>(Action action, string message = null) where T : Exception
    {
        try { action(); }
        catch (T error)
        {
            Check(message == null || error.Message.Contains(message), "Unexpected error: " + error.Message);
            return;
        }
        throw new Exception("Expected " + typeof(T).Name + ".");
    }
}
