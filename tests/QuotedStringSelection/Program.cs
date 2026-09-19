using AxialSqlTools.EditorSelection;
using System;
using System.Collections.Generic;

internal static class Program
{
    private static int _checks;

    private static void Main()
    {
        Check("SELECT 'abc-def-123';", "abc-def-123");
        Check("SELECT 'some text with spaces';", "some text with spaces");
        Check("SELECT N'Customer-Name';", "Customer-Name");
        Check("SELECT n'lowercase-prefix';", "lowercase-prefix");
        Check("SELECT 'O''Brien-Smith';", "O''Brien-Smith");
        Check("SELECT '''';", "''");
        Check("SELECT 'abc''';", "abc''");
        Check("SELECT 'first-line\r\nsecond-line';", "first-line\r\nsecond-line");
        Check("SELECT 'first\nsecond';", "first\nsecond");
        Check("SELECT N'Привет-世界-😀';", "Привет-世界-😀");
        Check("SELECT '--not a comment /* either */';", "--not a comment /* either */");
        Check("SELECT 'left-value', N'right-value';", "left-value", "right-value");
        Check("SELECT 'same-value', 'same-value';", "same-value", "same-value");
        Check("SELECT 'one' + 'two';", "one", "two");
        Check("SELECT 'http://host/a-b?x=1', 'a_b.c/d\\e';", "http://host/a-b?x=1", "a_b.c/d\\e");
        Check("SELECT 'before-error'; /* unfinished comment", "before-error");
        Check("SELECT 'before-unfinished', 'unfinished", "before-unfinished");
        Check("SELECT ''; SELECT N'';");
        Check("SELECT 'unfinished");
        Check("SELECT 'escaped-at-end''");
        Check("SELECT N'escaped-at-end''");
        Check("-- 'not-a-string'\r\nSELECT 'real-value';", "real-value");
        Check("/* 'not-a-string' /* nested 'still-comment' */ */ SELECT 'real-value';", "real-value");
        Check("SELECT [column'with-quotes], [escaped]]'identifier], \"quoted'identifier\";");
        Check("SELECT 1 - 2, @some_variable;");
        Check("SELECT 'x';\nGO\nSELECT N'y';", "x", "y");
        Check("SELECT 'incomplete statement'", "incomplete statement");
        Check("");
        Check(null);

        // Rebuild on a new snapshot: the old resolver must never supply offsets for edited text.
        var oldResolver = SqlStringSelectionResolver.Create("SELECT 'old-value';");
        var newResolver = SqlStringSelectionResolver.Create("-- a new line\nSELECT 'new-value';");
        Assert(oldResolver.TryGetContentSpan(8, out var oldSpan) && oldSpan.Start == 8, "old snapshot span");
        Assert(!newResolver.TryGetContentSpan(8, out _), "edited snapshot ignores old offsets");
        Assert(newResolver.TryGetContentSpan(22, out var newSpan) && newSpan.Start == 22, "edited snapshot uses new offsets");
        Console.WriteLine($"Passed {_checks} selection assertions.");
    }

    // Check every source position, including quotes, N prefixes, whitespace, and EOF.
    private static void Check(string sql, params string[] expectedContents)
    {
        var resolver = SqlStringSelectionResolver.Create(sql);
        sql = sql ?? string.Empty;
        var expected = new List<(int Start, int Length)>();
        int searchStart = 0;
        foreach (string content in expectedContents)
        {
            int opening = sql.IndexOf("'" + content + "'", searchStart, StringComparison.Ordinal);
            Assert(opening >= 0, "test fixture contains expected content");
            int start = opening + 1;
            expected.Add((start, content.Length));
            searchStart = start + content.Length;
        }

        for (int position = -1; position <= sql.Length; position++)
        {
            var match = expected.Find(s => position >= s.Start && position < s.Start + s.Length);
            bool selected = resolver.TryGetContentSpan(position, out var actual);
            Assert(selected == (match.Length > 0), $"selection at {position} in {sql}");
            if (selected)
                Assert(actual.Start == match.Start && actual.Length == match.Length,
                    $"complete literal contents at {position} in {sql}");
        }
    }

    private static void Assert(bool condition, string message)
    {
        _checks++;
        if (!condition) throw new InvalidOperationException(message);
    }
}
