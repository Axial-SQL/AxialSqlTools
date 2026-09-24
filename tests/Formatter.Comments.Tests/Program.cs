using AxialSqlTools;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using System.Text.RegularExpressions;

int passed = 0, failed = 0;
void Check(string name, Action action)
{
    try { action(); passed++; }
    catch (Exception ex) { failed++; Console.WriteLine("FAIL " + name + ": " + ex.Message); }
}
void Require(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}
TSqlFragment Parse(string sql)
{
    var fragment = new TSql170Parser(true).Parse(new StringReader(sql), out var errors);
    Require(errors.Count == 0, string.Join("; ", errors.Select(e => e.Message)) + "\n" + sql);
    return fragment;
}
bool IsComment(TSqlParserToken t) => t.TokenType == TSqlTokenType.SingleLineComment || t.TokenType == TSqlTokenType.MultilineComment;
string[] Comments(string sql) => Parse(sql).ScriptTokenStream.Where(IsComment).Select(t => t.Text).ToArray();
string Canonical(string sql)
{
    var generator = new Sql170ScriptGenerator();
    generator.GenerateScript(Parse(sql), out var result);
    return result;
}
SettingsManager.TSqlCodeFormatSettings Options(bool legacy, bool all, bool preserve = true)
{
    var settings = new SettingsManager.TSqlCodeFormatSettings();
    if (all)
        foreach (var field in typeof(SettingsManager.TSqlCodeFormatSettings).GetFields()) field.SetValue(settings, true);
    settings.disregardSsmsFormatterSettings = legacy;
    settings.preserveComments = preserve;
    return settings;
}
string Format(string sql, bool legacy, bool all = false, bool preserve = true, bool native = false)
{
    var generator = new Sql170ScriptGenerator();
    generator.Options.GetType().GetProperty("PreserveComments")?.SetValue(generator.Options, native);
    return TSqlFormatter.FormatCode(sql, Options(legacy, all, preserve), new TSql170Parser(true), generator);
}
string screenshot = @"while (1=0)
begin -- inline comment
select distinct top 10
    c.CustomerID, getDate(),
    CASE WHEN o.TotalAmount > 1000 THEN 'High' ELSE 'Low' END AS OrderSize
FROM Customers c
JOIN Orders o ON c.CustomerID = o.CustomerID CROSS JOIN Regions r
WHERE c.IsActive = 1; /* multi-line
comment */
SELECT dbo.func(p.ProductID), p.ProductName FROM Products p; EXEC dbo.test @a = 0, @b = 1;
end
if 1=0 begin select 1; declare @a int, @b varchar(10) = ''
end
go
create procedure dbo.test @a int, @b int = 0
as select 1;
";
var samples = new Dictionary<string, string>
{
    ["screenshot"] = screenshot,
    ["join line comment"] = "SELECT 1 FROM a JOIN -- join comment\nb ON a.id=b.id;",
    ["join block comment"] = "SELECT 1 FROM a JOIN /* join\ncomment */ b ON a.id=b.id;",
    ["inline BEGIN"] = "WHILE 1=0 BEGIN -- header\nSELECT 1; END;",
    ["trailing EOF"] = "SELECT 1; -- trailing EOF",
    ["trailing EOF without semicolon"] = "SELECT 1 -- trailing EOF",
    ["before semicolon"] = "SELECT 1 -- before terminator\n; SELECT 2;",
    ["standalone leading"] = "-- header\nSELECT 1;",
    ["leading blank lines"] = "\n\n-- header\nSELECT 1;",
    ["comment only"] = "-- one\n/* two\n lines */\n-- three",
    ["block only"] = "/* one */",
    ["consecutive comments"] = "SELECT 1; -- trailing\n-- leading\n/* block */\nSELECT 2;",
    ["adjacent blocks"] = "SELECT/*a*//*b*/1;",
    ["nested block"] = "SELECT /* outer /* inner */ outer */ 1;",
    ["strings with markers"] = "SELECT '--literal' AS [/*identifier*/], N'/*literal*/'; -- actual",
    ["select inline blocks"] = "SELECT /* first */ 1, /* second */ 2;",
    ["select standalone list"] = "SELECT TOP 10\n-- field\n a,\n/* field2 */\n b FROM t;",
    ["select trailing list"] = "SELECT DISTINCT a, -- first\nb, -- second\nc FROM t;",
    ["comments around delimiter"] = "SELECT 1 -- first\n; /* second */ SELECT 2;",
    ["block and line sequence"] = "SELECT 1; /* first */ -- second\nSELECT 2;",
    ["before comma"] = "SELECT 1 -- first\n, 2;",
    ["standalone before END"] = "BEGIN SELECT 1;\n-- end marker\nEND; SELECT 2;",
    ["batch comments"] = "SELECT 1; -- first\nGO -- batch\n-- second\nSELECT 2;",
    ["blank line stability"] = "SELECT 1;\n\n-- explanation\n\nSELECT 2;",
    ["case comments"] = "SELECT CASE -- case\nWHEN 1=1 -- condition\nTHEN /*yes*/ 2 ELSE -- no\n3 END;",
    ["exec comments"] = "EXEC -- procedure\ndbo.p @a=1, -- a\n@b=2;",
    ["declare comments"] = "DECLARE @a int=1, -- a\n@b varchar(10)='x';",
    ["procedure comments"] = "CREATE PROC dbo.p @a int, -- a\n@b int AS BEGIN -- begin\nSELECT @a + @b; END;",
    ["optional AS and INNER"] = "SELECT a.id /* alias */ x FROM a /* table */ aa JOIN b bb ON aa.id=bb.id;",
    ["repeated statements"] = "SELECT 1; -- first\nSELECT 1; -- second\nSELECT 1; -- third",
    ["CRLF"] = "BEGIN -- begin\r\nSELECT 1; /* end */\r\nSELECT 2;\r\nEND;",
    ["CR only"] = "SELECT 1; -- first\rSELECT 2; -- second",
    ["unicode"] = "SELECT N'привет'; -- café привет 😀",
    ["no comments"] = "SELECT 1; SELECT 'text';"
};
foreach (bool legacy in new[] { true, false })
foreach (bool all in new[] { false, true })
foreach (var sample in samples)
{
    string label = $"{sample.Key} legacy={legacy} all={all}";
    Check(label, () =>
    {
        string output = Format(sample.Value, legacy, all);
        Require(Comments(sample.Value).SequenceEqual(Comments(output)), "Comment text/order/count changed\n" + output);
        // Compare to formatting with comments disabled so Axial's deliberate casing changes
        // are allowed, while lost statements, changed literals, and hidden SQL are detected.
        string expected = Format(sample.Value, legacy, all, preserve: false);
        Require(Canonical(output) == Canonical(expected), "SQL meaning changed\n" + output);
        string twice = Format(output, legacy, all);
        Require(output == twice, "Formatting is not stable on a second pass\nFIRST:\n" + output + "\nSECOND:\n" + twice);
    });
}
foreach (bool legacy in new[] { true, false })
{
    Check("Screenshot placement legacy=" + legacy, () =>
    {
        string output = Format(screenshot, legacy, all: true);
        Require(output.Contains("BEGIN -- inline comment"), "Inline comment moved away from BEGIN");
        Require(Regex.IsMatch(output, @"BEGIN -- inline comment\r?\n {4}SELECT DISTINCT TOP 10"), "SELECT lost body indentation\n" + output);
        Require(output.Contains("; /* multi-line\ncomment */"), "Trailing block comment moved to next statement");
        Require(Regex.IsMatch(output, @"comment \*/(?:\r?\n)+ {4}SELECT"), "Statement after block comment lost indentation");
    });
    foreach(bool native in new[] {false, true})
    foreach(bool preserve in new[] {false, true})
        if (!native || typeof(SqlScriptGeneratorOptions).GetProperty("PreserveComments") != null)
        Check($"Preference legacy={legacy} native={native} override={preserve}", () =>
        {
            string output = Format("SELECT 1; -- comment", legacy, preserve: preserve, native: native);
            Require(Comments(output).Length == (preserve || (!legacy && native) ? 1 : 0), "Wrong preference behavior");
        });
}
Check("No generator option leakage", () =>
{
    var generator = new Sql170ScriptGenerator();
    var option = generator.Options.GetType().GetProperty("PreserveComments");
    option?.SetValue(generator.Options, false);
    TSqlFormatter.FormatCode("SELECT 1; -- first", Options(false, false), new TSql170Parser(true), generator);
    Require(!(option?.GetValue(generator.Options) is bool value && value), "Axial override mutated the shared generator");
    string next = TSqlFormatter.FormatCode("SELECT 2; -- second", Options(false, false, false), new TSql170Parser(true), generator);
    Require(Comments(next).Length == 0, "Earlier override leaked to next format");
});
Check("Direct helper does not duplicate native comments", () =>
{
    var generator = new Sql170ScriptGenerator();
    var option = generator.Options.GetType().GetProperty("PreserveComments");
    option?.SetValue(generator.Options, true);
    string sql = "SELECT 1; -- only once";
    string output = TsqlFormatterCommentInterleaver.GenerateWithComments(Parse(sql), generator, new TSql170Parser(true));
    Require(Comments(output).SequenceEqual(Comments(sql)), "Duplicated or changed comment");
    Require(option == null || (bool)option.GetValue(generator.Options), "Native option was not restored");
});
Check("Alignment matches LCS reference on repeated/random tokens", () =>
{
    var method = typeof(TsqlFormatterCommentInterleaver).GetMethod("LongestCommonSubsequence",
        System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static).MakeGenericMethod(typeof(string));
    var random = new Random(2371);
    for (int sample = 0; sample < 1000; sample++)
    {
        var a = Enumerable.Range(0, random.Next(0, 30)).Select(_ => random.Next(0, 5).ToString()).ToArray();
        var b = Enumerable.Range(0, random.Next(0, 30)).Select(_ => random.Next(0, 5).ToString()).ToArray();
        var expected = new int[a.Length + 1, b.Length + 1];
        for (int i = a.Length - 1; i >= 0; i--)
        for (int j = b.Length - 1; j >= 0; j--)
            expected[i,j] = a[i] == b[j] ? 1 + expected[i+1,j+1] : Math.Max(expected[i+1,j], expected[i,j+1]);
        var matches = (List<(int a, int b)>)method.Invoke(null, new object[] { a, b });
        Require(matches.Count == expected[0,0], "Alignment is not a longest common subsequence: " + sample);
        int prevA = -1, prevB = -1;
        foreach (var pair in matches)
        {
            Require(pair.a > prevA && pair.b > prevB && a[pair.a] == b[pair.b], "Invalid alignment");
            prevA = pair.a; prevB = pair.b;
        }
    }
});
Check("Large script with aliases and 2000 comments", () =>
{
    string sql = string.Join("\n", Enumerable.Range(0, 2000).Select(i =>
        $"SELECT a.id x FROM dbo.t a WHERE a.id={i}; -- row {i}"));
    long before = GC.GetAllocatedBytesForCurrentThread();
    var timer = System.Diagnostics.Stopwatch.StartNew();
    string output = Format(sql, true);
    long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
    Require(Comments(sql).SequenceEqual(Comments(output)), "Large script comment mismatch");
    Require(Canonical(sql) == Canonical(output), "Large script SQL changed");
    Require(allocated < 256L * 1024 * 1024, "Excessive allocation: " + allocated);
    Console.WriteLine($"Large script: {timer.ElapsedMilliseconds} ms, {allocated / 1024 / 1024} MiB allocated");
});
Console.WriteLine($"{passed} passed, {failed} failed");
return failed == 0 ? 0 : 1;
