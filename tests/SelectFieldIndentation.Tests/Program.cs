using AxialSqlTools;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static AxialSqlTools.SettingsManager;

internal static class Program
{
    private static int count;
    private static void Check(bool value, string message) { if (!value) throw new Exception(message); }
    private static void Test(string name, Action test) { test(); count++; Console.WriteLine("PASS " + name); }
    private sealed class Queries : TSqlFragmentVisitor
    {
        public List<QuerySpecification> Items = new List<QuerySpecification>();
        public override void ExplicitVisit(QuerySpecification node) { Items.Add(node); base.ExplicitVisit(node); }
    }
    private static TSqlFragment Parse(string sql)
    {
        IList<ParseError> errors;
        var result = new TSql170Parser(false).Parse(new StringReader(sql), out errors);
        Check(errors.Count == 0, "Output did not parse: " + string.Join("; ", errors.Select(e => e.Message)) + "\n" + sql);
        return result;
    }
    private static string Tokens(string sql) => string.Join("|", Parse(sql).ScriptTokenStream
        .Where(t => t.TokenType != TSqlTokenType.WhiteSpace && t.TokenType != TSqlTokenType.EndOfFile)
        .Select(t => t.TokenType + ":" + t.Text));
    private static void FormatCase(string name, string sql, TSqlCodeFormatSettings options = null, bool affected = true)
    {
        Test(name, () => {
            var settings = options ?? new TSqlCodeFormatSettings();
            settings.breakSelectFieldsAfterTopAndUnindent = false;
            string baseline = TSqlFormatter.FormatCode(sql, settings);
            settings.breakSelectFieldsAfterTopAndUnindent = true;
            string result = TSqlFormatter.FormatCode(sql, settings);
            Check(Tokens(baseline) == Tokens(result), "Changed SQL tokens or comment contents\n" + result);
            var visitor = new Queries(); Parse(result).Accept(visitor);
            var matches = visitor.Items.Where(q => q.TopRowFilter != null || q.UniqueRowFilter == UniqueRowFilter.Distinct).ToList();
            if (affected) Check(matches.Count > 0, "Fixture must contain a SELECT modifier");
            else Check(result == baseline, "An unrelated query changed");
            foreach (var query in matches)
            {
                int lastLine = query.TopRowFilter == null ? query.StartLine : query.ScriptTokenStream[query.TopRowFilter.LastTokenIndex].Line;
                foreach (var field in query.SelectElements)
                {
                    Check(field.StartLine > lastLine, "Field must start on a new line\n" + result);
                    Check(field.StartColumn == query.StartColumn + 4, "Field must be one normal indent from SELECT\n" + result);
                    lastLine = field.ScriptTokenStream[field.LastTokenIndex].Line;
                }
            }
            Check(TSqlFormatter.FormatCode(result, settings) == result, "Repeated formatting must be stable\n" + result);
        });
    }
    private static void Main()
    {
        FormatCase("DISTINCT only", "SELECT DISTINCT a,b,c FROM dbo.T");
        FormatCase("TOP only", "SELECT TOP (100) a,b,c FROM dbo.T");
        FormatCase("DISTINCT and TOP", "SELECT DISTINCT TOP (100) a,b,c FROM dbo.T");
        FormatCase("TOP legacy syntax", "SELECT TOP 10 a,b FROM dbo.T");
        FormatCase("TOP variable and PERCENT WITH TIES", "SELECT DISTINCT TOP (@n) PERCENT WITH TIES a,b FROM dbo.T ORDER BY a");
        FormatCase("TOP arithmetic expression", "SELECT TOP (1 + @n) a,b FROM dbo.T");
        FormatCase("single field", "SELECT DISTINCT a FROM dbo.T");
        FormatCase("star", "SELECT TOP (1) t.* FROM dbo.T t");
        FormatCase("SELECT without FROM", "SELECT DISTINCT TOP (1) 1 a, 2 b");
        FormatCase("INTO and remaining clauses", "SELECT DISTINCT TOP (10) a,b INTO #t FROM dbo.T WHERE a>0 GROUP BY a,b HAVING COUNT(*)>1 ORDER BY a");
        FormatCase("CTE", "WITH c AS (SELECT DISTINCT TOP (10) a,b FROM dbo.T) SELECT * FROM c");
        FormatCase("multiple CTEs", "WITH c AS (SELECT DISTINCT a,b FROM dbo.T), d AS (SELECT TOP (1) a,b FROM c) SELECT DISTINCT a,b FROM d");
        FormatCase("derived table", "SELECT x.a FROM (SELECT DISTINCT TOP (5) a,b FROM dbo.T) x");
        FormatCase("scalar subquery", "SELECT DISTINCT (SELECT TOP (1) a FROM dbo.S ORDER BY a) AS a, b FROM dbo.T");
        FormatCase("multiple nesting levels", "SELECT TOP (10) (SELECT TOP (1) (SELECT DISTINCT TOP (1) a FROM dbo.S) FROM dbo.R) AS x,b FROM dbo.T");
        FormatCase("nested query in predicate", "SELECT DISTINCT a,b FROM dbo.T WHERE a IN (SELECT TOP (2) a FROM dbo.S)");
        FormatCase("UNION", "SELECT DISTINCT a,b FROM dbo.T UNION ALL SELECT TOP (10) a,b FROM dbo.S");
        FormatCase("parenthesized expression", "SELECT TOP (10) COALESCE(a, (SELECT TOP (1) b FROM dbo.S)), b FROM dbo.T");
        FormatCase("function commas are not fields", "SELECT DISTINCT CONCAT(a,',',b),COALESCE(a,b,0) FROM dbo.T");
        FormatCase("window expression", "SELECT DISTINCT TOP (4) ROW_NUMBER() OVER (PARTITION BY a ORDER BY b),a,b FROM dbo.T");
        FormatCase("CASE expression", "SELECT DISTINCT CASE WHEN a=1 THEN 'x' ELSE 'y' END AS result,b FROM dbo.T");
        FormatCase("nested CASE and SELECT", "SELECT TOP (1) CASE WHEN a=1 THEN (SELECT DISTINCT TOP (1) x FROM dbo.S) ELSE 0 END AS c,b FROM dbo.T");
        FormatCase("multiline string stays unchanged", "SELECT TOP (1) N'line 1\n      SELECT DISTINCT TOP (5) a,b\nline 3',a FROM dbo.T");
        FormatCase("variable assignments", "SELECT TOP (1) @a=a,@b=b FROM dbo.T");
        FormatCase("plain SELECT stays unchanged", "SELECT a,b FROM dbo.T", affected:false);
        FormatCase("plain ALL stays unchanged", "SELECT ALL a,b FROM dbo.T", affected:false);
        FormatCase("DISTINCT aggregate stays unchanged", "SELECT COUNT(DISTINCT a),b FROM dbo.T GROUP BY b", affected:false);
        FormatCase("UPDATE TOP stays unchanged", "UPDATE TOP (5) dbo.T SET a=1,b=2", affected:false);
        FormatCase("keyword strings and identifiers", "SELECT [DISTINCT], [TOP], 'SELECT DISTINCT TOP (1) a,b' FROM dbo.T", affected:false);
        FormatCase("multiple batches", "SELECT DISTINCT a,b FROM dbo.T\nGO\nSELECT TOP (5) c,d FROM dbo.S");
        FormatCase("procedure and nested blocks", "CREATE PROCEDURE dbo.P AS BEGIN IF 1=1 BEGIN SELECT DISTINCT TOP (5) a,b FROM dbo.T END END");
        FormatCase("inline block comments", "SELECT DISTINCT /* header */ TOP (10) /* fields */ a /* first */, b FROM dbo.T", new TSqlCodeFormatSettings { preserveComments=true });
        FormatCase("line comments", "SELECT DISTINCT TOP (10) -- header\na, -- first\nb FROM dbo.T", new TSqlCodeFormatSettings { preserveComments=true });
        FormatCase("multiline block comment contents", "SELECT DISTINCT TOP (1) /* keep\n    spaces\n*/ a,b FROM dbo.T", new TSqlCodeFormatSettings { preserveComments=true });
        FormatCase("CASE option", "SELECT DISTINCT TOP (1) CASE WHEN a=1 THEN 1 ELSE 0 END AS c,b FROM dbo.T", new TSqlCodeFormatSettings { formatCaseAsMultiline=true });
        FormatCase("BEGIN END option", "WHILE 1=0 BEGIN SELECT DISTINCT TOP (1) a,b FROM dbo.T END", new TSqlCodeFormatSettings { unindentBeginEndBlocks=true });
        FormatCase("JOIN options", "SELECT DISTINCT TOP (5) t.a,s.b FROM dbo.T t JOIN dbo.S s ON t.a=s.a CROSS JOIN dbo.R r", new TSqlCodeFormatSettings { removeNewLineAfterJoin=true,addTabAfterJoinOn=true,moveCrossJoinToNewLine=true });
        var all = new TSqlCodeFormatSettings();
        foreach (var field in typeof(TSqlCodeFormatSettings).GetFields()) if (field.FieldType==typeof(bool)) field.SetValue(all,true);
        FormatCase("all advanced options", "CREATE PROCEDURE dbo.P @a int, @b int AS BEGIN SELECT DISTINCT TOP (10) t.a,CASE WHEN t.b>0 THEN GETDATE() ELSE NULL END AS c FROM dbo.T t JOIN dbo.S s ON t.a=s.a; END", all);
        Test("option defaults off and old settings load", () => Check(!JsonConvert.DeserializeObject<TSqlCodeFormatSettings>("{}").breakSelectFieldsAfterTopAndUnindent, "Default changed"));
        Test("option activates additional formatter by itself", () => Check(new TSqlCodeFormatSettings { breakSelectFieldsAfterTopAndUnindent=true }.HasAnyFormattingEnabled(), "Option was not registered"));
        Test("existing JSON key round-trips", () => {
            var saved=new TSqlCodeFormatSettings { breakSelectFieldsAfterTopAndUnindent=true };
            Check(JsonConvert.DeserializeObject<TSqlCodeFormatSettings>(JsonConvert.SerializeObject(saved)).breakSelectFieldsAfterTopAndUnindent, "Lost saved setting");
        });
        Console.WriteLine(count + " checks passed.");
    }
}
