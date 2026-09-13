using AxialSqlTools.QuerySafety;
using System;
using System.Collections.Generic;

internal static class Program
{
    private static int count;
    internal static void Check(bool condition, string message = "Assertion failed") { if (!condition) throw new Exception(message); }
    internal static void Test(string name, Action run) { run(); count++; Console.WriteLine("PASS " + name); }
    private static void Sql(string name, string sql, int actions)
    {
        Test(name, () => {
            var result = FatalActionAnalyzer.Analyze(sql);
            Check(result.Limitation == null, result.Limitation);
            Check(result.Actions.Count == actions, "Expected " + actions + " actions, found " + result.Actions.Count);
        });
    }

    private static void Main()
    {
        Sql("UPDATE without WHERE", "UPDATE dbo.T SET x=1", 1);
        Sql("DELETE without WHERE", "DELETE FROM dbo.T", 1);
        Sql("TRUNCATE", "TRUNCATE TABLE dbo.T", 1);
        Sql("partition TRUNCATE", "TRUNCATE TABLE dbo.T WITH (PARTITIONS (1 TO 3))", 1);
        Sql("UPDATE with WHERE", "UPDATE dbo.T SET x=1 WHERE id=2", 0);
        Sql("DELETE with WHERE", "DELETE dbo.T WHERE id=2", 0);
        Sql("TOP is not WHERE", "UPDATE TOP (1) dbo.T SET x=1; DELETE TOP (1) FROM dbo.T;", 2);
        Sql("comment WHERE is not a filter", "UPDATE dbo.T SET x=1 -- WHERE id=2\n; DELETE dbo.T /* WHERE id=2 */", 2);
        Sql("string WHERE is not a filter", "UPDATE dbo.T SET x=N'WHERE id=2'", 1);
        Sql("subquery WHERE does not filter the UPDATE", "UPDATE dbo.T SET x=(SELECT max(x) FROM dbo.S WHERE y=1)", 1);
        Sql("another statement WHERE does not filter DELETE", "SELECT * FROM dbo.S WHERE x=1; DELETE dbo.T;", 1);
        Sql("commented actions are not executed", "/* UPDATE dbo.T SET x=1; */ SELECT 1; -- TRUNCATE TABLE dbo.T", 0);
        Sql("strings are not statements", "SELECT N'UPDATE dbo.T SET x=1; DELETE dbo.T; TRUNCATE TABLE dbo.T'", 0);
        Sql("quoted identifiers are not keywords", "SELECT [DELETE], [TRUNCATE], [UPDATE] FROM dbo.T", 0);
        Sql("GO and control flow", "IF 1=1 BEGIN UPDATE dbo.T SET x=1; END\nGO 2\nBEGIN TRY DELETE dbo.T END TRY BEGIN CATCH SELECT 1 END CATCH", 2);
        Sql("GO repeat text inside string remains a string", "SELECT N'GO 2\nDELETE dbo.T';", 0);
        Sql("case-only alias cannot exempt a real target on case-sensitive databases", "UPDATE T SET x=1 FROM #t t", 1);
        Sql("case-only CTE cannot exempt a real target on case-sensitive databases", "WITH t AS (SELECT * FROM #t) DELETE T", 1);
        Sql("permanent target alias", "UPDATE t SET x=1 FROM dbo.T t JOIN #s s ON t.id=s.id; DELETE t FROM dbo.T t", 2);
        Sql("WHERE in join condition is not statement WHERE", "DELETE t FROM dbo.T t JOIN dbo.S s ON t.id=s.id", 1);
        Sql("direct local and global temp tables", "UPDATE #t SET x=1; DELETE FROM ##t; TRUNCATE TABLE #t", 0);
        Sql("qualified bracketed temp tables", "UPDATE tempdb..[#t] SET x=1; DELETE tempdb.dbo.[##t]; TRUNCATE TABLE [tempdb]..[#t]", 0);
        Sql("double quoted temp tables", "UPDATE \"#t\" SET x=1; DELETE FROM \"#t\"", 0);
        Sql("table variables", "DECLARE @t TABLE (x int); UPDATE @t SET x=1; DELETE FROM @t", 0);
        Sql("quoted at-sign name is a real table", "DELETE FROM dbo.[@t]; UPDATE [@t] SET x=1", 2);
        Sql("temp table aliases in joins", "UPDATE t SET x=1 FROM #t t JOIN dbo.S s ON t.id=s.id; DELETE t FROM #t t JOIN dbo.S s ON t.id=s.id", 0);
        Sql("table variable aliases", "DECLARE @t TABLE (x int); UPDATE t SET x=1 FROM @t t; DELETE t FROM @t t", 0);
        Sql("hash alias on a permanent target still warns", "DELETE [#alias] FROM dbo.T [#alias]", 1);
        Sql("subquery temp alias must not hide outer permanent target", "DELETE t FROM dbo.T t WHERE EXISTS (SELECT 1 FROM #s t); UPDATE t SET x=(SELECT max(x) FROM #s t) FROM dbo.T t", 1);
        Sql("temp CTE target", "WITH c AS (SELECT * FROM #t) DELETE FROM c", 0);
        Sql("table variable CTE target", "DECLARE @t TABLE (x int); WITH c AS (SELECT * FROM @t) UPDATE c SET x=1", 0);
        Sql("chained temp CTE", "WITH a AS (SELECT * FROM #t), b AS (SELECT * FROM a) DELETE FROM b", 0);
        Sql("temp CTE alias", "WITH c AS (SELECT * FROM #t) DELETE t FROM c t", 0);
        Sql("permanent CTE target", "WITH c AS (SELECT * FROM dbo.T WHERE id=1) DELETE FROM c", 1);
        Sql("qualified table is not same-name temp CTE", "WITH T AS (SELECT * FROM #t) DELETE FROM dbo.T", 1);
        Sql("mixed-source CTE cannot prove temporary target", "WITH c AS (SELECT t.x FROM #t t JOIN dbo.S s ON t.id=s.id) UPDATE c SET x=1", 1);
        Sql("recursive CTE terminates conservatively", "WITH c AS (SELECT * FROM #t UNION ALL SELECT * FROM c) DELETE FROM c", 1);
        Sql("temp and permanent actions in same script", "TRUNCATE TABLE #t; DELETE dbo.T; UPDATE @t SET x=1; TRUNCATE TABLE dbo.S", 2);
        Sql("OUTPUT does not hide DELETE target", "DELETE dbo.T OUTPUT deleted.id INTO #log", 1);
        Sql("OUTPUT into permanent table does not change temp target", "DELETE #t OUTPUT deleted.id INTO dbo.Log", 0);
        Sql("CREATE PROCEDURE is not running its body", "CREATE PROCEDURE dbo.P AS DELETE dbo.T; TRUNCATE TABLE dbo.T;", 0);
        Sql("ALTER PROCEDURE is not running its body", "ALTER PROCEDURE dbo.P AS UPDATE dbo.T SET x=1", 0);
        Sql("CREATE OR ALTER trigger is not running its body", "CREATE OR ALTER TRIGGER dbo.tr ON dbo.T AFTER INSERT AS DELETE dbo.S", 0);
        Sql("CREATE function is not running its body", "CREATE FUNCTION dbo.f() RETURNS @t TABLE (x int) AS BEGIN DELETE @t; RETURN; END", 0);
        Sql("DROP is outside the requested scope", "DROP TABLE dbo.T", 0);
        Sql("EXEC is not expanded", "EXEC dbo.P; EXEC(N'DELETE dbo.T');", 0);
        Sql("blank or safe SQL", " \nSELECT 1", 0);
        Test("incomplete analysis needs confirmation", () => { var r = FatalActionAnalyzer.Analyze("DELETE FROM dbo.T; nonsense syntax here;"); Check(r.RequiresConfirmation && r.Limitation != null); });
        Test("line numbers match execution text", () => Check(FatalActionAnalyzer.Analyze("SELECT 1;\n\nDELETE dbo.T").Actions[0].Line == 3));
        Test("unchanged execution keeps approval", () => { var s=new FatalActionSession(); s.Approve("DELETE dbo.T", "c"); Check(s.IsApproved("DELETE dbo.T", "c") && s.IsApproved("DELETE dbo.T", "c")); });
        Test("changed SQL invalidates approval including reverting", () => { var s=new FatalActionSession(); s.Approve("DELETE dbo.T", "c"); Check(!s.IsApproved("DELETE dbo.S", "c") && !s.IsApproved("DELETE dbo.T", "c")); });
        Test("connection change invalidates approval", () => { var s=new FatalActionSession(); s.Approve("DELETE dbo.T", "c1"); Check(!s.IsApproved("DELETE dbo.T", "c2") && !s.IsApproved("DELETE dbo.T", "c1")); });
        Test("unknown session cannot be approved", () => { var s=new FatalActionSession(); s.Approve("DELETE dbo.T", null); Check(!s.IsApproved("DELETE dbo.T", null)); });
        Test("approval is local to a window", () => { var a=new FatalActionSession(); var b=new FatalActionSession(); a.Approve("DELETE dbo.T", "c"); Check(!b.IsApproved("DELETE dbo.T", "c")); });
        Test("explicit reset removes approval", () => { var s=new FatalActionSession(); s.Approve("DELETE dbo.T", "c"); s.Clear(); Check(!s.IsApproved("DELETE dbo.T", "c")); });
        GuardTests.Run();
        Console.WriteLine(count + " checks passed.");
    }
}
