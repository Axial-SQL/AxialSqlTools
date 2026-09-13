using AxialSqlTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

internal static class Program
{
    private static int passed;
    private static int failed;

    private static int Main()
    {
        Name("bare", "Orders", null, null, "Orders");
        Name("schema", "sales.Orders", null, "sales", "Orders");
        Name("database", "Sales.dbo.Orders", "Sales", "dbo", "Orders");
        Name("omitted schema", "Sales..Orders", "Sales", null, "Orders");
        Name("delimited", " [Sales DB] . [my schema] . [Order Details]; ", "Sales DB", "my schema", "Order Details");
        Name("escaped brackets", "[Sales]]DB].[s]]].[a]]b]", "Sales]DB", "s]", "a]b");
        Name("double quotes", "\"Sales DB\".\"s\".\"Order\"\"Details\"", "Sales DB", "s", "Order\"Details");
        Name("punctuation in identifiers", "[a.b].[s'1].[o;/n]", "a.b", "s'1", "o;/n");
        Name("unicode", "[Продажи].[客户]", null, "Продажи", "客户");
        Name("multiline identifier", "dbo.[Order\nDetails]", null, "dbo", "Order\nDetails");
        Name("maximum length", new string('a', 128), null, null, new string('a', 128));
        foreach (string invalid in new[] { "", " ", "dbo.", ".Orders", "dbo..", "[broken", "\"broken", "[]",
            "server.db.dbo.Orders", "a...b", "Orders; DROP TABLE Orders", "Orders alias", "'Orders'", "N'Orders'",
            "dbo.Orders -- comment", "dbo.Orders /*comment*/", "Orders;;", "#Orders", "[#Orders]", "@Orders", new string('a', 129) })
            Check("reject " + invalid, () => Throws<ArgumentException>(() => SqlObjectName.Parse(invalid)));

        Caret("middle", "SELECT * FROM dbo.Or|ders;", "dbo.Orders");
        Caret("on schema", "SELECT * FROM d|bo.Orders;", "dbo.Orders");
        Caret("on dot", "SELECT * FROM dbo|.Orders;", "dbo.Orders");
        Caret("end", "SELECT * FROM dbo.Orders|;", "dbo.Orders");
        Caret("end of document", "dbo.Orders|", "dbo.Orders");
        Caret("quoted spaces", "SELECT * FROM [Sales DB].[dbo].[Order |Details];", "[Sales DB].[dbo].[Order Details]");
        Caret("escaped bracket", "SELECT * FROM [a]]|b];", "[a]]b]");
        Caret("multiline multipart", "SELECT * FROM [Sales DB]\r\n . dbo . Or|ders", "[Sales DB]\r\n . dbo . Orders");
        Caret("tabs and CRLF", "SELECT 1;\r\n\tSELECT * FROM dbo.Or|ders", "dbo.Orders");
        Caret("omitted schema", "SELECT * FROM Sales..Or|ders", "Sales..Orders");
        Caret("keep four-part input for rejection", "SELECT * FROM server.db.dbo.Or|ders", "server.db.dbo.Orders");
        Caret("unrelated incomplete SQL", "SELECT FROM; SELECT * FROM dbo.Or|ders", "dbo.Orders");
        Caret("double quoted caret", "SELECT * FROM \"Sales DB\".dbo.\"Order |Details\"", "\"Sales DB\".dbo.\"Order Details\"");
        Caret("avoid alias concatenation", "SELECT * FROM dbo.Or|ders o", "dbo.Orders");
        Caret("line comment", "-- dbo.Or|ders\nSELECT 1", null);
        Caret("multiline comment", "/* text\n dbo.Or|ders */\nSELECT 1", null);
        Caret("nested comment", "/* outer /* inner */ dbo.Or|ders */", null);
        Caret("string", "SELECT 'dbo.Or|ders';", null);
        Caret("unicode string", "SELECT N'dbo.Or|ders';", null);
        Caret("multiline string", "SELECT 'text\r\ndbo.Or|ders';", null);
        Caret("escaped string", "SELECT 'it''s dbo.Or|ders';", null);
        Caret("double quoted string after OFF", "SET QUOTED_IDENTIFIER OFF; SELECT \"dbo.Or|ders\";", null);
        Caret("double quoted identifier after ON", "SET QUOTED_IDENTIFIER OFF; SET QUOTED_IDENTIFIER ON; SELECT * FROM \"Or|ders\";", "\"Orders\"");
        Caret("brackets after OFF", "SET QUOTED_IDENTIFIER OFF; SELECT * FROM [Or|ders];", "[Orders]");
        Caret("keyword", "SEL|ECT * FROM dbo.Orders;", null);
        Caret("whitespace", "SELECT * FROM  | dbo.Orders;", null);

        Path("table", Item("USER_TABLE"), "/Database[@Name='Sales']/Table[@Name='Orders' and @Schema='dbo']");
        Path("view", Item("VIEW"), "/Database[@Name='Sales']/View[@Name='Orders' and @Schema='dbo']");
        foreach (string type in new[] { "SQL_SCALAR_FUNCTION", "SQL_INLINE_TABLE_VALUED_FUNCTION", "SQL_TABLE_VALUED_FUNCTION", "CLR_SCALAR_FUNCTION" })
            Path(type, Item(type), "/Database[@Name='Sales']/UserDefinedFunction[@Name='Orders' and @Schema='dbo']");
        Path("procedure", Item("SQL_STORED_PROCEDURE"), "/Database[@Name='Sales']/StoredProcedure[@Name='Orders' and @Schema='dbo']");
        Path("synonym", Item("SYNONYM"), "/Database[@Name='Sales']/Synonym[@Name='Orders' and @Schema='dbo']");
        Path("table type", Item("TYPE_TABLE"), "/Database[@Name='Sales']/UserDefinedTableType[@Name='Orders' and @Schema='dbo']");
        Path("index on view", Item("INDEX", "VIEW", "IX_Orders"), "/Database[@Name='Sales']/View[@Name='Orders' and @Schema='dbo']/Index[@Name='IX_Orders']");
        Path("trigger on view", Item("SQL_TRIGGER", "VIEW", "TR_Orders"), "/Database[@Name='Sales']/View[@Name='Orders' and @Schema='dbo']/Trigger[@Name='TR_Orders']");
        foreach (string type in new[] { "INDEX", "PRIMARY_KEY_CONSTRAINT", "UNIQUE_CONSTRAINT" })
            Path(type, Item(type, "USER_TABLE", "IX_Orders"), "/Database[@Name='Sales']/Table[@Name='Orders' and @Schema='dbo']/Index[@Name='IX_Orders']");
        Path("foreign key", Item("FOREIGN_KEY_CONSTRAINT", "USER_TABLE", "FK_Orders"), "/Database[@Name='Sales']/Table[@Name='Orders' and @Schema='dbo']/ForeignKey[@Name='FK_Orders']");
        Path("check", Item("CHECK_CONSTRAINT", "USER_TABLE", "CK_Orders"), "/Database[@Name='Sales']/Table[@Name='Orders' and @Schema='dbo']/Check[@Name='CK_Orders']");
        Path("default selects owning column", Item("DEFAULT_CONSTRAINT", "USER_TABLE", "DF_Orders"), "/Database[@Name='Sales']/Table[@Name='Orders' and @Schema='dbo']/Column[@Name='Status']");
        Path("escape URN names", new ScriptObjectSelectionItem("USER_TABLE", "s'c", "t'/x", 1, "db'o", 0, null),
            "/Database[@Name='db''o']/Table[@Name='t''/x' and @Schema='s''c']");
        Check("unsupported object fails clearly", () => Throws<NotSupportedException>(() => ObjectExplorerPath.GetRelativeSteps(Item("INTERNAL_TABLE"))));
        Check("missing parent fails clearly", () => Throws<InvalidOperationException>(() => ObjectExplorerPath.GetRelativeSteps(Item("INDEX"))));
        Check("database quoting", () => Equal("[db]]; DROP DATABASE x;--]", SqlObjectName.Quote("db]; DROP DATABASE x;--")));
        Check("same endpoint with comma spaces", () => True(Same("HOST, 1178", "host,1178")));
        Check("different ports stay different", () => True(!Same("HOST,1178", "HOST,1433")));
        Check("named instances stay different", () => True(!Same("HOST\\ONE", "HOST\\TWO")));
        Check("DNS names are not shortened", () => True(!Same("HOST.one.example", "HOST.two.example")));
        Check("different SQL logins stay different", () => True(!ObjectExplorerPath.SameConnection("HOST", "alice", false, "HOST", "bob", false)));
        Check("SQL and Windows auth stay different", () => True(!ObjectExplorerPath.SameConnection("HOST", "", true, "HOST", "", false)));

        Check("navigate existing connection", () =>
        {
            var host = new FakeHost { Connected = true };
            Navigate(host);
            Equal(0, host.ConnectCalls);
            Equal("Server[@Name='HOST']/Database[@Name='Sales']/Table[@Name='Orders' and @Schema='dbo']", host.Selected);
            True(host.Expanded[0] == "Server[@Name='HOST']");
        });
        Check("connect once then navigate after delayed registration", () =>
        {
            var host = new FakeHost { RegistrationDelay = 3 };
            Navigate(host);
            Equal(1, host.ConnectCalls);
            True(host.Selected != null);
        });
        Check("wait for lazy folders and target", () =>
        {
            var host = new FakeHost { Connected = true, LoadDelay = 2, SelectDelay = 3 };
            Navigate(host);
            True(host.Selected != null);
            Equal(4, host.SelectCalls);
        });
        Check("connection failure propagates", () =>
        {
            var host = new FakeHost { FailConnect = true };
            Throws<InvalidOperationException>(() => Navigate(host));
            Equal(1, host.ConnectCalls);
            Equal(null, host.Selected);
        });
        Check("cancellation before navigation does not connect", () =>
        {
            var host = new FakeHost();
            var token = new CancellationToken(true);
            Throws<OperationCanceledException>(() => ObjectExplorerNavigation.NavigateAsync(host,
                ObjectExplorerPath.GetRelativeSteps(Item("USER_TABLE")), token).GetAwaiter().GetResult());
            Equal(0, host.ConnectCalls);
        });
        Check("missing target can time out", () =>
        {
            var host = new FakeHost { Connected = true, SelectDelay = int.MaxValue };
            using (var cancellation = new CancellationTokenSource())
            {
                int waits = 0;
                Throws<OperationCanceledException>(() => ObjectExplorerNavigation.NavigateAsync(host,
                    ObjectExplorerPath.GetRelativeSteps(Item("USER_TABLE")), cancellation.Token, token =>
                    {
                        if (++waits == 12) cancellation.Cancel();
                        token.ThrowIfCancellationRequested();
                        return Task.CompletedTask;
                    }).GetAwaiter().GetResult());
                Equal(null, host.Selected);
            }
        });
        Console.WriteLine($"{passed} passed; {failed} failed.");
        return failed == 0 ? 0 : 1;
    }

    private static void Navigate(FakeHost host) => ObjectExplorerNavigation.NavigateAsync(host,
        ObjectExplorerPath.GetRelativeSteps(Item("USER_TABLE")), CancellationToken.None,
        token => Task.CompletedTask).GetAwaiter().GetResult();

    private static bool Same(string requested, string actual) => ObjectExplorerPath.SameConnection(requested, "", true, actual, "", true);
    private static ScriptObjectSelectionItem Item(string type, string parentType = null, string name = "Orders") =>
        new ScriptObjectSelectionItem(type, "dbo", name, 1, "Sales", parentType == null ? 0 : 2,
            parentType == null ? null : "Orders", parentType, "Status");
    private static void Path(string title, ScriptObjectSelectionItem item, string expected) => Check(title, () => Equal(expected, ObjectExplorerPath.GetRelativeSteps(item).Last()));
    private static void Name(string title, string input, string db, string schema, string name) => Check(title, () =>
    {
        var result = SqlObjectName.Parse(input);
        Equal(db, result.DatabaseName); Equal(schema, result.SchemaName); Equal(name, result.ObjectName);
    });
    private static void Caret(string title, string input, string expected) => Check(title, () =>
    {
        int offset = input.IndexOf('|');
        Equal(expected, SqlObjectName.GetNameAtCaret(input.Remove(offset, 1), offset));
    });
    private static void Check(string title, Action test)
    {
        try { test(); passed++; }
        catch (Exception ex) { failed++; Console.Error.WriteLine(title + ": " + ex.Message); }
    }
    private static void Equal<T>(T expected, T actual)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new Exception($"Expected '{expected}', got '{actual}'.");
    }
    private static void True(bool value) { if (!value) throw new Exception("Assertion failed."); }
    private static void Throws<T>(Action action) where T : Exception
    {
        try { action(); } catch (T) { return; }
        throw new Exception("Expected " + typeof(T).Name);
    }

    private sealed class FakeHost : IObjectExplorerNavigationHost
    {
        public bool Connected;
        public bool FailConnect;
        public int RegistrationDelay;
        public int LoadDelay;
        public int SelectDelay;
        public int ConnectCalls;
        public int SelectCalls;
        public string Selected;
        public List<string> Expanded = new List<string>();
        private readonly Dictionary<string, int> attempts = new Dictionary<string, int>();

        public string FindServerContext() => Connected && RegistrationDelay-- <= 0 ? "Server[@Name='HOST']" : null;
        public void Connect()
        {
            ConnectCalls++;
            if (FailConnect) throw new InvalidOperationException("Connection failed.");
            Connected = true;
        }
        public bool TryExpand(string urn)
        {
            attempts.TryGetValue(urn, out int count);
            attempts[urn] = count + 1;
            if (count < LoadDelay) return false;
            Expanded.Add(urn);
            return true;
        }
        public bool TrySelect(string urn)
        {
            SelectCalls++;
            if (SelectCalls <= SelectDelay) return false;
            Selected = urn;
            return true;
        }
    }
}
