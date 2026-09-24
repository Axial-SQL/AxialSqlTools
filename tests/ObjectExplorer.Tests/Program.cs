using AxialSqlTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

internal static partial class Program
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

        Check("unsupported object fails clearly", () => Throws<NotSupportedException>(() => ObjectExplorerPath.GetTreeSteps(Item("INTERNAL_TABLE"))));
        Check("missing parent fails clearly", () => Throws<InvalidOperationException>(() => ObjectExplorerPath.GetTreeSteps(Item("INDEX"))));
        Check("database quoting", () => Equal("[db]]; DROP DATABASE x;--]", SqlObjectName.Quote("db]; DROP DATABASE x;--")));
        Check("same endpoint with comma spaces", () => True(Same("HOST, 1178", "host,1178")));
        Check("different ports stay different", () => True(!Same("HOST,1178", "HOST,1433")));
        Check("named instances stay different", () => True(!Same(@"HOST\ONE", @"HOST\TWO")));
        Check("DNS names are not shortened", () => True(!Same("HOST.one.example", "HOST.two.example")));
        Check("different SQL logins stay different", () => True(!ObjectExplorerPath.SameConnection("HOST", "alice", false, "HOST", "bob", false)));
        Check("SQL and Windows auth stay different", () => True(!ObjectExplorerPath.SameConnection("HOST", "", true, "HOST", "", false)));
        TreeTests();
        HostTests();
        Console.WriteLine($"{passed} passed; {failed} failed.");
        return failed == 0 ? 0 : 1;
    }

    private static void Navigate(FakeHost host, ScriptObjectSelectionItem item = null) => ObjectExplorerNavigation.NavigateAsync(host,
        item ?? Item("USER_TABLE"), CancellationToken.None,
        token => Task.CompletedTask).GetAwaiter().GetResult();

    private static bool Same(string requested, string actual) => ObjectExplorerPath.SameConnection(requested, "", true, actual, "", true);
    private static ScriptObjectSelectionItem Item(string type, string parentType = null, string name = "Orders") =>
        new ScriptObjectSelectionItem(type, "dbo", name, 1, "Sales", parentType == null ? 0 : 2,
            parentType == null ? null : "Orders", parentType, "Status");
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
        public bool Connected = true;
        public bool FailConnect;
        public int RegistrationDelay;
        public int ConnectCalls;
        public IObjectExplorerTreeNode Selected;
        public FakeNode Root;

        public IObjectExplorerTreeNode FindServer() => Connected && RegistrationDelay-- <= 0 ? Root : null;
        public void Connect()
        {
            ConnectCalls++;
            if (FailConnect) throw new InvalidOperationException("Connection failed.");
            Connected = true;
        }
        public void Select(IObjectExplorerTreeNode node) => Selected = node;
    }

    private sealed class FakeNode : IObjectExplorerTreeNode
    {
        public string Name { get; set; }
        public string InvariantName { get; set; }
        public string UniqueName { get; set; }
        public string UrnPath { get; set; }
        public string NavigationContext { get; set; }
        public bool IsFolder { get; set; }
        public int Loads;
        public Action OnLoad;
        public List<IObjectExplorerTreeNode> Children = new List<IObjectExplorerTreeNode>();
        public IList<IObjectExplorerTreeNode> LoadChildren()
        {
            Loads++;
            OnLoad?.Invoke();
            return Children;
        }
    }
}
