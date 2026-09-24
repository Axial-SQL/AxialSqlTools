using AxialSqlTools;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

internal static class Program
{
    private static int passed;
    private static int Main()
    {
        try
        {
            ObjectResult("USER_TABLE", "Table", "UserTables", "Column");
            ObjectResult("VIEW", "View", "Views", "Definition");
            ObjectResult("SQL_STORED_PROCEDURE", "StoredProcedure", "StoredProcedures", "Parameter");
            ObjectResult("SQL_SCALAR_FUNCTION", "UserDefinedFunction", "Scalar-valuedFunctions", "Parameter");
            ObjectResult("SQL_INLINE_TABLE_VALUED_FUNCTION", "UserDefinedFunction", "Table-valuedFunctions", "Definition");
            ObjectResult("SQL_TABLE_VALUED_FUNCTION", "UserDefinedFunction", "Table-valuedFunctions", "Definition");

            var job = Object("Job", "Nightly / O'Brien ] job");
            var databases = Folder("Databases");
            var jobHost = new Host(Folder("Server", databases, Folder("SQLServerAgent", Folder("Jobs", job))));
            var jobResult = Result("SQL_AGENT_JOB", "msdb", "dbo", job.Name, "JobStep");
            Navigate(jobHost, jobResult);
            Equal(job, jobHost.Selected);
            Equal(0, databases.Loads);
            passed++;

            foreach (string column in new[] { "NavigationTypeDesc", "ScriptObjectName", "ScriptSchemaName", "ScriptDatabaseName" })
            {
                var invalid = Result("USER_TABLE", "Sales", "dbo", "Orders", "Table Name");
                invalid[column] = DBNull.Value;
                Throws<InvalidOperationException>(() => QuickSearchNavigationTarget.FromResult(invalid));
                passed++;
            }

            var missing = Result("USER_TABLE", "Sales", "dbo", "Orders", "Table Name");
            missing.Table.Columns.Remove("NavigationTypeDesc");
            Throws<InvalidOperationException>(() => QuickSearchNavigationTarget.FromResult(missing));
            passed++;

            var host = new Host(Folder("Server"));
            Throws<OperationCanceledException>(() => Navigate(host, jobResult, new CancellationToken(true)));
            Equal(null, host.Selected);
            Equal(0, host.Root.Loads);
            passed++;
            Console.WriteLine(passed + " Quick Search navigation checks passed.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }

    private static void ObjectResult(string type, string urnType, string folder, string matchLocation)
    {
        const string database = "Sales ] O'Brien";
        const string schema = "report.schema";
        const string name = "object.名 ] '";
        var target = Object(urnType, name, schema);
        // Identical displayed captions can have different schema/object boundaries.
        var collision = Object(urnType, "schema." + name, "report");
        var wrongDatabase = Object("Database", "Other", null, Folder(folder, Object(urnType, name, schema)));
        var branch = Folder(folder, collision, target);
        if (urnType == "UserDefinedFunction") branch = Folder("UsrDbFunctions", branch);
        if (urnType == "UserDefinedFunction" || urnType == "StoredProcedure") branch = Folder("UserProgrammability", branch);
        var correctDatabase = Object("Database", database, null, branch);
        var host = new Host(Folder("Server", Folder("Databases", wrongDatabase, correctDatabase)));
        Navigate(host, Result(type, database, schema, name, matchLocation));
        Equal(target, host.Selected);
        Equal(0, wrongDatabase.Loads);
        passed++;
    }

    private static DataRow Result(string type, string database, string schema, string name, string location)
    {
        var table = new DataTable();
        foreach (string column in new[] { "NavigationTypeDesc", "ScriptDatabaseName", "ScriptSchemaName", "ScriptObjectName", "MatchLocation", "ObjectName", "SourceText" })
            table.Columns.Add(column, typeof(string));
        // The result caption and text must never be parsed as an identifier.
        return table.Rows.Add(type, database, schema, name, location, "decorated display / Step 2", "SELECT 'unrelated';");
    }

    private static void Navigate(Host host, DataRow row, CancellationToken token = default(CancellationToken)) =>
        ObjectExplorerNavigation.NavigateAsync(host, QuickSearchNavigationTarget.FromResult(row), token,
            ct => Task.CompletedTask).GetAwaiter().GetResult();

    private static Node Folder(string id, params Node[] children) => new Node
    {
        Name = "localized caption", UniqueName = id, IsFolder = true,
        Children = children.Cast<IObjectExplorerTreeNode>().ToList()
    };
    private static Node Object(string type, string name, string schema = null, params Node[] children) => new Node
    {
        Name = name, InvariantName = schema == null ? name : schema + "." + name, UrnPath = "Server/" + type,
        NavigationContext = "Server/" + type + "[@Name='" + name.Replace("'", "''") + "'"
            + (schema == null ? "" : " and @Schema='" + schema.Replace("'", "''") + "'") + "]",
        Children = children.Cast<IObjectExplorerTreeNode>().ToList()
    };
    private static void Equal(object expected, object actual)
    {
        if (!Equals(expected, actual)) throw new Exception("Expected " + expected + ", got " + actual);
    }
    private static void Throws<T>(Action action) where T : Exception
    {
        try { action(); } catch (T) { return; }
        throw new Exception("Expected " + typeof(T).Name);
    }
    private sealed class Host : IObjectExplorerNavigationHost
    {
        public Host(Node root) { Root = root; }
        public Node Root;
        public IObjectExplorerTreeNode Selected;
        public IObjectExplorerTreeNode FindServer() => Root;
        public void Connect() => throw new Exception("Existing search connection should be reused.");
        public void Select(IObjectExplorerTreeNode node) { Selected = node; }
    }
    private sealed class Node : IObjectExplorerTreeNode
    {
        public string Name { get; set; }
        public string InvariantName { get; set; }
        public string UniqueName { get; set; }
        public string UrnPath { get; set; }
        public string NavigationContext { get; set; }
        public bool IsFolder { get; set; }
        public int Loads;
        public IList<IObjectExplorerTreeNode> Children;
        public IList<IObjectExplorerTreeNode> LoadChildren() { Loads++; return Children; }
    }
}
