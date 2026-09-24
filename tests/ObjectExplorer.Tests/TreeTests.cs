using AxialSqlTools;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

internal static partial class Program
{
    private static FakeNode Folder(string id, params FakeNode[] children) => new FakeNode
    {
        Name = "localized " + id, UniqueName = id, IsFolder = true,
        Children = children.Cast<IObjectExplorerTreeNode>().ToList()
    };
    private static FakeNode Object(string type, string name, string schema = null, params FakeNode[] children) => new FakeNode
    {
        Name = name, InvariantName = schema == null ? name : schema + "." + name,
        UrnPath = "Server/Database/" + type,
        NavigationContext = "Server[@Name='HOST']/" + type + "[@Name='" + name.Replace("'", "''") + "'"
            + (schema == null ? "" : " and @Schema='" + schema.Replace("'", "''") + "'") + "]",
        Children = children.Cast<IObjectExplorerTreeNode>().ToList()
    };
    private static FakeHost Host(params FakeNode[] children) => new FakeHost
    {
        Root = Object("Server", "HOST", null, Folder("Databases", Object("Database", "Sales", null, children)))
    };
    private static void TreeTests()
    {
        Check("cold tree loads synchronously before reading children", () =>
        {
            var target = Object("Table", "Orders", "dbo");
            var folder = Folder("UserTables");
            folder.OnLoad = () => folder.Children.Add(target);
            var host = Host(folder);
            Navigate(host);
            Equal(target, host.Selected);
            Equal(1, folder.Loads);
            Equal(0, host.ConnectCalls);
        });
        Check("auto-connect once then wait for root registration", () =>
        {
            var target = Object("Table", "Orders", "dbo");
            var host = Host(Folder("UserTables", target));
            host.Connected = false;
            host.RegistrationDelay = 3;
            Navigate(host);
            Equal(1, host.ConnectCalls);
            Equal(target, host.Selected);
        });
        Check("connection failures propagate", () =>
        {
            var host = Host(); host.Connected = false; host.FailConnect = true;
            Throws<InvalidOperationException>(() => Navigate(host));
            Equal(1, host.ConnectCalls); Equal(null, host.Selected);
        });
        foreach (bool connecting in new[] { false, true })
            Check("cancellation while " + (connecting ? "connecting" : "loading"), () =>
            {
                var host = Host(Folder("UserTables", Object("Table", "Orders", "dbo")));
                host.Connected = !connecting;
                host.RegistrationDelay = connecting ? int.MaxValue : 0;
                using (var cancellation = new CancellationTokenSource())
                {
                    int waits = 0;
                    Throws<OperationCanceledException>(() => ObjectExplorerNavigation.NavigateAsync(host, Item("USER_TABLE"),
                        cancellation.Token, token =>
                        {
                            if (++waits == 3) cancellation.Cancel();
                            token.ThrowIfCancellationRequested(); return Task.CompletedTask;
                        }).GetAwaiter().GetResult());
                    Equal(null, host.Selected);
                }
            });
        Check("cancelled command does not connect", () =>
        {
            var host = Host(); host.Connected = false;
            Throws<OperationCanceledException>(() => ObjectExplorerNavigation.NavigateAsync(host, Item("USER_TABLE"),
                new CancellationToken(true)).GetAwaiter().GetResult());
            Equal(0, host.ConnectCalls);
        });
        Check("cancellation during native enumeration prevents selection", () =>
        {
            using (var cancellation = new CancellationTokenSource())
            {
                var folder = Folder("UserTables", Object("Table", "Orders", "dbo"));
                folder.OnLoad = cancellation.Cancel;
                var host = Host(folder);
                Throws<OperationCanceledException>(() => ObjectExplorerNavigation.NavigateAsync(host, Item("USER_TABLE"),
                    cancellation.Token, token => Task.CompletedTask).GetAwaiter().GetResult());
                Equal(null, host.Selected);
            }
        });
        Check("filtered object stops without repeated enumeration", () =>
        {
            var folder = Folder("UserTables");
            var host = Host(folder);
            Throws<InvalidOperationException>(() => Navigate(host));
            Equal(1, folder.Loads); Equal(null, host.Selected);
        });
        Check("unrelated databases and objects are not expanded", () =>
        {
            var wrong = Object("Table", "Other", "dbo");
            var target = Object("Table", "Orders", "dbo");
            var security = Folder("Security");
            var host = Host(Folder("UserTables", wrong, target), security);
            var otherDb = Object("Database", "OtherDb");
            ((FakeNode)host.Root.Children[0]).Children.Insert(0, otherDb);
            Navigate(host);
            Equal(target, host.Selected); Equal(0, wrong.Loads); Equal(0, otherDb.Loads); Equal(0, security.Loads);
        });
        Check("system databases nested under System Databases", () =>
        {
            var target = Object("Table", "Orders", "dbo");
            var host = new FakeHost { Root = Object("Server", "HOST", null,
                Folder("Databases", Folder("SystemDatabases", Object("Database", "master", null, Folder("UserTables", target))))) };
            Navigate(host, new ScriptObjectSelectionItem("USER_TABLE", "dbo", "Orders", 1, "master", 0, null));
            Equal(target, host.Selected);
        });
        foreach (bool schemasAboveType in new[] { false, true })
            Check("schema grouping " + schemasAboveType, () =>
            {
                var target = Object("Table", "Orders", "dbo");
                // A schema-grouped tree may omit both schema prefix and context.
                target.InvariantName = "Orders"; target.NavigationContext = null;
                var wrongSchema = Folder("other", Object("Table", "Orders", "other"));
                wrongSchema.Name = "other";
                var schema = Folder("dbo", schemasAboveType ? Folder("UserTables", target) : target);
                schema.Name = "dbo";
                var groups = Folder("Schemas", wrongSchema, schema);
                var host = Host(schemasAboveType ? groups : Folder("UserTables", groups));
                Navigate(host); Equal(target, host.Selected); Equal(0, wrongSchema.Loads);
            });
        Check("localized database folder with native collection metadata", () =>
        {
            var target = Object("Table", "Orders", "dbo");
            var host = Host(Folder("UserTables", target));
            var folder = (FakeNode)host.Root.Children[0];
            folder.Name = "Datenbanken"; folder.UniqueName = "version-specific"; folder.UrnPath = "Server/Database";
            Navigate(host); Equal(target, host.Selected);
        });
        Check("schema metadata node can contain grouped objects", () =>
        {
            var target = Object("Table", "Orders"); target.NavigationContext = null;
            var schema = Object("Schema", "dbo", null, Folder("UserTables", target));
            var host = Host(Folder("Schemas", schema));
            Navigate(host); Equal(target, host.Selected);
        });
        Check("schema-grouped procedures can omit Programmability", () =>
        {
            var target = Object("StoredProcedure", "Orders", "dbo");
            var host = Host(Folder("StoredProcedures", target));
            Navigate(host, Item("SQL_STORED_PROCEDURE")); Equal(target, host.Selected);
        });
        Check("versioned programmability internal name", () =>
        {
            var target = Object("StoredProcedure", "Orders", "dbo");
            var host = Host(Folder("Sql2005UserProgrammability", Folder("StoredProcedures", target)));
            Navigate(host, Item("SQL_STORED_PROCEDURE")); Equal(target, host.Selected);
        });
        foreach (string type in new[] { "SQL_SCALAR_FUNCTION", "CLR_SCALAR_FUNCTION", "SQL_TABLE_VALUED_FUNCTION", "SQL_INLINE_TABLE_VALUED_FUNCTION", "CLR_TABLE_VALUED_FUNCTION" })
            Check("function " + type, () =>
            {
                bool scalar = type.Contains("SCALAR");
                var target = Object("UserDefinedFunction", "Orders", "dbo");
                var wrong = Folder(scalar ? "Table-valuedFunctions" : "Scalar-valuedFunctions");
                var host = Host(Folder("UserProgrammability", Folder("UsrDbFunctions", wrong,
                    Folder(scalar ? "Scalar-valuedFunctions" : "Table-valuedFunctions", target))));
                Navigate(host, Item(type)); Equal(target, host.Selected); Equal(0, wrong.Loads);
            });
        foreach (string type in new[] { "VIEW", "SYNONYM", "TYPE_TABLE", "SEQUENCE_OBJECT", "CLR_AGGREGATE_FUNCTION", "SYSTEM_TABLE", "EXTERNAL_TABLE" })
            Check("object " + type, () =>
            {
                var step = ObjectExplorerPath.GetTreeSteps(Item(type)).Last();
                var target = Object(step.Type, "Orders", "dbo");
                FakeNode branch = target;
                string[] route;
                switch (type)
                {
                    case "VIEW": route = new[] { "Views" }; break;
                    case "SYNONYM": route = new[] { "Synonyms" }; break;
                    case "TYPE_TABLE": route = new[] { "UserProgrammability", "Types", "UserDefinedTableTypes" }; break;
                    case "SEQUENCE_OBJECT": route = new[] { "UserProgrammability", "Sequences" }; break;
                    case "CLR_AGGREGATE_FUNCTION": route = new[] { "UserProgrammability", "UsrDbFunctions", "AggregateFunctions" }; break;
                    case "SYSTEM_TABLE": route = new[] { "UserTables", "SystemTables" }; break;
                    default: route = new[] { "UserTables", "ExternalTables" }; break;
                }
                foreach (var folder in route.Reverse()) branch = Folder(folder, branch);
                var host = Host(branch); Navigate(host, Item(type)); Equal(target, host.Selected);
            });
        foreach (string type in new[] { "INDEX", "PRIMARY_KEY_CONSTRAINT", "UNIQUE_CONSTRAINT", "FOREIGN_KEY_CONSTRAINT", "CHECK_CONSTRAINT", "DEFAULT_CONSTRAINT", "SQL_TRIGGER", "CLR_TRIGGER" })
            foreach (string parentType in new[] { "USER_TABLE", "VIEW" })
                Check(type + " on " + parentType, () =>
                {
                    var item = Item(type, parentType, "child");
                    string urnType, folder;
                    switch (type)
                    {
                        case "PRIMARY_KEY_CONSTRAINT": urnType = "Index"; folder = "Keys"; break;
                        case "FOREIGN_KEY_CONSTRAINT": urnType = "ForeignKey"; folder = "Keys"; break;
                        case "CHECK_CONSTRAINT": urnType = "Check"; folder = "Constraints"; break;
                        case "DEFAULT_CONSTRAINT": urnType = "Column"; folder = "Columns"; break;
                        case "SQL_TRIGGER": case "CLR_TRIGGER": urnType = "Trigger"; folder = "Triggers"; break;
                        default: urnType = "Index"; folder = "Indexes"; break;
                    }
                    var target = Object(urnType, type == "DEFAULT_CONSTRAINT" ? "Status" : "child");
                    var parent = Object(parentType == "VIEW" ? "View" : "Table", "Orders", "dbo", Folder(folder, target));
                    var host = Host(Folder(parentType == "VIEW" ? "Views" : "UserTables", parent));
                    Navigate(host, item); Equal(target, host.Selected);
                });
        Check("exact case and schema selection", () =>
        {
            var target = Object("Table", "Orders", "dbo");
            var host = Host(Folder("UserTables", Object("Table", "orders", "dbo"), Object("Table", "Orders", "sales"), target));
            Navigate(host); Equal(target, host.Selected);
        });
        Check("dots in schema and name do not collide", () =>
        {
            var target = Object("Table", "c", "a.b");
            var host = Host(Folder("UserTables", Object("Table", "b.c", "a"), target));
            Navigate(host, new ScriptObjectSelectionItem("USER_TABLE", "a.b", "c", 1, "Sales", 0, null));
            Equal(target, host.Selected);
        });
        Check("quoted unusual names use exact metadata", () =>
        {
            const string name = "a'b]/Table[@Name='c / 雪", schema = "s'.x";
            var target = Object("Table", name, schema);
            var host = Host(Folder("UserTables", target));
            Navigate(host, new ScriptObjectSelectionItem("USER_TABLE", schema, name, 1, "Sales", 0, null));
            Equal(target, host.Selected);
        });
        Check("native reflection finds inherited item and context", () =>
        {
            var node = new NativeNode();
            Equal("UserTables", ObjectExplorerReflection.UniqueName(node));
            Equal(true, ObjectExplorerReflection.Read(ObjectExplorerReflection.Read(node, "containedItem"), "IsFolder"));
            ObjectExplorerReflection.EnumerateChildren(node);
            Equal(false, node.Async); Equal(1, node.Calls);
        });
        Check("missing native enumeration fails explicitly", () =>
            Throws<NotSupportedException>(() => ObjectExplorerReflection.EnumerateChildren(new object())));
        Check("native enumeration exception is unwrapped", () =>
            Throws<UnauthorizedAccessException>(() => ObjectExplorerReflection.EnumerateChildren(new FailingNativeNode())));
        Check("native Show and fallback detection", () =>
        {
            var service = new NativeService();
            True(ObjectExplorerReflection.TryShow(service)); Equal(1, service.Calls);
            True(!ObjectExplorerReflection.TryShow(new object()));
        });
    }

    private class NativeNodeBase
    {
#pragma warning disable CS0414
        private readonly object containedItem = new NativeItem();
#pragma warning restore CS0414
        public bool Async = true;
        public int Calls;
        private void EnumerateChildren(bool async) { Async = async; Calls++; }
    }
    private sealed class NativeNode : NativeNodeBase { }
    private class NativeItemBase
    {
#pragma warning disable CS0414
        private readonly object context = new NativeContext();
#pragma warning restore CS0414
        public bool IsFolder => true;
    }
    private sealed class NativeItem : NativeItemBase { }
    private sealed class NativeContext { public object this[string name] => name == "UniqueName" ? "UserTables" : null; }
    private sealed class FailingNativeNode { public void EnumerateChildren(bool async) => throw new UnauthorizedAccessException("Metadata denied."); }
    private sealed class NativeService { public int Calls; public void Show() { Calls++; } }
}
