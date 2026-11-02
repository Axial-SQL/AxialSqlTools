using AxialSqlTools;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

internal static class SqlObjectResolverTests
{
    private const string Connection = "Server=test;Initial Catalog=App;User ID=test;Password=secret;Encrypt=True";
    private static int passed;

    private static int Main()
    {
        try
        {
            Run("Azure current database does not require master", AzureCurrent);
            Run("Azure missing object reports not found without master", AzureMissing);
            Run("Database-scoped and unknown engines do not query master", DatabaseScopedEngines);
            Run("Explicit database is selected before opening", ExplicitDatabase);
            Run("SQL Server keeps current and master matches", SqlServerAmbiguity);
            Run("SQL Server finds master-only objects", MasterOnly);
            Run("Managed Instance retains master fallback", ManagedInstance);
            Run("Already connected to master performs one lookup", CurrentMaster);
            Run("Object and schema values remain parameterized", Parameters);
            Run("Unqualified schema remains SQL NULL", UnqualifiedSchema);
            Run("Cancellation before opening propagates", PreCancelled);
            Run("Cancellation during metadata read propagates", ReadCancelled);
            Run("Cancellation during master open propagates", MasterCancelled);
            Run("Metadata errors propagate unchanged", MetadataFailure);
            Console.WriteLine("Passed " + passed + " object resolver tests.");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 1;
        }
    }

    private static void Run(string label, Action test)
    {
        FakeSql.Reset();
        SqlObjectName.Next = new SqlObjectName(null, "dbo", "Orders");
        test();
        Assert(FakeSql.Connections.All(c => c.Disposed), "Every connection must be disposed");
        Assert(FakeSql.Commands.All(c => c.Disposed), "Every command must be disposed");
        Assert(FakeSql.Readers.All(r => r.Disposed), "Every reader must be disposed");
        passed++;
        Console.WriteLine("PASS: " + label);
    }

    private static List<ScriptObjectSelectionItem> Find(string text = "dbo.Orders", CancellationToken token = default(CancellationToken))
        => SqlObjectResolver.FindObjectsAsync(Connection, text, token).GetAwaiter().GetResult();

    private static void AzureCurrent()
    {
        FakeSql.EngineEdition = 5;
        FakeSql.Add("App");
        Equal("App", Find().Single().DatabaseName);
        Equal("App", string.Join(",", FakeSql.Connections.Select(c => c.Database)));
        Equal(1, FakeSql.MetadataCalls.Count);
    }

    private static void AzureMissing()
    {
        FakeSql.EngineEdition = 5;
        var error = Throws<InvalidOperationException>(() => Find());
        Assert(error.Message.Contains("not found") && error.Message.Contains("dbo.Orders"), "Expected object-not-found detail");
        Equal(1, FakeSql.Connections.Count);
    }

    private static void DatabaseScopedEngines()
    {
        foreach (int edition in new[] { 6, 11, 12, 999 })
        {
            FakeSql.EngineEdition = edition;
            FakeSql.Rows.Clear();
            FakeSql.Add("App");
            int before = FakeSql.Connections.Count;
            Equal(1, Find().Count);
            Equal(before + 1, FakeSql.Connections.Count);
        }
    }

    private static void ExplicitDatabase()
    {
        const string database = "Sales];Archive DB";
        FakeSql.EngineEdition = 5;
        FakeSql.Add(database);
        SqlObjectName.Next = new SqlObjectName(database, "dbo", "Orders");
        Equal(database, Find("[Sales]];Archive DB].[dbo].[Orders]").Single().DatabaseName);
        Equal(database, FakeSql.Connections.Single().OpenedDatabase);
        Equal("[Sales]];Archive DB].[dbo].[Orders]", SqlObjectName.LastText);
    }

    private static void SqlServerAmbiguity()
    {
        FakeSql.EngineEdition = 3;
        FakeSql.Add("App");
        FakeSql.Add("master");
        Equal("App,master", string.Join(",", Find().Select(m => m.DatabaseName)));
        Equal("App,master", string.Join(",", FakeSql.Connections.Select(c => c.OpenedDatabase)));
        Assert(FakeSql.Connections.All(c => c.OriginalConnectionString.IndexOf("password=secret", StringComparison.OrdinalIgnoreCase) >= 0), "Credentials must survive the master connection");
        Assert(FakeSql.Connections.All(c => c.OriginalConnectionString.IndexOf("encrypt=True", StringComparison.OrdinalIgnoreCase) >= 0), "Other options must survive the master connection");
    }

    private static void MasterOnly()
    {
        FakeSql.EngineEdition = 3;
        FakeSql.Add("master");
        Equal("master", Find().Single().DatabaseName);
    }

    private static void ManagedInstance()
    {
        FakeSql.EngineEdition = 8;
        FakeSql.Add("master");
        Equal("master", Find().Single().DatabaseName);
    }

    private static void CurrentMaster()
    {
        FakeSql.EngineEdition = 3;
        FakeSql.Add("master");
        SqlObjectName.Next = new SqlObjectName("master", "dbo", "Orders");
        Equal("master", Find("master.dbo.Orders").Single().DatabaseName);
        Equal(1, FakeSql.MetadataCalls.Count);
        Equal(1, FakeSql.Connections.Count);
        Equal(0, FakeSql.ScalarCalls);
    }

    private static void Parameters()
    {
        const string name = "Order'; DROP TABLE x;--";
        const string schema = "Odd]Schema";
        FakeSql.EngineEdition = 5;
        FakeSql.Add("App", schema, name);
        SqlObjectName.Next = new SqlObjectName(null, schema, name);
        Equal(name, Find("[Odd]]Schema].[Order'; DROP TABLE x;--]").Single().ObjectName);
        var command = FakeSql.MetadataCalls.Single();
        Assert(!command.CommandText.Contains(name) && !command.CommandText.Contains(schema), "Names must not be interpolated into SQL");
        Equal(name, command.Parameters["@objectName"].Value);
        Equal(schema, command.Parameters["@schemaName"].Value);
        Equal(SqlDbType.NVarChar, command.Parameters["@objectName"].Type);
        Equal(128, command.Parameters["@objectName"].Size);
        Equal(128, command.Parameters["@schemaName"].Size);
        Equal(2, command.CommandText.Split(new[] { "@schemaName IS NULL OR s.name = @schemaName" }, StringSplitOptions.None).Length - 1);
    }

    private static void UnqualifiedSchema()
    {
        FakeSql.EngineEdition = 5;
        FakeSql.Add("App", "sales");
        SqlObjectName.Next = new SqlObjectName(null, null, "Orders");
        Equal("sales", Find("Orders").Single().SchemaName);
        Assert(FakeSql.MetadataCalls.Single().Parameters["@schemaName"].Value == DBNull.Value, "Absent schema must use SQL NULL");
    }

    private static void PreCancelled()
    {
        using (var cancel = new CancellationTokenSource())
        {
            cancel.Cancel();
            Throws<OperationCanceledException>(() => Find(token: cancel.Token));
            Equal(0, FakeSql.Commands.Count);
        }
    }

    private static void ReadCancelled()
    {
        using (var cancel = new CancellationTokenSource())
        {
            FakeSql.Add("App");
            FakeSql.BeforeOperation = stage => { if (stage == "read:App") cancel.Cancel(); };
            Throws<OperationCanceledException>(() => Find(token: cancel.Token));
            Equal(1, FakeSql.Connections.Count);
        }
    }

    private static void MasterCancelled()
    {
        using (var cancel = new CancellationTokenSource())
        {
            FakeSql.EngineEdition = 3;
            FakeSql.BeforeOperation = stage => { if (stage == "open:master") cancel.Cancel(); };
            Throws<OperationCanceledException>(() => Find(token: cancel.Token));
            Equal(2, FakeSql.Connections.Count);
        }
    }

    private static void MetadataFailure()
    {
        var expected = new InvalidOperationException("Metadata permission denied");
        FakeSql.BeforeOperation = stage => { if (stage == "query:App") throw expected; };
        var actual = Throws<InvalidOperationException>(() => Find());
        Assert(ReferenceEquals(expected, actual), "Do not mask metadata errors");
        Equal(1, FakeSql.Connections.Count);
    }

    private static T Throws<T>(Action action) where T : Exception
    {
        try { action(); }
        catch (T error) { return error; }
        throw new Exception("Expected " + typeof(T).Name);
    }

    private static void Equal<T>(T expected, T actual)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new Exception("Expected " + expected + ", got " + actual);
    }

    private static void Assert(bool value, string message)
    {
        if (!value) throw new Exception(message);
    }
}
