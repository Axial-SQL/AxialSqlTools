using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

// These fakes exercise the production orchestration without a server or SSMS.
// They deliberately reject USE and strip passwords from open connections.
internal static class FakeSql
{
    public static int EngineEdition;
    public static int ScalarCalls;
    public static Action<string> BeforeOperation;
    public static readonly List<Microsoft.Data.SqlClient.SqlConnection> Connections = new List<Microsoft.Data.SqlClient.SqlConnection>();
    public static readonly List<Microsoft.Data.SqlClient.SqlCommand> Commands = new List<Microsoft.Data.SqlClient.SqlCommand>();
    public static readonly List<Microsoft.Data.SqlClient.SqlCommand> MetadataCalls = new List<Microsoft.Data.SqlClient.SqlCommand>();
    public static readonly List<Microsoft.Data.SqlClient.SqlDataReader> Readers = new List<Microsoft.Data.SqlClient.SqlDataReader>();
    public static readonly List<object[]> Rows = new List<object[]>();

    public static void Reset()
    {
        EngineEdition = 5;
        ScalarCalls = 0;
        BeforeOperation = null;
        Connections.Clear();
        Commands.Clear();
        MetadataCalls.Clear();
        Readers.Clear();
        Rows.Clear();
    }

    public static void Add(string database, string schema = "dbo", string name = "Orders")
        => Rows.Add(new object[] { "USER_TABLE", schema, name, 42, database, 0, null, null, null });

    public static async Task Operation(string stage, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        await Task.Yield();
        BeforeOperation?.Invoke(stage);
        token.ThrowIfCancellationRequested();
    }
}

namespace Microsoft.Data.SqlClient
{
    internal sealed class SqlConnectionStringBuilder
    {
        private readonly DbConnectionStringBuilder builder = new DbConnectionStringBuilder();
        public SqlConnectionStringBuilder(string text) { builder.ConnectionString = text; }
        public string InitialCatalog
        {
            get => builder.TryGetValue("Initial Catalog", out var value) ? (string)value : "master";
            set => builder["Initial Catalog"] = value;
        }
        public string ConnectionString => builder.ConnectionString;
    }

    internal sealed class SqlConnection : IDisposable
    {
        public SqlConnection(string text)
        {
            OriginalConnectionString = text;
            Database = new SqlConnectionStringBuilder(text).InitialCatalog;
            FakeSql.Connections.Add(this);
        }
        public string OriginalConnectionString { get; }
        public string Database { get; }
        public string OpenedDatabase { get; private set; }
        public bool Opened { get; private set; }
        public bool Disposed { get; private set; }
        public string ConnectionString
        {
            get
            {
                var builder = new DbConnectionStringBuilder { ConnectionString = OriginalConnectionString };
                if (Opened) builder.Remove("Password");
                return builder.ConnectionString;
            }
        }
        public async Task OpenAsync(CancellationToken token)
        {
            await FakeSql.Operation("open:" + Database, token);
            if (FakeSql.EngineEdition == 5 && Database == "master")
                throw new InvalidOperationException("Contained user cannot access virtual master");
            Opened = true;
            OpenedDatabase = Database;
        }
        public void Dispose() { Disposed = true; }
    }

    internal sealed class SqlParameter
    {
        public object Value { get; set; }
        public SqlDbType Type { get; set; }
        public int Size { get; set; }
    }

    internal sealed class SqlParameterCollection
    {
        private readonly Dictionary<string, SqlParameter> parameters = new Dictionary<string, SqlParameter>();
        public SqlParameter Add(string name, SqlDbType type, int size)
        {
            var parameter = new SqlParameter { Type = type, Size = size };
            parameters.Add(name, parameter);
            return parameter;
        }
        public SqlParameter this[string name] => parameters[name];
    }

    internal sealed class SqlCommand : IDisposable
    {
        private readonly SqlConnection connection;
        public SqlCommand(string text, SqlConnection connection)
        {
            CommandText = text;
            this.connection = connection;
            FakeSql.Commands.Add(this);
        }
        public string CommandText { get; }
        public SqlParameterCollection Parameters { get; } = new SqlParameterCollection();
        public bool Disposed { get; private set; }

        public async Task<object> ExecuteScalarAsync(CancellationToken token)
        {
            RequireOpen();
            if (!CommandText.Contains("SERVERPROPERTY('EngineEdition')"))
                throw new InvalidOperationException("Unexpected scalar query");
            FakeSql.ScalarCalls++;
            await FakeSql.Operation("edition:" + connection.Database, token);
            return FakeSql.EngineEdition;
        }

        public async Task<SqlDataReader> ExecuteReaderAsync(CancellationToken token)
        {
            RequireOpen();
            if (CommandText.TrimStart().StartsWith("USE ", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("USE is not supported by this test server");
            FakeSql.MetadataCalls.Add(this);
            await FakeSql.Operation("query:" + connection.Database, token);
            var name = (string)Parameters["@objectName"].Value;
            var schema = Parameters["@schemaName"].Value;
            var rows = FakeSql.Rows.Where(row => (string)row[4] == connection.Database
                && (string)row[2] == name && (schema == DBNull.Value || (string)row[1] == (string)schema)).ToList();
            return new SqlDataReader(rows, connection.Database);
        }

        private void RequireOpen()
        {
            if (!connection.Opened || connection.Disposed) throw new InvalidOperationException("Connection is not open");
        }
        public void Dispose() { Disposed = true; }
    }

    internal sealed class SqlDataReader : IDisposable
    {
        private readonly List<object[]> rows;
        private readonly string database;
        private int position = -1;
        public SqlDataReader(List<object[]> rows, string database)
        {
            this.rows = rows;
            this.database = database;
            FakeSql.Readers.Add(this);
        }
        public bool Disposed { get; private set; }
        public async Task<bool> ReadAsync(CancellationToken token)
        {
            await FakeSql.Operation("read:" + database, token);
            return ++position < rows.Count;
        }
        public string GetString(int column) => (string)rows[position][column];
        public int GetInt32(int column) => (int)rows[position][column];
        public async Task<bool> IsDBNullAsync(int column, CancellationToken token)
        {
            await FakeSql.Operation("null:" + database, token);
            return rows[position][column] == null;
        }
        public void Dispose() { Disposed = true; }
    }
}

namespace AxialSqlTools
{
    // Name parsing itself depends on ScriptDom and is outside this harness.
    internal sealed class SqlObjectName
    {
        public static SqlObjectName Next;
        public static string LastText;
        public SqlObjectName(string database, string schema, string name)
        {
            DatabaseName = database;
            SchemaName = schema;
            ObjectName = name;
        }
        public string DatabaseName { get; }
        public string SchemaName { get; }
        public string ObjectName { get; }
        public static SqlObjectName Parse(string text) { LastText = text; return Next; }
        internal static string Quote(string name) => "[" + name.Replace("]", "]]") + "]";
    }
    internal sealed class ScriptObjectPickerDialog
    {
        public ScriptObjectPickerDialog(IList<ScriptObjectSelectionItem> items, string action) { }
        public ScriptObjectSelectionItem SelectedObject => throw new NotSupportedException();
        public bool? ShowDialog() => throw new NotSupportedException();
    }
}

namespace Microsoft.VisualStudio.Shell
{
    internal static class ThreadHelper { public static void ThrowIfNotOnUIThread() { } }
    internal static class Package { public static object GetGlobalService(Type type) => null; }
}
namespace Microsoft.VisualStudio.Shell.Interop
{
    internal sealed class SVsUIShell { }
    internal interface IVsUIShell { int GetDialogOwnerHwnd(out IntPtr hwnd); }
}
namespace System.Windows.Interop
{
    internal sealed class WindowInteropHelper
    {
        public WindowInteropHelper(object dialog) { }
        public IntPtr Owner { get; set; }
    }
}
