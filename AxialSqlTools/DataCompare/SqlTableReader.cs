using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace AxialSqlTools.DataCompare
{
    public static class SqlTableReader
    {
        public static List<TableSchema> ListTables(string connectionString, CancellationToken token)
        {
            using (var connection = Open(connectionString, token))
            using (var command = new SqlCommand(@"SELECT s.name, t.name FROM sys.tables t
JOIN sys.schemas s ON s.schema_id=t.schema_id WHERE t.is_ms_shipped=0 ORDER BY s.name,t.name;", connection))
            using (CancelWith(command, token))
            using (var reader = command.ExecuteReader())
            {
                var tables = new List<TableSchema>();
                while (reader.Read())
                {
                    token.ThrowIfCancellationRequested();
                    tables.Add(new TableSchema { Schema = reader.GetString(0), Name = reader.GetString(1) });
                }
                return tables;
            }
        }

        public static TableSchema LoadSchema(string connectionString, string schema, string table, CancellationToken token)
        {
            using (var connection = Open(connectionString, token))
            using (var command = new SqlCommand(@"
SELECT CONVERT(nvarchar(256),SERVERPROPERTY('ServerName')), DB_NAME(), t.temporal_type
FROM sys.tables t JOIN sys.schemas s ON s.schema_id=t.schema_id WHERE s.name=@schema AND t.name=@table;
SELECT c.name, TYPE_NAME(c.system_type_id), c.max_length, c.precision, c.scale,
 c.is_nullable, c.is_identity, c.is_computed, c.generated_always_type, c.encryption_type,
 c.default_object_id, c.collation_name
FROM sys.columns c JOIN sys.tables t ON t.object_id=c.object_id JOIN sys.schemas s ON s.schema_id=t.schema_id
WHERE s.name=@schema AND t.name=@table ORDER BY c.column_id;
SELECT i.name, i.is_primary_key, c.name
FROM sys.indexes i JOIN sys.tables t ON t.object_id=i.object_id JOIN sys.schemas s ON s.schema_id=t.schema_id
JOIN sys.index_columns ic ON ic.object_id=i.object_id AND ic.index_id=i.index_id
JOIN sys.columns c ON c.object_id=ic.object_id AND c.column_id=ic.column_id
WHERE s.name=@schema AND t.name=@table AND i.is_unique=1 AND i.is_disabled=0 AND i.has_filter=0
AND i.is_hypothetical=0 AND ic.key_ordinal>0 ORDER BY i.is_primary_key DESC,i.index_id,ic.key_ordinal;", connection))
            using (CancelWith(command, token))
            {
                command.Parameters.Add("@schema", SqlDbType.NVarChar, 128).Value = schema;
                command.Parameters.Add("@table", SqlDbType.NVarChar, 128).Value = table;
                using (var reader = command.ExecuteReader())
                {
                    if (!reader.Read()) throw new InvalidOperationException("The selected table no longer exists or is not visible to this login.");
                    var result = new TableSchema { Schema = schema, Name = table, Server = reader.GetString(0), Database = reader.GetString(1), HistoryTable = reader.GetByte(2) == 1 };
                    reader.NextResult();
                    while (reader.Read())
                    {
                        token.ThrowIfCancellationRequested();
                        result.Columns.Add(new TableColumn { Name = reader.GetString(0), Type = reader.IsDBNull(1) ? "unsupported" : reader.GetString(1), Length = reader.GetInt16(2),
                            Precision = reader.GetByte(3), Scale = reader.GetByte(4), Nullable = reader.GetBoolean(5), Identity = reader.GetBoolean(6), Computed = reader.GetBoolean(7),
                            Generated = reader.GetByte(8) != 0, Encrypted = !reader.IsDBNull(9), HasDefault = reader.GetInt32(10) != 0,
                            Collation = reader.IsDBNull(11) ? null : reader.GetString(11) });
                    }
                    reader.NextResult();
                    while (reader.Read())
                    {
                        token.ThrowIfCancellationRequested();
                        string name = reader.GetString(0);
                        var key = result.Keys.FirstOrDefault(k => k.Name == name);
                        if (key == null) { key = new TableKey { Name = name, Primary = reader.GetBoolean(1) }; result.Keys.Add(key); }
                        key.Columns.Add(reader.GetString(2));
                    }
                    return result;
                }
            }
        }

        public static IEnumerable<object[]> ReadRows(string connectionString, ComparisonPlan plan, bool source, CancellationToken token)
        {
            TableSchema table = source ? plan.Source : plan.Target;
            var columns = plan.Columns.Select(m => source ? m.Source : m.Target).ToArray();
            using (var connection = Open(connectionString, token))
            using (var transaction = plan.Options.UseSnapshotIsolation ? connection.BeginTransaction(IsolationLevel.Snapshot) : null)
            using (var command = new SqlCommand("SELECT " + string.Join(",", columns.Select(c => c.Type == "xml" ? "CONVERT(nvarchar(max), " + SqlText.Quote(c.Name) + ")" : SqlText.Quote(c.Name))) + " FROM " + table.QualifiedName + ";", connection, transaction))
            using (CancelWith(command, token))
            {
                command.CommandTimeout = plan.Options.CommandTimeoutSeconds;
                // SequentialAccess bounds large value reads. No checksum/hash approximation or NOLOCK.
                using (var reader = command.ExecuteReader(CommandBehavior.SequentialAccess))
                    while (reader.Read())
                    {
                        token.ThrowIfCancellationRequested();
                        var row = new object[columns.Length];
                        long rowBytes = 64 + columns.Length * 32L;
                        for (int i = 0; i < columns.Length; i++)
                        {
                            token.ThrowIfCancellationRequested();
                            if (reader.IsDBNull(i)) continue;
                            string type = columns[i].Type;
                            if (SqlText.IsText(type))
                            {
                                using (var textReader = reader.GetTextReader(i))
                                {
                                    var text = new StringBuilder(); var buffer = new char[8192]; int length;
                                    while ((length = textReader.Read(buffer, 0, buffer.Length)) > 0)
                                    {
                                        token.ThrowIfCancellationRequested();
                                        if ((text.Length + (long)length) * 2 > plan.Options.MaxCellBytes) throw CellLimit(columns[i]);
                                        text.Append(buffer, 0, length);
                                    }
                                    row[i] = text.ToString();
                                }
                            }
                            else if (SqlText.IsBinary(type))
                            {
                                using (var stream = reader.GetStream(i))
                                using (var bytes = new MemoryStream())
                                {
                                    var buffer = new byte[8192]; int length;
                                    while ((length = stream.Read(buffer, 0, buffer.Length)) > 0)
                                    {
                                        token.ThrowIfCancellationRequested();
                                        if (bytes.Length + length > plan.Options.MaxCellBytes) throw CellLimit(columns[i]);
                                        bytes.Write(buffer, 0, length);
                                    }
                                    row[i] = bytes.ToArray();
                                }
                            }
                            else if (type == "decimal" || type == "numeric") row[i] = reader.GetSqlDecimal(i);
                            else row[i] = reader.GetValue(i);
                            if (row[i] is string textValue) rowBytes += textValue.Length * 2L + 32;
                            if (row[i] is byte[] binaryValue) rowBytes += binaryValue.Length + 32;
                            if (rowBytes > 32L * 1024 * 1024)
                                throw new InvalidOperationException("A comparison row exceeds 32 MiB. Exclude large columns and compare again.");
                        }
                        yield return row;
                    }
                transaction?.Commit();
            }
        }

        private static InvalidOperationException CellLimit(TableColumn column) => new InvalidOperationException("Column " + column.Name + " contains a value larger than 16 MiB. Exclude it and compare again.");

        internal static SqlConnection Open(string connectionString, CancellationToken token)
        {
            var connection = new SqlConnection(connectionString);
            try { connection.OpenAsync(token).GetAwaiter().GetResult(); return connection; }
            catch { connection.Dispose(); throw; }
        }

        internal static CancellationTokenRegistration CancelWith(SqlCommand command, CancellationToken token) => token.Register(() =>
        {
            try { command.Cancel(); }
            catch (InvalidOperationException) { }
            catch (SqlException) { }
        });
    }
}
