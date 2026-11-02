using Microsoft.Data.SqlClient;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Interop;

namespace AxialSqlTools
{
    internal static class SqlObjectResolver
    {
        // Both commands use the same metadata query and platform-aware master fallback. ConfigureAwait(false)
        // also allows the existing synchronous scripting command to use this lookup.
        public static async Task<List<ScriptObjectSelectionItem>> FindObjectsAsync(
            string connectionString, string selectedName, CancellationToken cancellationToken)
        {
            var name = SqlObjectName.Parse(selectedName);
            var builder = new SqlConnectionStringBuilder(connectionString);
            // Azure SQL Database requires a new connection to change databases. Set the
            // catalog before opening, including for explicit database.schema.object names.
            if (name.DatabaseName != null)
                builder.InitialCatalog = name.DatabaseName;
            using (var connection = new SqlConnection(builder.ConnectionString))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                var matches = await QueryObjectsAsync(connection, name, cancellationToken).ConfigureAwait(false);
                if (!string.Equals(connection.Database, "master", StringComparison.OrdinalIgnoreCase)
                    && await SupportsMasterLookupAsync(connection, cancellationToken).ConfigureAwait(false))
                {
                    // Preserve the original credentials. Connection.ConnectionString can
                    // omit the password after Open when Persist Security Info is false.
                    builder.InitialCatalog = "master";
                    using (var master = new SqlConnection(builder.ConnectionString))
                    {
                        await master.OpenAsync(cancellationToken).ConfigureAwait(false);
                        matches.AddRange(await QueryObjectsAsync(master, name, cancellationToken).ConfigureAwait(false));
                    }
                }

                if (matches.Count == 0)
                    throw new InvalidOperationException($"The specified object was not found: '{selectedName}'.");
                return matches;
            }
        }

        public static ScriptObjectSelectionItem SelectObject(IList<ScriptObjectSelectionItem> matches, string action = "script")
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (matches.Count == 1)
                return matches[0];

            var dialog = new ScriptObjectPickerDialog(matches, action);
            var uiShell = Package.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
            if (uiShell != null && uiShell.GetDialogOwnerHwnd(out var hwnd) == 0 && hwnd != IntPtr.Zero)
                new WindowInteropHelper(dialog).Owner = hwnd;
            return dialog.ShowDialog() == true ? dialog.SelectedObject : null;
        }

        private static async Task<bool> SupportsMasterLookupAsync(SqlConnection connection, CancellationToken cancellationToken)
        {
            using (var command = new SqlCommand("SELECT CONVERT(int, SERVERPROPERTY('EngineEdition'));", connection))
            {
                var edition = await command.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false);
                // SQL Server, Managed Instance and SQL Edge share an instance-level master.
                // Database-scoped Azure services must not require access to virtual master.
                switch (Convert.ToInt32(edition))
                {
                    case 1: case 2: case 3: case 4: case 8: case 9:
                        return true;
                    default:
                        return false;
                }
            }
        }

        private static async Task<List<ScriptObjectSelectionItem>> QueryObjectsAsync(SqlConnection connection,
            SqlObjectName name, CancellationToken cancellationToken)
        {
            string sql = @"
SELECT o.type_desc, s.name, COALESCE(tt.name, o.name), o.object_id, DB_NAME(),
       o.parent_object_id, COALESCE(pt.name, p.name), p.type_desc, c.name
FROM sys.objects AS o
LEFT JOIN sys.table_types AS tt ON tt.type_table_object_id = o.object_id
JOIN sys.schemas AS s ON s.schema_id = COALESCE(tt.schema_id, o.schema_id)
LEFT JOIN sys.objects AS p ON p.object_id = o.parent_object_id
LEFT JOIN sys.table_types AS pt ON pt.type_table_object_id = p.object_id
LEFT JOIN sys.columns AS c ON c.object_id = o.parent_object_id AND c.default_object_id = o.object_id
WHERE COALESCE(tt.name, o.name) = @objectName AND (@schemaName IS NULL OR s.name = @schemaName)
UNION ALL
SELECT 'INDEX', s.name, i.name, o.object_id, DB_NAME(), o.object_id, COALESCE(tt.name, o.name), o.type_desc, NULL
FROM sys.indexes AS i
JOIN sys.objects AS o ON o.object_id = i.object_id
LEFT JOIN sys.table_types AS tt ON tt.type_table_object_id = o.object_id
JOIN sys.schemas AS s ON s.schema_id = COALESCE(tt.schema_id, o.schema_id)
WHERE i.name = @objectName AND (@schemaName IS NULL OR s.name = @schemaName);";

            using (var command = new SqlCommand(sql, connection))
            {
                command.Parameters.Add("@objectName", SqlDbType.NVarChar, 128).Value = name.ObjectName;
                command.Parameters.Add("@schemaName", SqlDbType.NVarChar, 128).Value = (object)name.SchemaName ?? DBNull.Value;
                var matches = new List<ScriptObjectSelectionItem>();
                using (var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false))
                {
                    while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                    {
                        matches.Add(new ScriptObjectSelectionItem(reader.GetString(0), reader.GetString(1),
                            reader.GetString(2), reader.GetInt32(3), reader.GetString(4), reader.GetInt32(5),
                            await reader.IsDBNullAsync(6, cancellationToken).ConfigureAwait(false) ? null : reader.GetString(6),
                            await reader.IsDBNullAsync(7, cancellationToken).ConfigureAwait(false) ? null : reader.GetString(7),
                            await reader.IsDBNullAsync(8, cancellationToken).ConfigureAwait(false) ? null : reader.GetString(8)));
                    }
                }
                return matches;
            }
        }
    }
}
