using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Sdk.Sfc;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Interop;
using static AxialSqlTools.ScriptFactoryAccess;

namespace AxialSqlTools
{
    public sealed class ScriptObjectSelectionItem
    {
        public ScriptObjectSelectionItem(string typeDesc, string schemaName, string objectName, int objectId, string databaseName, int parentObjectId, string parentObjectName)
        {
            TypeDesc = typeDesc;
            SchemaName = schemaName;
            ObjectName = objectName;
            ObjectId = objectId;
            DatabaseName = databaseName;
            ParentObjectId = parentObjectId;
            ParentObjectName = parentObjectName;
        }

        public string TypeDesc { get; }

        public string SchemaName { get; }

        public string ObjectName { get; }

        public int ObjectId { get; }

        public string DatabaseName { get; }

        public int ParentObjectId { get; }

        public string ParentObjectName { get; }

        public string DisplayName =>
            string.IsNullOrEmpty(ParentObjectName)
                ? $"{DatabaseName}.{SchemaName}.{ObjectName} ({TypeDesc})"
                : $"{DatabaseName}.{SchemaName}.{ParentObjectName}.{ObjectName} ({TypeDesc})";
    }

    public static class ScriptObjectDefinition
    {

        public static string GetText(AsyncPackage package, string selectedObjectName)
        {
            var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfo();

            ScriptObjectSelectionItem selectedObject = null;

            using (SqlConnection currentServerConnection = new SqlConnection(connectionInfo.FullConnectionString))
            {
                currentServerConnection.Open();

                string commandText = $@"
                        SELECT PARSENAME(@FullName, 3) AS DatabaseName,
                               PARSENAME(@FullName, 2) AS SchemaName,
                               PARSENAME(@FullName, 1) AS ObjectName;
                        ";

                string currentDatabase = string.Empty;

                ParsedObjectName parsedObjectName = null;

                using (SqlCommand cmd = new SqlCommand(commandText, currentServerConnection))
                {
                    cmd.Parameters.Add(new SqlParameter("FullName", selectedObjectName));

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            currentDatabase = reader.IsDBNull(0) ? currentServerConnection.Database : reader.GetString(0);
                            parsedObjectName = new ParsedObjectName(
                                reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                                reader.IsDBNull(2) ? string.Empty : reader.GetString(2)
                            );
                        }
                    }
                }

                if (string.IsNullOrEmpty(parsedObjectName.ObjectName))
                {
                    throw new Exception($"Failed to extract a table name from the provided object string: {selectedObjectName}");
                }

                List<ScriptObjectSelectionItem> matches = QueryObjects(
                    currentServerConnection,
                    currentDatabase,
                    parsedObjectName.ObjectName,
                    parsedObjectName.SchemaName);

                if (!string.Equals(currentDatabase, "master", StringComparison.OrdinalIgnoreCase))
                {
                    matches.AddRange(QueryObjects(
                        currentServerConnection,
                        "master",
                        parsedObjectName.ObjectName,
                        parsedObjectName.SchemaName));
                }

                matches = DeduplicateMatches(matches);

                if (matches.Count == 0)
                {
                    throw new Exception($"The specified object was not found: '{selectedObjectName}'.");
                }

                if (matches.Count == 1)
                {
                    selectedObject = matches[0];
                }
                else
                {
                    var dialog = new ScriptObjectPickerDialog(matches);
                    var uiShell = Package.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
                    if (uiShell != null && uiShell.GetDialogOwnerHwnd(out var hwnd) == 0 && hwnd != IntPtr.Zero)
                    {
                        new WindowInteropHelper(dialog).Owner = hwnd;
                    }

                    bool? result = dialog.ShowDialog();
                    if (result != true)
                    {
                        return string.Empty;
                    }

                    selectedObject = dialog.SelectedObject;
                    if (selectedObject == null)
                    {
                        return string.Empty;
                    }
                }
            }

            ServerConnection SmoConnection = new ServerConnection();
            SmoConnection.ConnectionString = connectionInfo.FullConnectionString;
            Server server = new Server(SmoConnection);

            Scripter scripter = new Scripter(server) { Options = new ScriptingOptions() };

            if (selectedObject.TypeDesc == "USER_TABLE")
            {
                scripter.Options.ScriptData = false;
                scripter.Options.DriAllKeys = true;

                scripter.Options.Indexes = true;
                scripter.Options.Triggers = true;
                scripter.Options.Default = true;
                scripter.Options.DriAll = true;

                scripter.Options.ScriptDataCompression = true;
                scripter.Options.NoCollation = true;
            }
            else
            {
                scripter.Options.ScriptForCreateOrAlter = true;
                scripter.Options.EnforceScriptingOptions = true;
            }
            // scripter.Options.ScriptBatchTerminator = true; -> this doesn't work for some reason..

            Database db = server.Databases[selectedObject.DatabaseName];
            SqlSmoObject dbObject = null;

            if (db.Tables.Contains(selectedObject.ObjectName, selectedObject.SchemaName))
            {
                dbObject = db.Tables[selectedObject.ObjectName, selectedObject.SchemaName];
            }
            else if (db.StoredProcedures.Contains(selectedObject.ObjectName, selectedObject.SchemaName))
            {
                dbObject = db.StoredProcedures[selectedObject.ObjectName, selectedObject.SchemaName];
            }
            else if (db.UserDefinedFunctions.Contains(selectedObject.ObjectName, selectedObject.SchemaName))
            {
                dbObject = db.UserDefinedFunctions[selectedObject.ObjectName, selectedObject.SchemaName];
            }
            else if (db.Views.Contains(selectedObject.ObjectName, selectedObject.SchemaName))
            {
                dbObject = db.Views[selectedObject.ObjectName, selectedObject.SchemaName];
            }
            else if (db.Synonyms.Contains(selectedObject.ObjectName, selectedObject.SchemaName))
            {
                dbObject = db.Synonyms[selectedObject.ObjectName, selectedObject.SchemaName];
            }
            else if (db.UserDefinedTableTypes.Contains(selectedObject.ObjectName, selectedObject.SchemaName))
            {
                dbObject = db.UserDefinedTableTypes[selectedObject.ObjectName, selectedObject.SchemaName];
            }
            else if (db.UserDefinedTypes.Contains(selectedObject.ObjectName, selectedObject.SchemaName))
            {
                dbObject = db.UserDefinedTypes[selectedObject.ObjectName, selectedObject.SchemaName];
            }
            else if (selectedObject.TypeDesc == "SQL_TRIGGER")
            {
                dbObject = FindTableTrigger(db, selectedObject.ParentObjectName, selectedObject.ObjectName, selectedObject.SchemaName);
            }
            else if (selectedObject.TypeDesc == "INDEX")
            {
                dbObject = FindTableIndex(db, selectedObject.ParentObjectName, selectedObject.ObjectName, selectedObject.SchemaName);
            }
            else if (selectedObject.TypeDesc == "PRIMARY_KEY_CONSTRAINT" || selectedObject.TypeDesc == "UNIQUE_CONSTRAINT")
            {
                dbObject = FindTableIndex(db, selectedObject.ParentObjectName, selectedObject.ObjectName, selectedObject.SchemaName);
            }
            else if (selectedObject.TypeDesc == "FOREIGN_KEY_CONSTRAINT")
            {
                dbObject = FindTableForeignKey(db, selectedObject.ParentObjectName, selectedObject.ObjectName, selectedObject.SchemaName);
            }
            else if (selectedObject.TypeDesc == "CHECK_CONSTRAINT")
            {
                dbObject = FindTableCheck(db, selectedObject.ParentObjectName, selectedObject.ObjectName, selectedObject.SchemaName);
            }
            else if (selectedObject.TypeDesc == "DEFAULT_CONSTRAINT")
            {
                dbObject = FindDefaultConstraint(db, selectedObject.ParentObjectName, selectedObject.ObjectName, selectedObject.SchemaName);
            }

            string fullScriptResult = String.Empty;

            if (dbObject != null)
            {
                System.Collections.Specialized.StringCollection sc = scripter.Script(new Urn[] { dbObject.Urn });

                StringBuilder sb = new StringBuilder();
                foreach (string line in sc)
                {
                    sb.AppendLine(line);
                    sb.AppendLine("GO");
                }
                fullScriptResult = sb.ToString();

                // additional format to make it pretty
                if (selectedObject.TypeDesc == "USER_TABLE")
                {
                    TSql170Parser sqlParser = new TSql170Parser(false);
                    IList<ParseError> parseErrors = new List<ParseError>();
                    TSqlFragment result = sqlParser.Parse(new StringReader(fullScriptResult), out parseErrors);

                    // leave it as is if for some reason we can't format it
                    if (parseErrors.Count == 0)
                    {
                        Sql170ScriptGenerator gen = new Sql170ScriptGenerator();
                        gen.Options.AlignClauseBodies = false;
                        gen.Options.IncludeSemicolons = false;
                        gen.GenerateScript(result, out fullScriptResult);
                    }
                }

            }
            else
            {
                throw new Exception($"The specified object was not found: '{selectedObjectName}'.");
            }

            return fullScriptResult;
        }



        private static List<ScriptObjectSelectionItem> QueryObjects(
            SqlConnection connection,
            string databaseName,
            string objectName,
            string schemaName)
        {
            string commandText = $@"USE [{databaseName}];
            SELECT o.type_desc,
                   s.name,
                   o.name,
                   o.object_id,
                   DB_NAME(),
                   parent_object_id,
                   OBJECT_NAME(parent_object_id) as parent_object_name      
            FROM sys.objects o
            JOIN sys.schemas s ON o.schema_id = s.schema_id
            WHERE o.name = @objectName
              AND (@schemaName IS NULL OR s.name = @schemaName)";
            commandText += @"
            UNION ALL
            SELECT 'INDEX' AS type_desc,
                   s.name,
                   i.name,
                   o.object_id,
                   DB_NAME(),
                   i.object_id,
                   OBJECT_NAME(i.object_id) 
            FROM sys.indexes i
            JOIN sys.objects o ON i.object_id = o.object_id
            JOIN sys.schemas s ON o.schema_id = s.schema_id
            WHERE i.name = @objectName
              AND (@schemaName IS NULL OR s.name = @schemaName);";

            using (SqlCommand cmd = new SqlCommand(commandText, connection))
            {
                cmd.Parameters.Add(new SqlParameter("objectName", objectName));
                cmd.Parameters.Add(new SqlParameter("schemaName", string.IsNullOrWhiteSpace(schemaName) ? (object)DBNull.Value : schemaName));

                List<ScriptObjectSelectionItem> matches = new List<ScriptObjectSelectionItem>();
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        matches.Add(new ScriptObjectSelectionItem(
                           reader.GetString(0),
                           reader.GetString(1),
                           reader.GetString(2),
                           reader.GetInt32(3),
                           reader.GetString(4),
                           reader.IsDBNull(5) ? 0 : reader.GetInt32(5),
                           reader.IsDBNull(6) ? string.Empty : reader.GetString(6)
                       ));
                    }
                }

                return matches;
            }
        }

        private static List<ScriptObjectSelectionItem> DeduplicateMatches(List<ScriptObjectSelectionItem> matches)
        {
            Dictionary<string, ScriptObjectSelectionItem> unique = new Dictionary<string, ScriptObjectSelectionItem>(StringComparer.OrdinalIgnoreCase);

            foreach (ScriptObjectSelectionItem match in matches)
            {
                string key = $"{match.DatabaseName}|{match.SchemaName}|{match.ObjectName}|{match.TypeDesc}|{match.ObjectId}";
                if (!unique.ContainsKey(key))
                {
                    unique[key] = match;
                }
            }

            return new List<ScriptObjectSelectionItem>(unique.Values);
        }

        private static SqlSmoObject FindTableTrigger(Database db, string ParentObjectName, string triggerName, string schemaName)
        {

            Table table = db.Tables[ParentObjectName, schemaName];
            if (table.Triggers.Contains(triggerName))
            {
                return table.Triggers[triggerName];
            }

            return null;
        }

        private static SqlSmoObject FindTableIndex(Database db, string ParentObjectName, string indexName, string schemaName)
        {
            Table table = db.Tables[ParentObjectName, schemaName];

            if (table.Indexes.Contains(indexName))
            {
                return table.Indexes[indexName];
            }


            return null;
        }

        private static SqlSmoObject FindTableForeignKey(Database db, string ParentObjectName, string constraintName, string schemaName)
        {
            Table table = db.Tables[ParentObjectName, schemaName];
            if (table.ForeignKeys.Contains(constraintName))
            {
                return table.ForeignKeys[constraintName];
            }

            return null;
        }

        private static SqlSmoObject FindTableCheck(Database db, string ParentObjectName, string constraintName, string schemaName)
        {
            Table table = db.Tables[ParentObjectName, schemaName];

            if (table.Checks.Contains(constraintName))
            {
                return table.Checks[constraintName];
            }

            return null;
        }

        private static SqlSmoObject FindDefaultConstraint(Database db, string ParentObjectName, string constraintName, string schemaName)
        {
            Table table = db.Tables[ParentObjectName, schemaName];

            foreach (Column column in table.Columns)
            {
                if (column.DefaultConstraint != null
                    && string.Equals(column.DefaultConstraint.Name, constraintName, StringComparison.OrdinalIgnoreCase))
                {
                    return column.DefaultConstraint;
                }
            }

            return null;
        }

        private sealed class ParsedObjectName
        {
            public ParsedObjectName(string schemaName, string objectName)
            {
                SchemaName = schemaName;
                ObjectName = objectName ?? throw new ArgumentNullException(nameof(objectName));
            }

            public string SchemaName { get; }

            public string ObjectName { get; }
        }

    }


}
