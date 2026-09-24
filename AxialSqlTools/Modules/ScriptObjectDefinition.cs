using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Sdk.Sfc;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace AxialSqlTools
{
    public static class ScriptObjectDefinition
    {

        public static string GetText(AsyncPackage package, string selectedObjectName)
        {
            var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfo();

            if (connectionInfo == null)
                throw new InvalidOperationException("Connect the SQL query window to a server before scripting an object.");

            var matches = SqlObjectResolver.FindObjectsAsync(connectionInfo.FullConnectionString,
                selectedObjectName, System.Threading.CancellationToken.None).GetAwaiter().GetResult();
            var selectedObject = SqlObjectResolver.SelectObject(matches);
            if (selectedObject == null)
                return string.Empty;

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

    }
}
