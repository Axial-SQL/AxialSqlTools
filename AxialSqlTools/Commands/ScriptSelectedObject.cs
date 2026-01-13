using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Interop;
using EnvDTE;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Sdk.Sfc;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ScriptSelectedObject
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4134;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScriptSelectedObject"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ScriptSelectedObject(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ScriptSelectedObject Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ScriptSelectedObject's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new ScriptSelectedObject(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE dte = Package.GetGlobalService(typeof(DTE)) as DTE;

            if (dte?.ActiveDocument != null)
            {

                try
                {
                    TextSelection selection = dte.ActiveDocument.Selection as TextSelection;

                    string selectedObjectName = selection.Text.Trim().TrimEnd(';');

                    if (string.IsNullOrEmpty(selectedObjectName))
                    {
                        throw new Exception("Nothing has been selected");
                    }

                    //Could be:
                    // - Table
                    // - Sproc
                    // - Function

                    // find all object with this name
                    // script properly from a current connection
                    // display in a single window

                    var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfo();

                    ScriptObjectSelectionItem selectedObject = null;
                    ParsedObjectName parsedObjectName = ParseSelectedObjectName(selectedObjectName);

                    using (SqlConnection currentServerConnection = new SqlConnection(connectionInfo.FullConnectionString))
                    {
                        currentServerConnection.Open();
                        string currentDatabase = currentServerConnection.Database;

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
                            throw new Exception("Unable to find this object");
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
                                return;
                            }

                            selectedObject = dialog.SelectedObject;
                            if (selectedObject == null)
                            {
                                return;
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
                        dbObject = FindTableTrigger(db, selectedObject.ObjectName, selectedObject.SchemaName);
                    } // Add more else-if blocks for other types like Functions, Triggers, etc.


                    if (dbObject != null)
                    {
                        System.Collections.Specialized.StringCollection sc = scripter.Script(new Urn[] { dbObject.Urn });

                        StringBuilder sb = new StringBuilder();
                        foreach (string line in sc)
                        {
                            sb.AppendLine(line);
                            sb.AppendLine("GO");
                        }
                        string fullScriptResult = sb.ToString();

                        // additional format to make it pretty
                        if (selectedObject.TypeDesc == "USER_TABLE")
                        {
                            TSql160Parser sqlParser = new TSql160Parser(false);
                            IList<ParseError> parseErrors = new List<ParseError>();
                            TSqlFragment result = sqlParser.Parse(new StringReader(fullScriptResult), out parseErrors);

                            // leave it as is if for some reason we can't format it
                            if (parseErrors.Count == 0)
                            {
                                Sql160ScriptGenerator gen = new Sql160ScriptGenerator();
                                gen.Options.AlignClauseBodies = false;
                                gen.Options.IncludeSemicolons = false;
                                gen.GenerateScript(result, out fullScriptResult);
                            }
                        }

                        ServiceCache.ScriptFactory.CreateNewBlankScript(ScriptType.Sql, connectionInfo.ActiveConnectionInfo, null);

                        // insert SQL definition to document
                        EnvDTE.TextDocument doc = (EnvDTE.TextDocument)ServiceCache.ExtensibilityModel.Application.ActiveDocument.Object(null);

                        doc.EndPoint.CreateEditPoint().Insert(fullScriptResult);

                    }
                   

                }
                catch (Exception ex)
                {

                    // Show a message box to prove we were here
                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        ex.Message,
                        "Error getting selected object",
                        OLEMSGICON.OLEMSGICON_WARNING,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

                }
            }

        }

        private static ParsedObjectName ParseSelectedObjectName(string selectedObjectName)
        {
            if (string.IsNullOrWhiteSpace(selectedObjectName))
            {
                throw new ArgumentException("Selected object name cannot be empty", nameof(selectedObjectName));
            }

            List<string> parts = SplitIdentifierParts(selectedObjectName);
            if (parts.Count == 0)
            {
                throw new ArgumentException("Selected object name cannot be empty", nameof(selectedObjectName));
            }

            string objectName = UnquoteIdentifier(parts[^1]);
            string schemaName = null;

            if (parts.Count >= 2)
            {
                schemaName = UnquoteIdentifier(parts[^2]);
            }

            return new ParsedObjectName(schemaName, objectName);
        }

        private static List<string> SplitIdentifierParts(string identifier)
        {
            List<string> parts = new List<string>();
            StringBuilder current = new StringBuilder();
            bool inBrackets = false;

            foreach (char ch in identifier)
            {
                if (ch == '[')
                {
                    inBrackets = true;
                }
                else if (ch == ']')
                {
                    inBrackets = false;
                }

                if (ch == '.' && !inBrackets)
                {
                    AddPart(parts, current);
                }
                else
                {
                    current.Append(ch);
                }
            }

            AddPart(parts, current);
            return parts;
        }

        private static void AddPart(List<string> parts, StringBuilder current)
        {
            if (current.Length == 0)
            {
                return;
            }

            parts.Add(current.ToString().Trim());
            current.Clear();
        }

        private static string UnquoteIdentifier(string identifier)
        {
            string trimmed = identifier?.Trim();
            if (string.IsNullOrEmpty(trimmed))
            {
                return trimmed;
            }

            if (trimmed.StartsWith("[", StringComparison.Ordinal) && trimmed.EndsWith("]", StringComparison.Ordinal))
            {
                return trimmed.Substring(1, trimmed.Length - 2).Replace("]]", "]", StringComparison.Ordinal);
            }

            return trimmed;
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
       DB_NAME()
FROM sys.objects o
JOIN sys.schemas s ON o.schema_id = s.schema_id
WHERE o.name = @objectName
  AND (@schemaName IS NULL OR s.name = @schemaName);";

            using (SqlCommand cmd = new SqlCommand(commandText, connection))
            {
                cmd.Parameters.Add(new SqlParameter("objectName", objectName));
                cmd.Parameters.Add(new SqlParameter("schemaName", (object)schemaName ?? DBNull.Value));

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
                            reader.GetString(4)));
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

        private static SqlSmoObject FindTableTrigger(Database db, string triggerName, string schemaName)
        {
            foreach (Table table in db.Tables)
            {
                if (!string.IsNullOrEmpty(schemaName) && !string.Equals(table.Schema, schemaName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (table.Triggers.Contains(triggerName))
                {
                    return table.Triggers[triggerName];
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

    internal sealed class ScriptObjectSelectionItem
    {
        public ScriptObjectSelectionItem(string typeDesc, string schemaName, string objectName, int objectId, string databaseName)
        {
            TypeDesc = typeDesc;
            SchemaName = schemaName;
            ObjectName = objectName;
            ObjectId = objectId;
            DatabaseName = databaseName;
        }

        public string TypeDesc { get; }

        public string SchemaName { get; }

        public string ObjectName { get; }

        public int ObjectId { get; }

        public string DatabaseName { get; }

        public string DisplayName => $"{DatabaseName}.{SchemaName}.{ObjectName} ({TypeDesc})";
    }
}
