using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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

                    string selectedObjectName = selection.Text.Trim();

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

                    SqlConnection currentServerConnection = new SqlConnection(connectionInfo.FullConnectionString);
                    currentServerConnection.Open();

                    string command = $@"
                    SELECT [type_desc],
                           SCHEMA_NAME([schema_id]),
                           [name],
                           [object_id], 
                           DB_NAME()
                    FROM sys.objects
                    WHERE [object_id] = OBJECT_ID(@selectedObjectName)";

                    SqlCommand cmd = new SqlCommand(command, currentServerConnection);
                    cmd.Parameters.Add(new SqlParameter("selectedObjectName", selectedObjectName));

                    string object_type = null;
                    string object_schema = null;
                    string object_name = null;
                    int object_id = 0;
                    string database_name = null;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            object_type = reader.GetString(0);
                            object_schema = reader.GetString(1);
                            object_name = reader.GetString(2);
                            object_id = reader.GetInt32(3);
                            database_name = reader.GetString(4);

                        }
                        reader.Close();
                    }                    

                    if (object_id == 0)
                    {
                        //let's check master database too
                        // can't do in a single query with UNION because could be a collation conflict
                        string commandMaster = $@"USE [master];
                        SELECT [type_desc],
                               SCHEMA_NAME([schema_id]),
                               [name],
                               [object_id], 
                               DB_NAME()
                        FROM sys.objects
                        WHERE [object_id] = OBJECT_ID(@selectedObjectName)";

                        SqlCommand cmdMaster = new SqlCommand(commandMaster, currentServerConnection);
                        cmdMaster.Parameters.Add(new SqlParameter("selectedObjectName", selectedObjectName));

                        using (SqlDataReader reader = cmdMaster.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                object_type = reader.GetString(0);
                                object_schema = reader.GetString(1);
                                object_name = reader.GetString(2);
                                object_id = reader.GetInt32(3);
                                database_name = reader.GetString(4);
                            }
                            reader.Close();
                        }
                        currentServerConnection.Close();

                        if (object_id == 0)
                        {
                            throw new Exception("Unable to find this object");
                        }
                    }

                    currentServerConnection.Close();

                    ServerConnection SmoConnection = new ServerConnection();
                    SmoConnection.ConnectionString = connectionInfo.FullConnectionString;
                    Server server = new Server(SmoConnection);

                    Scripter scripter = new Scripter(server) { Options = new ScriptingOptions() };

                    if (object_type == "USER_TABLE")
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
                       
                    Database db = server.Databases[database_name];
                    SqlSmoObject dbObject = null;

                    if (db.Tables.Contains(object_name, object_schema))
                    {
                        dbObject = db.Tables[object_name, object_schema];
                    }
                    else if (db.StoredProcedures.Contains(object_name, object_schema))
                    {
                        dbObject = db.StoredProcedures[object_name, object_schema];
                    }
                    else if (db.UserDefinedFunctions.Contains(object_name, object_schema))
                    {
                        dbObject = db.UserDefinedFunctions[object_name, object_schema];
                    }
                    else if (db.Views.Contains(object_name, object_schema))
                    {
                        dbObject = db.Views[object_name, object_schema];
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
                        if (object_type == "USER_TABLE")
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
    }
}
