using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
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


            ThreadHelper.ThrowIfNotOnUIThread();
            DTE dte = Package.GetGlobalService(typeof(DTE)) as DTE;

            if (dte?.ActiveDocument != null)
            {

                try
                {
                    TextSelection selection = dte.ActiveDocument.Selection as TextSelection;

                    string selectedObjectName = selection.Text.Trim();

                    //Could be:
                    // - Table
                    // - Sproc
                    // - Function

                    // find all object with this name
                    // script properly from a current connection
                    // display in a single window

                    //////UIConnectionInfo connection = ServiceCache.ScriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo;

                    //////SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

                    //////builder.DataSource = connection.ServerName;
                    //////builder.IntegratedSecurity = string.IsNullOrEmpty(connection.Password);
                    //////builder.Password = connection.Password;
                    //////builder.UserID = connection.UserName;
                    //////builder.InitialCatalog = connection.AdvancedOptions["DATABASE"];
                    //////builder.ApplicationName = "Axial SQL Tools";

                    //////string connectionString = builder.ToString();

                    //////SqlConnection currentServerConnetion = new SqlConnection(connectionString);
                    //////currentServerConnetion.Open();

                    //////string command = "SELECT * FROM sys.objects WHERE [object_id] = OBJECT_ID(@selectedObjectName)";
                    //////SqlCommand cmd = new SqlCommand(command, currentServerConnetion);
                    //////cmd.Parameters.Add(new SqlParameter[] { new SqlParameter("selectedObjectName", selectedObjectName) });

                    //////string object_type = null;
                    //////string object_schema = null;
                    //////string object_name = null;

                    //////using (SqlDataReader reader = cmd.ExecuteReader())
                    //////{
                    //////    if (reader.Read())
                    //////    {
                    //////        object_type = reader.GetString(0);  
                    //////        object_schema = reader.GetString(0);
                    //////        object_name = reader.GetString(0);

                    //////    }
                    //////    reader.Close();
                    //////}


                    //////var a = 0;
                  
                    

                    //ServerConnection SmoConnection = new ServerConnection(currentServerConnetion);
                    //Server server = new Server(SmoConnection);



                    //Scripter scripter = new Scripter(server)
                    //{
                    //    Options = {
                    //        ScriptData = false,
                    //        ScriptSchema = true,
                    //        ScriptDrops = false,
                    //        WithDependencies = true,
                    //        Indexes = true, // Include indexes
                    //        NoCollation = false, // Specify collation
                    //        // Other options as needed
                            
                    //    }
                    //};

                    //// Select the object to script
                    //Database db = server.Databases[databaseName];

                    
                    //Table myTable = db.Tables["your_table_name", "your_schema_name"];

                    //// Generate the script
                    //IEnumerable<string> script = scripter.Script(new Urn[] { myTable.Urn });

                    //// Output the script
                    //foreach (var line in script)
                    //{
                    //    Console.WriteLine(line);
                    //}



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


            //string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            //string title = "ScriptSelectedObject";

            //// Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.package,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
