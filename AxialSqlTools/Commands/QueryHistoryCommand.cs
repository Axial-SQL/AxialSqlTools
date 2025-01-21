using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;



namespace AxialSqlTools
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class QueryHistoryCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4144;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="QueryHistoryCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private QueryHistoryCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static QueryHistoryCommand Instance
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
            // Switch to the main thread - the call to AddCommand in QueryHistoryCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new QueryHistoryCommand(package, commandService);
        }

        private UIConnectionInfo CreateUIConnectionInfo(string connectionString)
        {
            var builder = new SqlConnectionStringBuilder(connectionString);
            var connInfo = new UIConnectionInfo();

            // SSMS expects these keys in AdvancedOptions (the name is case-sensitive):
            // "Server", "Database", "UserName", "Password", "AuthenticationType", etc.

            connInfo.ApplicationName = "AxialSQLTools";
            connInfo.ServerName = builder.DataSource;
            connInfo.PersistPassword = false;
            connInfo.OtherParams = null;

            // We usually store the database name under "DATABASE" or "Initial Catalog" in AdvancedOptions.
            connInfo.AdvancedOptions["DATABASE"] = builder.InitialCatalog;

            // Decide if integrated (Windows) auth or SQL auth:
            if (builder.IntegratedSecurity)
            {
                connInfo.AuthenticationType = 0;  // 0 => Windows
            }
            else
            {
                connInfo.AuthenticationType = 1;  // 1 => SQL Server auth
                connInfo.UserName = builder.UserID;
                connInfo.Password = builder.Password;
            }

            // Optionally set advanced options (e.g., connection timeout, etc.) if needed:
            // connInfo.AdvancedOptions["Connect Timeout"] = builder.ConnectTimeout.ToString();

            // Encrypt
            // (true => "YES", false => "NO"—SSMS uses strings here)
            if (builder.Encrypt)
            {
                connInfo.AdvancedOptions["ENCRYPT_CONNECTION"] = "True";
            }
            else
            {
                connInfo.AdvancedOptions["ENCRYPT_CONNECTION"] = "False";
            }

            // TrustServerCertificate
            // (true => "YES", false => "NO")
            if (builder.TrustServerCertificate)
            {
                connInfo.AdvancedOptions["TRUST_SERVER_CERTIFICATE"] = "True";
            }
            else
            {
                connInfo.AdvancedOptions["TRUST_SERVER_CERTIFICATE"] = "False";
            }

            return connInfo;
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

            try
            {

                /* - needs more work. XML approach works but the password is not persisted.
                 * 
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(SettingsManager.GetQueryHistoryConnectionString());

                var ci = ScriptFactoryAccess.GetCurrentConnectionInfo();

                var conn = ci.ActiveConnectionInfo.Copy();
                conn.ServerType = Guid.NewGuid();

                string xmlConnInfo = ScriptFactoryAccess.GetXmlFromUIConnectionInfo(conn);

                var uiConn = ScriptFactoryAccess.CreateConnectionInfoFromXml(xmlConnInfo);
                uiConn.ApplicationName = conn.ApplicationName;
                if (!string.IsNullOrEmpty(builder.Password))
                {
                    uiConn.Password = builder.Password;
                    uiConn.PersistPassword = true;
                }   
                */

                string QueryText = @"SELECT TOP (1000) [QueryID]
      ,[StartTime]
      ,[FinishTime]
      ,[ElapsedTime]
      ,[TotalRowsReturned]
      ,[ExecResult]
      ,[QueryText]
      ,[DataSource]
      ,[DatabaseName]
      ,[LoginName]
      ,[WorkstationId]
FROM [dbo].[QueryHistory]
WHERE 1 = 1
ORDER BY [QueryId] DESC;";

                // it doesn't accept the conn info I generate...
                object newScript = ServiceCache.ScriptFactory.CreateNewBlankScript(ScriptType.Sql); //, uiConn, null);                

                // insert SQL definition to document
                EnvDTE.TextDocument doc = (EnvDTE.TextDocument)ServiceCache.ExtensibilityModel.Application.ActiveDocument.Object(null);

                doc.EndPoint.CreateEditPoint().Insert(QueryText);

                // ServiceCache.ExtensibilityModel.Application.ExecuteCommand("Query.Execute");

            }
            catch (Exception ex)
            {

                string title = "Query History Command";

                VsShellUtilities.ShowMessageBox(
                    this.package,
                    ex.Message,
                    title,
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            }

           



        }
    }
}
