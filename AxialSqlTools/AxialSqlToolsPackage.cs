using Aurora;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;
using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio;
using System.Collections;
using NLog;
using NLog.Targets;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Linq;
using AxialSqlTools.Properties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Net;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AxialSqlTools
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(AxialSqlToolsPackage.PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(SettingsWindow))]
    [ProvideToolWindow(typeof(AboutWindow))]
    [ProvideToolWindow(typeof(ToolWindowGridToEmail))]
    [ProvideToolWindow(typeof(HealthDashboard_Server))]
    [ProvideToolWindow(typeof(HealthDashboard_Servers))]
    [ProvideToolWindow(typeof(DataTransferWindow))]
    [ProvideToolWindow(typeof(AskChatGptWindow))]
    [ProvideToolWindow(typeof(SqlServerBuildsWindow))]
    [ProvideToolWindow(typeof(QueryHistoryWindow))]
    public sealed class AxialSqlToolsPackage : AsyncPackage
    {

        public class SQLVersionInfo
        {
            public string SqlVersion { get; set; }    // e.g. "SQL Server 2022"
            public Version BuildNumber { get; set; }   // e.g. "16.0.1000"
            public DateTime ReleaseDate { get; set; }
            public string UpdateName { get; set; }    // e.g. "CU5" or "Security Update XYZ"
            public string Url { get; set; }
        }

        public class SQLBuildsData
        {
            public Dictionary<string, List<SQLVersionInfo>> Builds { get; set; } = new Dictionary<string, List<SQLVersionInfo>>();
        }

        public SQLBuildsData SQLBuildsDataInfo;

        #region QueryHistory
        private class QueryHistoryEntry
        {
            public DateTime StartTime;
            public DateTime FinishTime;
            public string ElapsedTime;
            public long TotalRowsReturned;
            public string ExecResult;
            public string QueryText;
            public string DataSource;
            public string DatabaseName;
            public string LoginName;
            public string WorkstationId;
        }

        private static ConcurrentQueue<QueryHistoryEntry> _queryHistoryQueue = new ConcurrentQueue<QueryHistoryEntry>();
        public static Logger _logger;

        private void InitializeLogging()
        {

            var logDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "AxialSQL",
                        "AxialSQLToolsLog"
                );
            Directory.CreateDirectory(logDirectory);

            // If using the NLog.config approach:
            LogManager.Setup()
                  .LoadConfiguration(builder =>
                  {
                      // Create a file target
                      var fileTarget = new FileTarget("fileLog")
                      {
                          FileName = Path.Combine(logDirectory, "log_${shortdate}.log"),
                          Layout = "${longdate}|${level}|${logger}|${message}${exception:format=ToString}",

                          // Optionally, configure archive settings, etc.
                          ArchiveFileName = Path.Combine(logDirectory, "archive/log.{###}.txt"),
                          ArchiveNumbering = ArchiveNumberingMode.Rolling,
                          ArchiveAboveSize = 1024 * 1024 * 5, // 5 MB, for example
                          MaxArchiveFiles = 5
                      };

                      // Add the file target to the builder
                      // builder.AddTarget(fileTarget);

                      // Create a rule: "Write all logs from Info to Fatal to fileTarget"
                      builder.ForLogger()
                             .FilterMinLevel(LogLevel.Info)
                             .WriteTo(fileTarget);
                  });

            _logger = LogManager.GetCurrentClassLogger();
            // If needed, create directories here if they do not exist
            // Or do nothing if the config is specifying a folder that NLog will create automatically
        }

        private static void EnqueueDataForProcessing(QueryHistoryEntry data)
        {
            _queryHistoryQueue.Enqueue(data);
            _ = Task.Run(() => ProcessDataAsync());
        }

        private static async Task ProcessDataAsync()
        {
            while (_queryHistoryQueue.TryDequeue(out QueryHistoryEntry data))
            {
                await PersistDataAsync(data);
            }

        }

        private static async Task PersistDataAsync(QueryHistoryEntry data)
        {
            try
            {
                string connectionString = SettingsManager.GetQueryHistoryConnectionString();
                string qhTableName = SettingsManager.GetQueryHistoryTableNameOrDefault();
                string indexNameGuid = Guid.NewGuid().ToString(); // too much complexity trying to incorporate all possible table name combinations into proper index name

                if (!string.IsNullOrEmpty(connectionString))
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        await connection.OpenAsync();
                        string sql = $@"
                        IF OBJECT_ID('{qhTableName}') IS NULL
                        BEGIN
                            CREATE TABLE {qhTableName} (
                                [QueryID]           INT            IDENTITY (1, 1) NOT NULL,
                                [StartTime]         DATETIME       NOT NULL,
                                [FinishTime]        DATETIME       NOT NULL,
                                [ElapsedTime]       VARCHAR (15)   NOT NULL,
                                [TotalRowsReturned] BIGINT         NOT NULL,
                                [ExecResult]        VARCHAR (100)  NOT NULL,
                                [QueryText]         NVARCHAR (MAX) NOT NULL,
                                [DataSource]        NVARCHAR (128) NOT NULL,
                                [DatabaseName]      NVARCHAR (128) NOT NULL,
                                [LoginName]         NVARCHAR (128) NOT NULL,
                                [WorkstationId]     NVARCHAR (128) NOT NULL,
                                PRIMARY KEY CLUSTERED ([QueryID]),
                                INDEX [IDX_{indexNameGuid}_1] ([StartTime]),
                                INDEX [IDX_{indexNameGuid}_2] ([FinishTime]),
                                INDEX [IDX_{indexNameGuid}_3] ([DataSource]),
                                INDEX [IDX_{indexNameGuid}_4] ([DatabaseName])
                            );
                        END

                        INSERT INTO {qhTableName}
                            (StartTime, FinishTime, ElapsedTime, TotalRowsReturned, 
                                ExecResult, QueryText, DataSource, DatabaseName, LoginName, WorkstationId) 
                        VALUES (@StartTime, @FinishTime, @ElapsedTime, @TotalRowsReturned, 
                                    @ExecResult, @QueryText, @DataSource, @DatabaseName, @LoginName, @WorkstationId)
                        ";

                        using (SqlCommand command = new SqlCommand(sql, connection))
                        {
                            command.Parameters.AddWithValue("@StartTime", data.StartTime);
                            command.Parameters.AddWithValue("@FinishTime", data.FinishTime);
                            command.Parameters.AddWithValue("@ElapsedTime", data.ElapsedTime);
                            command.Parameters.AddWithValue("@TotalRowsReturned", data.TotalRowsReturned);
                            command.Parameters.AddWithValue("@ExecResult", data.ExecResult);
                            command.Parameters.AddWithValue("@QueryText", data.QueryText.Trim());
                            command.Parameters.AddWithValue("@DataSource", data.DataSource);
                            command.Parameters.AddWithValue("@DatabaseName", data.DatabaseName);
                            command.Parameters.AddWithValue("@LoginName", data.LoginName);
                            command.Parameters.AddWithValue("@WorkstationId", data.WorkstationId);
                            await command.ExecuteNonQueryAsync();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "[QueryHistory-PersistDataAsync]: An exception occurred");
            }

        }
        #endregion

        public const string PackageGuidString = "82ff597d-c4bc-469f-b990-637219074984";
        public const string PackageGuidGroup = "d8ef26a8-e88c-4ad1-85fd-ddc48a207530";

        private Plugin m_plugin = null;
        private CommandRegistry m_commandRegistry = null;
        private CommandBar m_commandBarQueryTemplates = null;

        public CommandEvents m_queryExecuteEvent { get; private set; }

        private int numberOfWindowsOpen = 0;

        public int GetNextToolWindowId()
        {
            numberOfWindowsOpen += 1;

            return numberOfWindowsOpen;
        }

        public Dictionary<string, string> globalSnippets = new Dictionary<string, string>();

        public static AxialSqlToolsPackage PackageInstance { get; private set; }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            PackageInstance = this;

            InitializeLogging();

            try
            {
                await FormatQueryCommand.InitializeAsync(this);
                await RefreshTemplatesCommand.InitializeAsync(this);
                await OpenTemplatesFolderCommand.InitializeAsync(this);
                await ExportGridToExcelCommand.InitializeAsync(this);
                await SettingsWindowCommand.InitializeAsync(this);
                await AboutWindowCommand.InitializeAsync(this);
                await ScriptSelectedObject.InitializeAsync(this);
                await ExportGridToAsInsertsCommand.InitializeAsync(this);
                await ToolWindowGridToEmailCommand.InitializeAsync(this);
                await HealthDashboard_ServerCommand.InitializeAsync(this);
                await HealthDashboard_ServersCommand.InitializeAsync(this);
                await DataTransferWindowCommand.InitializeAsync(this);
                await CheckAddinVersionCommand.InitializeAsync(this);
                await ResultGridCopyAsInsertCommand.InitializeAsync(this);
                await AskChatGptCommand.InitializeAsync(this);
                await SqlServerBuildsWindowCommand.InitializeAsync(this);
                await QueryHistoryWindowCommand.InitializeAsync(this);

            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred");
            }

            try
            {

                DTE2 application = GetGlobalService(typeof(DTE)) as DTE2;
                IVsProfferCommands3 profferCommands3 = await base.GetServiceAsync(typeof(SVsProfferCommands)) as IVsProfferCommands3;
                OleMenuCommandService oleMenuCommandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;

                /*
                var command = application.Commands.Item("Query.Execute");
                m_queryExecuteEvent = application.Events.get_CommandEvents(command.Guid, command.ID);
                m_queryExecuteEvent.BeforeExecute += this.CommandEvents_BeforeExecute;
                m_queryExecuteEvent.AfterExecute += this.CommandEvents_AfterExecute;
                */

                EnvDTE80.Events2 events = (EnvDTE80.Events2)application.Events;
                EnvDTE.WindowEvents windowEvents = events.WindowEvents;

                windowEvents.WindowCreated += new _dispWindowEvents_WindowCreatedEventHandler(WindowCreated_Event);

                // "File.ConnectObjectExplorer"
                // "Query.Connect"
                // 

                //---------------------------------------------------------------------------
                // Query Templates
                ImageList icons = new ImageList();
                icons.Images.Add(Resources.script);

                m_plugin = new Plugin(application, profferCommands3, icons, oleMenuCommandService, "AxialSqlTools", "Aurora.Connect");

                CommandBar commandBar = m_plugin.AddCommandBar("Axial SQL Tools", MsoBarPosition.msoBarTop);
                m_commandRegistry = new CommandRegistry(m_plugin, commandBar, new Guid(PackageGuidString), new Guid(PackageGuidGroup));

                m_commandBarQueryTemplates = m_plugin.AddCommandBarMenu("Query Templates", MsoBarPosition.msoBarTop, null);

                //---------------------------------------------------------------------------
                LoadGlobalSnippets();

                //---------------------------------------------------------------------------
                RefreshTemplatesList();

                //---------------------------------------------------------------------------
                // check for a new version
                MenuCommand Cmd = m_plugin.MenuCommandService.FindCommand(new CommandID(CheckAddinVersionCommand.CommandSet, CheckAddinVersionCommand.CommandId));

                try
                {
                    Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
                    string currentVersionString = currentVersion.ToString();
                    bool isNewVersionAvailable = false;

                    using (var client = new HttpClient())
                    {
                        // GitHub API versioning
                        client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("Axial-SQL-Tools", "Latest"));
                        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github.v3+json"));

                        // Request the latest release from GitHub API
                        var url = $"https://api.github.com/repos/Axial-SQL/AxialSqlTools/releases/latest";
                        var response = await client.GetStringAsync(url);

                        dynamic latestRelease = JsonConvert.DeserializeObject(response);
                        var latestVersion = (string)latestRelease.tag_name;

                        isNewVersionAvailable = (Version.Parse(latestVersion) > Version.Parse(currentVersionString));

                    }

                    Cmd.Visible = isNewVersionAvailable;

                }
                catch { Cmd.Visible = false; }


            }
            catch (Exception ex)
            {

                _logger.Error(ex, "An exception occurred");

                // Show a message box to prove we were here
                VsShellUtilities.ShowMessageBox(
                    this,
                    ex.Message,
                    "Oops! Something went wrong. Please report this issue on our GitHub repository and attach error log.",
                    OLEMSGICON.OLEMSGICON_WARNING,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            }

            try
            {
                SQLBuildsDataInfo = await Task.Run(() => DownloadSqlServerBuildInfo());

                MenuCommand CmdSqlServerBuilds = m_plugin.MenuCommandService.FindCommand(new CommandID(SqlServerBuildsWindowCommand.CommandSet, SqlServerBuildsWindowCommand.CommandId));
                CmdSqlServerBuilds.Visible = true;

            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred");
            }

            // needed for the OxyPlot library
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);           
            
        }

        #endregion

        private void WindowCreated_Event(EnvDTE.Window Window)
        {

            ThreadHelper.ThrowIfNotOnUIThread();

            if (SettingsManager.GetUseSnippets())
            {
                try
                {
                    // snippet processor 
                    var DocData = GridAccess.GetProperty(Window.Object, "DocData");
                    var txtMgr = (IVsTextManager)GridAccess.GetProperty(DocData, "TextManager");

                    IVsTextView textView;
                    if (txtMgr != null && txtMgr.GetActiveView(0, null, out textView) == VSConstants.S_OK)
                    {
                        //seems that you don't need to keep the object in memory
                        var CommandFilter = new KeypressCommandFilter(this, textView);
                        CommandFilter.AddToChain();
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "An exception occurred");
                }
            }

            // subscribe to the execution completed event
            try
            {

                var SQLResultsControl = GridAccess.GetNonPublicField(Window.Object, "m_sqlResultsControl");

                EventHandler eventHandler = SQLResultsControl_ScriptExecutionCompleted;

                Type targetType = SQLResultsControl.GetType();

                EventInfo eventInfo = targetType.GetEvent("ScriptExecutionCompleted");
                if (eventInfo != null)
                {
                    Delegate handlerDelegate = Delegate.CreateDelegate(eventInfo.EventHandlerType, eventHandler.Target, eventHandler.Method);
                    eventInfo.RemoveEventHandler(SQLResultsControl, handlerDelegate);
                    eventInfo.AddEventHandler(SQLResultsControl, handlerDelegate);
                }

            }
            catch (Exception ex) 
            {
                _logger.Error(ex, "An exception occurred");
            }
        }

        public void LoadGlobalSnippets()
        {

            if (SettingsManager.GetUseSnippets())
            {

                var snippetFolder = SettingsManager.GetSnippetFolder();

                if (Directory.Exists(snippetFolder))
                {
                    var allFiles = Directory.EnumerateFiles(snippetFolder, "*.sql");

                    foreach (var file in allFiles)
                    {

                        FileInfo fi = new FileInfo(file);

                        if (fi.Length < 1024 * 1024)
                        {
                            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fi.Name);

                            globalSnippets.Add(fileNameWithoutExtension, System.IO.File.ReadAllText(fi.FullName));
                        }

                    }


                }

            }



        }

        // I don't understand the purpose, but it works
        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            // add this into main module -> AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);

            if (args.Name.Contains("OxyPlot"))
                return AppDomain.CurrentDomain.Load(args.Name);
            else return null;
        }
        //----------------


        // This method aligns all numeric values to the right
        public static void SQLResultsControl_ScriptExecutionCompleted(object QEOLESQLExec, object b)
        {

            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                //1. Align numeric types to the right
                CollectionBase gridContainers = GridAccess.GetGridContainers();

                foreach (var gridContainer in gridContainers)
                {
                    var grid = GridAccess.GetNonPublicField(gridContainer, "m_grid") as GridControl;
                    var gridStorage = grid.GridStorage;
                    var schemaTable = GridAccess.GetNonPublicField(gridStorage, "m_schemaTable") as DataTable;

                    var gridColumns = GridAccess.GetNonPublicField(grid, "m_Columns") as GridColumnCollection;
                    if (gridColumns != null)
                    {

                        string[] typeToAlignRight = new string[] { "tinyint", "smallint", "int", "bigint", "money", "decimal", "numeric" };

                        List<int> columnsToAlignRight = new List<int> { };

                        for (int c = 0; c < schemaTable.Rows.Count; c++)
                        {
                            int columnOrdinal = (int)schemaTable.Rows[c][1];
                            var sqlDataTypeName = schemaTable.Rows[c][24];

                            if (typeToAlignRight.Contains(sqlDataTypeName))
                            {
                                columnsToAlignRight.Add(columnOrdinal);
                            }
                        }

                        foreach (Microsoft.SqlServer.Management.UI.Grid.GridColumn gridColumn in gridColumns)
                        {

                            if (columnsToAlignRight.Contains(gridColumn.ColumnIndex - 1) || gridColumn.ColumnIndex == 0)
                            {
                                var textAlignField = GridAccess.GetNonPublicFieldInfo(gridColumn, "TextAlign");
                                if (textAlignField != null)
                                {
                                    textAlignField.SetValue(gridColumn, System.Windows.Forms.HorizontalAlignment.Right);
                                }
                                var textAlignField2 = GridAccess.GetNonPublicFieldInfo(gridColumn, "m_myAlign");
                                if (textAlignField2 != null)
                                {
                                    textAlignField2.SetValue(gridColumn, System.Windows.Forms.HorizontalAlignment.Right);
                                }
                                var textAlignField3 = GridAccess.GetNonPublicFieldInfo(gridColumn, "m_textFormat");
                                if (textAlignField3 != null)
                                {
                                    System.Windows.Forms.TextFormatFlags flags = (System.Windows.Forms.TextFormatFlags)GridAccess.GetNonPublicField(gridColumn, "m_textFormat");
                                    textAlignField3.SetValue(gridColumn, flags | System.Windows.Forms.TextFormatFlags.Right);
                                }
                            }

                        }
                    }

                    grid.Refresh();

                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred");
            }

            try
            {
                // 2. Get open transaction info
                int openTranCount = 0;

                var SQLResultsControl = GridAccess.GetSQLResultsControl();
                var m_SqlExec = GridAccess.GetNonPublicField(SQLResultsControl, "m_sqlExec");

                Microsoft.Data.SqlClient.SqlConnection connection = GridAccess.GetNonPublicField(m_SqlExec, "m_conn") as Microsoft.Data.SqlClient.SqlConnection;
                if (connection.State == ConnectionState.Open)
                {
                    using (Microsoft.Data.SqlClient.SqlCommand command = new Microsoft.Data.SqlClient.SqlCommand("SELECT @@TRANCOUNT", connection))
                    {
                        var result = command.ExecuteScalar();
                        openTranCount = result != DBNull.Value ? Convert.ToInt32(result) : 0;
                    }
                }

                var editorProperties = GridAccess.GetNonPublicField(m_SqlExec, "editorProperties");
                var editorProperties_ElapsedTime = (string)GridAccess.GetProperty(editorProperties, "ElapsedTime");

                GridAccess.ChangeStatusBarContent(openTranCount, editorProperties_ElapsedTime);

            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred");
            }

            // query history
            try
            {

                var editorProperties = GridAccess.GetNonPublicField(QEOLESQLExec, "editorProperties");

                var textSpan = GridAccess.GetNonPublicField(QEOLESQLExec, "textSpan");

                var mConn = GridAccess.GetNonPublicField(QEOLESQLExec, "m_conn");

                var QueryHistoryObj = new QueryHistoryEntry();
                QueryHistoryObj.StartTime = (DateTime)GridAccess.GetNonPublicField(editorProperties, "startTime");
                QueryHistoryObj.FinishTime = (DateTime)GridAccess.GetNonPublicField(editorProperties, "finishTime");
                QueryHistoryObj.ElapsedTime = (string)GridAccess.GetProperty(editorProperties, "ElapsedTime");
                QueryHistoryObj.TotalRowsReturned = (long)GridAccess.GetProperty(editorProperties, "TotalRowsReturned");

                // Success, Failure -> not really clear how to track batch execution results...
                QueryHistoryObj.ExecResult = GridAccess.GetNonPublicField(QEOLESQLExec, "m_execResult").ToString();

                QueryHistoryObj.QueryText = (string)GridAccess.GetProperty(textSpan, "Text");
                QueryHistoryObj.DataSource = (string)GridAccess.GetProperty(mConn, "DataSource");
                QueryHistoryObj.DatabaseName = (string)GridAccess.GetProperty(mConn, "Database");
                QueryHistoryObj.LoginName = (string)GridAccess.GetProperty(editorProperties, "ChildLoginName");
                QueryHistoryObj.WorkstationId = (string)GridAccess.GetProperty(mConn, "WorkstationId");

                EnqueueDataForProcessing(QueryHistoryObj);

            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred");
            }


        }
        private void CommandEvents_BeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
        {
            //ThreadHelper.ThrowIfNotOnUIThread();            
        }

        //it has been executed, but the Grid hasn't been created yet...
        private void CommandEvents_AfterExecute(string Guid, int ID, object CustomIn, object CustomOut)
        {
            //ThreadHelper.ThrowIfNotOnUIThread();
        }

        public void RefreshTemplatesList()
        {

            //Delete existing autogenerated controls
            for (int idx = m_commandBarQueryTemplates.Controls.Count; idx >= 3; idx--)
            {
                CommandBarControl control = m_commandBarQueryTemplates.Controls[idx];
                if (control == null) continue;

                control.Delete();
            }

            Dictionary<string, string> fileNamesCache = new Dictionary<string, string>();

            string Folder = SettingsManager.GetTemplatesFolder();
            int i = 2;
            CreateCommands(ref i, ref fileNamesCache, Folder, m_commandRegistry, m_commandBarQueryTemplates);

            UpdateRenamedTemplatesControls(m_commandBarQueryTemplates, fileNamesCache);

        }

        private void UpdateRenamedTemplatesControls(CommandBar commandBarFolder, Dictionary<string, string> fileNamesCache)
        {
            foreach (CommandBarControl control in commandBarFolder.Controls)
            {
                string keyToFind = control.Caption;
                if (fileNamesCache.TryGetValue(keyToFind, out string value))
                {
                    control.Caption = value;
                    control.Tag = keyToFind;
                }
                else if (fileNamesCache.TryGetValue(control.Tag, out string value2)) //This is the case when the file was renamed
                {
                    control.Caption = value2;
                }
            }
        }

        private void CreateCommands(ref int i, ref Dictionary<string, string> fileNamesCache, string Folder,
                CommandRegistry m_commandRegistry, CommandBar commandBarFolder)
        {

            var dirs = Directory.GetDirectories(Folder);

            foreach (var dirStr in dirs)
            {
                var di = new FileInfo(dirStr);

                string controlName = "Folder_" + i;

                fileNamesCache.Add(controlName, di.Name);
                i = i + 1;

                CommandBar commandBarFolderNext = m_plugin.AddCommandBarMenu(controlName, MsoBarPosition.msoBarMenuBar, commandBarFolder);

                CreateCommands(ref i, ref fileNamesCache, Path.Combine(Folder, dirStr), m_commandRegistry, commandBarFolderNext);

            }

            var files = Directory.GetFiles(Folder);

            foreach (var file in files)
            {
                var fi = new FileInfo(file);

                string controlName = "Template_" + i;
                fileNamesCache.Add(controlName, fi.Name);

                var nc = new CommandProcessor(m_plugin, controlName, controlName, "");
                nc.package = this;
                nc.FullFileName = fi.FullName;
                m_commandRegistry.RegisterCommand(true, nc, true, commandBarFolder);

                i = i + 1;

            }

            UpdateRenamedTemplatesControls(commandBarFolder, fileNamesCache);

        }


        private static SQLBuildsData DownloadSqlServerBuildInfo()
        {
            try
            {
                return DownloadAndParseExcel();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred");

                return null;
            }            

        }

        private static SQLBuildsData DownloadAndParseExcel()
        {
            SQLBuildsData data = new SQLBuildsData();
            string excelUrl = "https://aka.ms/sqlserverbuilds";

            // Download the Excel file into a byte array.
            byte[] excelBytes;
            using (WebClient webClient = new WebClient())
            {
                excelBytes = webClient.DownloadData(excelUrl);
            }

            List<string> sheetNames = new List<string> { "2025", "2022", "2019", "2017", "2016", "2014", "2012"}; // not going for older versions...

            // Load the Excel file into a MemoryStream.
            using (MemoryStream stream = new MemoryStream(excelBytes))
            {
                // Open the SpreadsheetDocument for read-only access.
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;

                    // Iterate over each sheet in the workbook.
                    foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                    {
                        string sheetName = sheet.Name;
                        WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                        if (worksheetPart == null)
                        {
                            continue;
                        }

                        if(!sheetNames.Contains(sheetName))
                        {
                            continue;
                        }

                        // Get the SheetData element.
                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                        if (sheetData == null)
                        {
                            continue;
                        }

                        var buildList = new List<SQLVersionInfo>();
                        Dictionary<int, string> headers = null;
                        bool isHeaderRow = true;

                        // Iterate over the rows in the worksheet.
                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            // The first row is assumed to be the header.
                            if (isHeaderRow)
                            {
                                headers = new Dictionary<int, string>();
                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    int colIndex = GetColumnIndex(cell.CellReference);
                                    string headerValue = GetCellValue(cell, sharedStringPart);
                                    headers[colIndex] = headerValue;
                                }
                                isHeaderRow = false;
                            }
                            else
                            {
                                // Build a dictionary of header -> cell value for this row.
                                Dictionary<string, string> rowData = new Dictionary<string, string>();

                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    int colIndex = GetColumnIndex(cell.CellReference);
                                    if (headers != null && headers.ContainsKey(colIndex))
                                    {
                                        string header = headers[colIndex];
                                        string cellValue = GetCellValue(cell, sharedStringPart);
                                        rowData[header] = cellValue;
                                    }
                                }

                                // Only add the row if it contains a non-empty "Build" value.
                                if (rowData.ContainsKey("Build Number") && !string.IsNullOrWhiteSpace(rowData["Build Number"]))
                                {
                                    SQLVersionInfo info = new SQLVersionInfo
                                    {
                                        SqlVersion = sheetName,
                                        UpdateName = rowData.ContainsKey("Cumulative Update or Security ID") ? rowData["Cumulative Update or Security ID"] : null,
                                        Url = rowData.ContainsKey("KB URL") ? rowData["KB URL"] : null
                                    };

                                    string BuildNumber = rowData.ContainsKey("Build Number") ? rowData["Build Number"] : null;
                                    if (BuildNumber != null)
                                    {
                                        try
                                        {
                                            info.BuildNumber = new Version(BuildNumber.Trim());
                                        }
                                        catch {                                            
                                        }
                                    }
                                    // Parse the release date if present.
                                    if (rowData.ContainsKey("Release Date"))
                                    {
                                        string releaseDateStr = rowData["Release Date"];
                                        if (double.TryParse(releaseDateStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
                                        {
                                            try
                                            {
                                                info.ReleaseDate = DateTime.FromOADate(result);
                                            }
                                            catch { }
                                        }
                                                                         
                                    }
                                    buildList.Add(info);
                                }
                            }
                        }

                        // Store the build list keyed by the worksheet (SQL Server version) name.
                        data.Builds[sheetName] = buildList;
                    }
                }
            }

            return data;
        }

        /// <summary>
        /// Retrieves the text value of a cell. If the cell uses the shared string table,
        /// the method returns the appropriate string.
        /// </summary>
        /// <param name="cell">The cell to read.</param>
        /// <param name="sharedStringPart">The shared string table part (may be null).</param>
        /// <returns>The cell value as a string.</returns>
        private static string GetCellValue(Cell cell, SharedStringTablePart sharedStringPart)
        {
            if (cell == null)
                return null;

            string value = cell.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (sharedStringPart != null && int.TryParse(value, out int index))
                {
                    return sharedStringPart.SharedStringTable.ElementAt(index).InnerText;
                }
            }
            return value;
        }

        /// <summary>
        /// Converts a cell reference (e.g. "A1", "B2", "AA3") to a zero-based column index.
        /// </summary>
        /// <param name="cellReference">The cell reference string.</param>
        /// <returns>The zero-based column index.</returns>
        private static int GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
                return -1;

            // Extract only the letters.
            string columnReference = new string(cellReference.Where(c => Char.IsLetter(c)).ToArray());
            int columnIndex = 0;
            int factor = 1;
            for (int pos = columnReference.Length - 1; pos >= 0; pos--)
            {
                columnIndex += (columnReference[pos] - 'A' + 1) * factor;
                factor *= 26;
            }
            return columnIndex - 1; // Convert to zero-based index.
        }


    }
}
