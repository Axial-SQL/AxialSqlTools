using Aurora;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;
using AxialSqlTools.Properties;
using NLog;
using NLog.Targets;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

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
    [ProvideToolWindow(typeof(SettingsWindow), Style = VsDockStyle.MDI)]
    [ProvideToolWindow(typeof(AboutWindow), Style = VsDockStyle.MDI)]
    [ProvideToolWindow(typeof(ToolWindowGridToEmail), Style = VsDockStyle.MDI)]
    [ProvideToolWindow(typeof(HealthDashboard_Server), Style = VsDockStyle.MDI)]
    [ProvideToolWindow(typeof(DataCompare.DataCompareWindow), MultiInstances = true, Style = VsDockStyle.MDI)]
    [ProvideToolWindow(typeof(DataTransferWindow), Style = VsDockStyle.MDI)]
    [ProvideToolWindow(typeof(SqlServerBuildsWindow), Style = VsDockStyle.MDI)]
    [ProvideToolWindow(typeof(QueryHistoryWindow), Style = VsDockStyle.MDI)]
    [ProvideToolWindow(typeof(StatisticsSummaryWindow))]
    [ProvideToolWindow(typeof(DatabaseScripterToolWindow), Style = VsDockStyle.MDI)]
    [ProvideToolWindow(typeof(DataImportWindow), Style = VsDockStyle.MDI)]
    [ProvideToolWindow(typeof(QuickSearchWindow), Style = VsDockStyle.MDI)]
    [ProvideToolWindow(typeof(SnippetManagerWindow), Style = VsDockStyle.MDI)]
    public sealed partial class AxialSqlToolsPackage : AsyncPackage
    {
        public SQLBuildsData SQLBuildsDataInfo = new SQLBuildsData();

        public SQLBuildsLoadResult SQLBuildsLoadState { get; private set; }

        public SQLBuildsLoadResult SQLBuildsLastSuccess { get; private set; }

        public bool SQLBuildsIsLoading { get; private set; }

        public event EventHandler SQLBuildsChanged;

        private Task _sqlBuildsRefreshTask;

        public Task RefreshSqlServerBuildsAsync(string localFile = null)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_sqlBuildsRefreshTask != null && !_sqlBuildsRefreshTask.IsCompleted) return _sqlBuildsRefreshTask;
            return _sqlBuildsRefreshTask = RefreshSqlServerBuildsCoreAsync(localFile);
        }

        private async Task RefreshSqlServerBuildsCoreAsync(string localFile)
        {
            SQLBuildsIsLoading = true;
            SQLBuildsChanged?.Invoke(this, EventArgs.Empty);
            try
            {
                var result = await Task.Run(() => SQLBuilds.DownloadSqlServerBuildInfo(localFile));
                await JoinableTaskFactory.SwitchToMainThreadAsync();
                SQLBuildsLoadState = result;
                if (result.HasData)
                {
                    SQLBuildsDataInfo = result.Data;
                    SQLBuildsLastSuccess = result;
                }
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "SQL Server build refresh failed.");
                SQLBuildsLoadState = new SQLBuildsLoadResult { Source = localFile ?? SQLBuilds.SourceUrl,
                    Error = "SQL Server build refresh failed. Retry the download or open a readable .xlsx workbook.",
                    Details = ex.GetType().Name + ": " + ex.Message };
            }
            finally
            {
                SQLBuildsIsLoading = false;
                SQLBuildsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

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

        public const string PackageGuidString = "82ff597d-c4bc-469f-b990-637219074984";

        public const string PackageGuidGroup = "d8ef26a8-e88c-4ad1-85fd-ddc48a207530";

        private int numberOfWindowsOpen = 0;

        public int GetNextToolWindowId()
        {
            numberOfWindowsOpen += 1;

            return numberOfWindowsOpen;
        }

        public static AxialSqlToolsPackage PackageInstance { get; private set; }

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
                await ExportGridToGoogleSheetCommand.InitializeAsync(this);
                await SettingsWindowCommand.InitializeAsync(this);
                await AboutWindowCommand.InitializeAsync(this);
                await ScriptSelectedObject.InitializeAsync(this);
                await ExportGridToAsInsertsCommand.InitializeAsync(this);
                await ToolWindowGridToEmailCommand.InitializeAsync(this);
                await HealthDashboard_ServerCommand.InitializeAsync(this);
                await DataTransferWindowCommand.InitializeAsync(this);
                await DataCompare.DataCompareWindowCommand.InitializeAsync(this);
                await DataImportWindowCommand.InitializeAsync(this);
                await ResultGridCopyAsInsertCommand.InitializeAsync(this);
                await SqlServerBuildsWindowCommand.InitializeAsync(this);
                await QueryHistoryWindowCommand.InitializeAsync(this);
                await StatisticsSummaryWindowCommand.InitializeAsync(this);
                await DatabaseScripterToolWindowCommand.InitializeAsync(this);
                await QuickSearchWindowCommand.InitializeAsync(this);
                await SnippetManagerWindowCommand.InitializeAsync(this);
                await SelectCurrentStatementCommand.InitializeAsync(this);
                await ToggleBlockCommentCommand.InitializeAsync(this);
                await SortSelectedTextCommand.InitializeAsync(this);

                UpdateChecker.ScheduleCheck(this, SettingsManager.GetEnableUpdateChecks());

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

                var command = application.Commands.Item("Query.Execute");
                m_queryExecuteEvent = application.Events.get_CommandEvents(command.Guid, command.ID);
                m_queryExecuteEvent.BeforeExecute += this.CommandEvents_BeforeExecute;
                m_queryExecuteEvent.AfterExecute += this.CommandEvents_AfterExecute;

                EnvDTE80.Events2 events = (EnvDTE80.Events2)application.Events;
                EnvDTE.WindowEvents windowEvents = events.WindowEvents;

                windowEvents.WindowCreated += new _dispWindowEvents_WindowCreatedEventHandler(WindowCreated_Event);
                windowEvents.WindowActivated += new _dispWindowEvents_WindowActivatedEventHandler(WindowActivated_Event);
                windowEvents.WindowClosing += new _dispWindowEvents_WindowClosingEventHandler(WindowClosing_Event);

                StartActiveWindowConnectionMonitor();

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
                SnippetService.ReloadSnippets();

                //---------------------------------------------------------------------------
                RefreshTemplatesList();

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
                await JoinableTaskFactory.SwitchToMainThreadAsync();
                MenuCommand CmdSqlServerBuilds = m_plugin.MenuCommandService.FindCommand(new CommandID(SqlServerBuildsWindowCommand.CommandSet, SqlServerBuildsWindowCommand.CommandId));
                CmdSqlServerBuilds.Visible = true;
                await RefreshSqlServerBuildsAsync();

            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred");
            }

            // needed for the OxyPlot library
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
            
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                UpdateChecker.LaunchDeferredUpdateOnClose();
            }

            base.Dispose(disposing);
        }

        // I don't understand the purpose, but it works
        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            // add this into main module -> AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);

            if (args.Name.Contains("OxyPlot"))
                return AppDomain.CurrentDomain.Load(args.Name);
            else return null;
        }
    }
}
