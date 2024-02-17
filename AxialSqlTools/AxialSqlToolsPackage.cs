using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using Task = System.Threading.Tasks.Task;

using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Windows.Forms;
using Aurora;
using System.Reflection;
using System.IO;

using AxialSqlTools.Properties;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Collections;
using Microsoft.SqlServer.Management.UI.Grid;
using System.Linq;
using System.Data;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using System.Drawing;

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
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(AxialSqlToolsPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]

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

    public sealed class AxialSqlToolsPackage : AsyncPackage
    {

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



        /// <summary>
        /// Initializes a new instance of the <see cref="AxialSqlToolsPackage"/> class.
        /// </summary>
        public AxialSqlToolsPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

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

            try
            {
                await FormatQueryCommand.InitializeAsync(this);
                await RefreshTemplatesCommand.InitializeAsync(this);
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

                await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

                DTE2 application = GetGlobalService(typeof(DTE)) as DTE2;
                IVsProfferCommands3 profferCommands3 = await base.GetServiceAsync(typeof(SVsProfferCommands)) as IVsProfferCommands3;
                OleMenuCommandService oleMenuCommandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;

                var command = application.Commands.Item("Query.Execute");
                m_queryExecuteEvent = application.Events.get_CommandEvents(command.Guid, command.ID);
                m_queryExecuteEvent.BeforeExecute += this.CommandEvents_BeforeExecute;
                m_queryExecuteEvent.AfterExecute += this.CommandEvents_AfterExecute;



                // "File.ConnectObjectExplorer"
                // "Query.Connect"
                // 

                //---------------------------------------------------------------------------
                // Query Templates
                string optionsFileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "AxialSqlTools.xml");
                Config options = Config.Load(optionsFileName);

                ImageList icons = new ImageList();
                icons.Images.Add(Resources.script);

                m_plugin = new Plugin(application, profferCommands3, icons, oleMenuCommandService, "AxialSqlTools", "Aurora.Connect", options);

                CommandBar commandBar = m_plugin.AddCommandBar("Axial SQL Tools", MsoBarPosition.msoBarTop);
                m_commandRegistry = new CommandRegistry(m_plugin, commandBar, new Guid(PackageGuidString), new Guid(PackageGuidGroup));

                m_commandBarQueryTemplates = m_plugin.AddCommandBarMenu("Query Templates", MsoBarPosition.msoBarTop, null);

                RefreshTemplatesList();

                //---------------------------------------------------------------------------
                // check for a new version
                MenuCommand Cmd = m_plugin.MenuCommandService.FindCommand(new CommandID(CheckAddinVersionCommand.CommandSet, CheckAddinVersionCommand.CommandId));

                try
                {
                    Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
                    string currentVersionString = currentVersion.ToString();

                    var checker = new GitHubReleaseChecker();

                    //TODO - fails in SSMS18
                    bool isNewVersionAvailable = await checker.IsNewVersionAvailableAsync(currentVersionString);

                    Cmd.Visible = isNewVersionAvailable;

                }
                catch { Cmd.Visible = false; }

                


            }
            catch (Exception ex)
            {

                // Show a message box to prove we were here
                VsShellUtilities.ShowMessageBox(
                    this,
                    ex.Message,
                    "Something went wrong... Please report this to info@axial-sql.com!",
                    OLEMSGICON.OLEMSGICON_WARNING,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            }

        }

        // TODO - might be needed for OxyPlot if I end up using it
        //private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        //{
        //    // add this into main module -> AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);

        //    if (args.Name.Contains("OxyPlot"))
        //        return AppDomain.CurrentDomain.Load(args.Name);
        //    else return null; 
        //}
        //----------------

        public static void AddEventToSqlResultConrol_ScriptExecutionCompleted()
        {
            var SQLResultsControl = GridAccess.GetSQLResultsControl();

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
        
        // This method aligns all numeric values to the right
        public static void SQLResultsControl_ScriptExecutionCompleted(object a, object b)
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

                        foreach (GridColumn gridColumn in gridColumns)
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
            catch // (Exception ex)
            {

            }

            try
            {
                // 2. Get open transaction info
                int openTranCount = 0;

                var SQLResultsControl = GridAccess.GetSQLResultsControl();
                var m_SqlExec = GridAccess.GetNonPublicField(SQLResultsControl, "m_sqlExec");

#if SSMS19DLL
                Microsoft.Data.SqlClient.SqlConnection connection = GridAccess.GetNonPublicField(m_SqlExec, "m_conn") as Microsoft.Data.SqlClient.SqlConnection;
                if (connection.State == ConnectionState.Open)
                {
                    using (Microsoft.Data.SqlClient.SqlCommand command = new Microsoft.Data.SqlClient.SqlCommand("SELECT @@TRANCOUNT", connection))
                    {
                        var result = command.ExecuteScalar();
                        openTranCount = result != DBNull.Value ? Convert.ToInt32(result) : 0;
                    }
                }
#endif
#if SSMS18DLL
                System.Data.SqlClient.SqlConnection connection = GridAccess.GetNonPublicField(m_SqlExec, "m_conn") as System.Data.SqlClient.SqlConnection;
                if (connection.State == ConnectionState.Open)
                    {
                        using (System.Data.SqlClient.SqlCommand command = new System.Data.SqlClient.SqlCommand("SELECT @@TRANCOUNT", connection))
                    {
                        var result = command.ExecuteScalar();
                        openTranCount = result != DBNull.Value ? Convert.ToInt32(result) : 0;
                    }
                }
#endif

                GridAccess.ChangeCurrentWindowTitle(openTranCount);

            } catch (Exception ex)
            {
                var msg = ex.Message;
            }            

        }


        private void CommandEvents_BeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                AddEventToSqlResultConrol_ScriptExecutionCompleted();
            }
            catch {

            }

        }

        //it has been executed, but the Grid hasn't been created yet...
        private void CommandEvents_AfterExecute(string Guid, int ID, object CustomIn, object CustomOut)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
        }

        public void RefreshTemplatesList()
        {
            // ThreadHelper.ThrowIfNotOnUIThread();

            //// TODO - Need to handle this properly somehow... 
            //// delete all named command bars and commands
            //DTE2 application = GetGlobalService(typeof(DTE)) as DTE2;

            //int ii = 1;
            //while (true)
            //{
            //    int removeCommandBarResult = 1;
            //    int removeCommandResult = 1;

            //    string folderName = "Folder_" + ii;
            //    string commandName = "Template_" + ii;

            //    try
            //    {
            //        CommandBars cmdBars = (Microsoft.VisualStudio.CommandBars.CommandBars)application.CommandBars;
            //        CommandBar existingCmdBar = null;
            //        existingCmdBar = cmdBars[folderName];
            //        removeCommandBarResult = m_plugin.ProfferCommands.RemoveCommandBar(existingCmdBar);
            //    } catch {}

            //    try
            //    {
            //        removeCommandResult = m_plugin.ProfferCommands.RemoveNamedCommand(commandName);
            //    } catch {}

            //    ii += 1;

            //    if ((removeCommandBarResult == 1 && removeCommandResult == 1) || ii > 1000) {break;}                
            //}
            
            Dictionary<string, string> fileNamesCache = new Dictionary<string, string>();

            string Folder = SettingsManager.GetTemplatesFolder();
            int i = 1;
            CreateCommands(ref i, ref fileNamesCache, Folder, m_commandRegistry, m_commandBarQueryTemplates);

            foreach (CommandBarControl control in m_commandBarQueryTemplates.Controls)
            {
                string keyToFind = control.Caption;
                if (fileNamesCache.TryGetValue(keyToFind, out string value1))
                {
                    control.Caption = value1;
                    control.Tag = keyToFind;
                } else if (fileNamesCache.TryGetValue(control.Tag, out string value2)) //This is the case when the file was renamed
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

            foreach (CommandBarControl control in commandBarFolder.Controls)
            {
                string keyToFind = control.Caption;
                if (fileNamesCache.TryGetValue(keyToFind, out string value))
                {
                    control.Caption = value;
                    control.Tag = keyToFind;
                } else if (fileNamesCache.TryGetValue(control.Tag, out string value2)) //This is the case when the file was renamed
                {
                    control.Caption = value2;
                }
            }

        }
       
#endregion
    }
}
