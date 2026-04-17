using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ExportGridToExcelCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4133;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExportGridToExcelCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ExportGridToExcelCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static ExportGridToExcelCommand Instance
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
            // Switch to the main thread - the call to AddCommand in ExportGridToExcelCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new ExportGridToExcelCommand(package, commandService);
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

            // detect shift state
            bool isShiftPressed = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);      

            List<DataTable> dataTables = GridAccess.GetDataTables();

            if (dataTables.Count > 0)
            {

                string filePath = ShowSaveFileDialog();
                if (string.IsNullOrEmpty(filePath))
                {
                    return;
                }

                ExcelExport.SaveDataTableToExcel(dataTables, filePath, isShiftPressed);

                VsShellUtilities.ShowMessageBox(
                    this.package,
                    "The data has been successfully exported to:\n" + filePath,
                    "Export Complete",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
            else {
                VsShellUtilities.ShowMessageBox(
                   this.package,
                   "No Data Available",
                   "No result sets are available for export.",
                   OLEMSGICON.OLEMSGICON_WARNING,
                   OLEMSGBUTTON.OLEMSGBUTTON_OK,
                   OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }

        }

        static string ShowSaveFileDialog()
        {
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {

                var settings = SettingsManager.GetExcelExportSettings();

                // Suggest Desktop as the initial folder
                saveDialog.InitialDirectory = settings.GetDefaultDirectory();
                saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                saveDialog.DefaultExt = "xlsx"; 

                // Suggest a timestamped filename
                saveDialog.FileName = ExcelExport.ExpandDateWildcards(settings.GetDefaultFileName());

                DialogResult result = saveDialog.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(saveDialog.FileName))
                {
                    return saveDialog.FileName;
                }
                else
                {
                    return null;
                }
            }
        }

    }
}
