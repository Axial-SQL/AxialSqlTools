using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    internal sealed class ExportGridToGoogleSheetCommand
    {
        public const int CommandId = 4137;

        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");

        private readonly AsyncPackage package;

        private ExportGridToGoogleSheetCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static ExportGridToGoogleSheetCommand Instance
        {
            get;
            private set;
        }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new ExportGridToGoogleSheetCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ThreadHelper.JoinableTaskFactory.RunAsync(async delegate
            {
                await ExecuteAsync();
            });
        }

        private async Task ExecuteAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            bool isShiftPressed = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);

            List<System.Data.DataTable> dataTables = GridAccess.GetDataTables();
            if (dataTables.Count == 0)
            {
                VsShellUtilities.ShowMessageBox(
                   this.package,
                   "No Data Available",
                   "No result sets are available for export.",
                   OLEMSGICON.OLEMSGICON_WARNING,
                   OLEMSGBUTTON.OLEMSGBUTTON_OK,
                   OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return;
            }

            var settings = SettingsManager.GetGoogleSheetsSettings();
            if (!settings.HasClientConfiguration())
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    "Google Sheets client ID and client secret are required. Open the Settings window and configure the Google Sheets section before exporting.",
                    "Missing Google Sheets configuration",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return;
            }

            if (string.IsNullOrWhiteSpace(settings.refreshToken))
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    "Google Sheets is not authorized yet. Use the Settings window to authorize access before exporting.",
                    "Authorization Required",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return;
            }

            try
            {
                var result = await GoogleSheetsExport.ExportToNewSpreadsheetAsync(dataTables, settings, isShiftPressed, CancellationToken.None);

                string message = "The data has been exported to Google Sheets.";
                if (!string.IsNullOrEmpty(result.SpreadsheetUrl))
                {
                    message += $"\n\n{result.SpreadsheetUrl}";
                }

                VsShellUtilities.ShowMessageBox(
                    this.package,
                    message,
                    "Export Complete",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
            catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    $"Export failed: {ex.Message}",
                    "Export Failed",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }
    }
}
