using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;
using static AxialSqlTools.AxialSqlToolsPackage;
using System.Windows.Input;
using System.Windows.Interop;

namespace AxialSqlTools
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class FormatQueryCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4131;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="FormatQueryCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private FormatQueryCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static FormatQueryCommand Instance
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
            // Switch to the main thread - the call to AddCommand in FormatQueryCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new FormatQueryCommand(package, commandService);
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
                
                var formatSettings = SettingsManager.GetTSqlCodeFormatSettings();
                if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
                {
                    // Own the dialog by VS main window
                    var dlg = new FormatOptionsDialog(formatSettings);
                    var uiShell = Package.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
                    if (uiShell != null && uiShell.GetDialogOwnerHwnd(out var hwnd) == 0 && hwnd != IntPtr.Zero)
                    {
                        new WindowInteropHelper(dlg).Owner = hwnd;
                    }

                    var ok = dlg.ShowDialog();
                    if (ok != true)
                    {
                        return; // user canceled
                    }
                    formatSettings = dlg.Settings; 
                }


                try
                {
                    TextSelection selection = dte.ActiveDocument.Selection as TextSelection;

                    string existingCommandText = selection.Text.Trim();

                    if (!string.IsNullOrEmpty(existingCommandText))
                    {
                        string result = TSqlFormatter.FormatCode(existingCommandText, formatSettings);
                        selection.Delete();
                        selection.Insert(result);
                        return;
                    }

                    // continue formatiing the entire document when nothing is selected                    
                    TextDocument textDoc = dte.ActiveDocument.Object("TextDocument") as TextDocument;
                    if (textDoc != null)
                    {
                        existingCommandText = textDoc.CreateEditPoint(textDoc.StartPoint).GetText(textDoc.EndPoint).Trim();

                        if (!string.IsNullOrEmpty(existingCommandText))
                        {
                            string result = TSqlFormatter.FormatCode(existingCommandText, formatSettings);


                            EditPoint startPoint = textDoc.StartPoint.CreateEditPoint();
                            startPoint.ReplaceText(textDoc.EndPoint, result, (int)vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
                            
                            return;
                        }
                    }

                }
                catch (Exception ex)
                {

                    // Show a message box to prove we were here
                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        ex.Message,
                        "Error parsing the code",
                        OLEMSGICON.OLEMSGICON_WARNING,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

                }
            }

        }        

    }
}
