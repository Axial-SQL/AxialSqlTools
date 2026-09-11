using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;

namespace AxialSqlTools.DataCompare
{
    internal sealed class DataCompareWindowCommand
    {
        public const int CommandId = 4156;
        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");
        private readonly AsyncPackage package;
        private DataCompareWindowCommand(AsyncPackage package, OleMenuCommandService commands)
        {
            this.package = package;
            commands.AddCommand(new MenuCommand(Execute, new CommandID(CommandSet, CommandId)));
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commands = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commands == null) throw new InvalidOperationException("SSMS command service is unavailable.");
            new DataCompareWindowCommand(package, commands);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            int id = ((AxialSqlToolsPackage)package).GetNextToolWindowId();
            var window = package.FindToolWindow(typeof(DataCompareWindow), id, true);
            if (!(window?.Frame is IVsWindowFrame frame)) throw new InvalidOperationException("Cannot create the Data Compare window.");
            ErrorHandler.ThrowOnFailure(frame.SetProperty((int)__VSFPROPID.VSFPROPID_FrameMode, VSFRAMEMODE.VSFM_MdiChild));
            ErrorHandler.ThrowOnFailure(frame.Show());
        }
    }
}
