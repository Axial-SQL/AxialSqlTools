using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace SSMS_EnvTabs
{
    internal sealed class OpenConfigCommand
    {
        public const int CommandId = 0x0103;
        public static readonly Guid CommandSet = SSMS_EnvTabsPackage.PackageCmdSetGuid;

        private readonly AsyncPackage package;

        private OpenConfigCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            new OpenConfigCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            try 
            {
                TabGroupConfigLoader.EnsureDefaultConfigExists();
                string configPath = TabGroupConfigLoader.GetUserConfigPath();
                VsShellUtilities.OpenDocument(this.package, configPath);
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"OpenConfigCommand failed: {ex.Message}");
            }
        }
    }
}
