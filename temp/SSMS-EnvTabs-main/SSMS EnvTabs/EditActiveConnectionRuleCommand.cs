using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace SSMS_EnvTabs
{
    internal sealed class EditActiveConnectionRuleCommand
    {
        public const int CommandId = 0x0105;
        public static readonly Guid CommandSet = SSMS_EnvTabsPackage.PackageCmdSetGuid;

        private readonly SSMS_EnvTabsPackage package;

        private EditActiveConnectionRuleCommand(SSMS_EnvTabsPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static async Task InitializeAsync(SSMS_EnvTabsPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            new EditActiveConnectionRuleCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (package.RdtEventManagerInstance == null)
                {
                    MessageBox.Show(
                        "EnvTabs is not fully initialized yet. Please try again.",
                        "SSMS EnvTabs",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                bool handled = package.RdtEventManagerInstance.EditRuleForActiveConnection();
                if (!handled)
                {
                    MessageBox.Show(
                        "No editable connection rule was found for the active tab.",
                        "SSMS EnvTabs",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"EditActiveConnectionRuleCommand failed: {ex.Message}");
            }
        }
    }
}