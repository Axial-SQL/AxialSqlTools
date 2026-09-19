using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace AxialSqlTools
{
    internal sealed class StatisticsSummaryWindowCommand
    {
        public const int CommandId = 4152;
        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");

        private readonly AsyncPackage package;

        private StatisticsSummaryWindowCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandId = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandId);
            commandService.AddCommand(menuItem);
        }

        public static StatisticsSummaryWindowCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new StatisticsSummaryWindowCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            package.JoinableTaskFactory.RunAsync(async delegate
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                try
                {
                    var window = await package.ShowToolWindowAsync(typeof(StatisticsSummaryWindow), 0, true, package.DisposalToken);
                    if (window?.Frame == null)
                    {
                        throw new NotSupportedException("Cannot create tool window");
                    }

                    StatisticsSummaryStore.SetWindowOpen(true);
                    AxialSqlToolsPackage.EnsureStatisticsExecutionHookForActiveWindow("statistics-window-open");
                }
                catch (Exception ex)
                {
                    VsShellUtilities.ShowMessageBox(
                        package,
                        "Failed to open Statistics Summary. " + ex.Message,
                        "Axial SQL Tools",
                        OLEMSGICON.OLEMSGICON_CRITICAL,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }
            });
        }
    }
}