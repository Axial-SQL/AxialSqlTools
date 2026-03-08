using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace SSMS_EnvTabs
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(UIContextGuids.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(SSMS_EnvTabsPackage.PackageGuidString)]
    public sealed class SSMS_EnvTabsPackage : AsyncPackage
    {
        /// <summary>
        /// SSMS_EnvTabsPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "a155434a-a5a5-4d8f-b211-b26305addff6";

        public const string PackageCmdSetGuidString = "2E302824-2C62-4275-801B-55919612C26D";
        public static readonly Guid PackageCmdSetGuid = new Guid(PackageCmdSetGuidString);
        
        public const int cmdidCalibrate = 0x0100;
        public const int cmdidGenerateSalt = 0x0101;
        public const int cmdidCaptureData = 0x0102;
        public const int cmdidOpenConfig = 0x0103;
        public const int cmdidCheckUpdates = 0x0104;
        public const int cmdidEditActiveConnectionRule = 0x0105;

        private RdtEventManager rdtEventManager;

        internal RdtEventManager RdtEventManagerInstance => rdtEventManager;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            TabGroupConfigLoader.EnsureDefaultConfigExists();
            var initialConfig = TabGroupConfigLoader.LoadOrNull();
            TabGroupConfigLoader.UpdateConfigVersionIfNeeded(initialConfig, UpdateChecker.GetCurrentVersion());
            if (initialConfig?.Settings != null)
            {
                EnvTabsLog.Enabled = initialConfig.Settings.EnableLogging;
                EnvTabsLog.VerboseEnabled = initialConfig.Settings.EnableVerboseLogging;
                SsmsSettingsUpdater.EnsureRegexTabColorizationEnabled(initialConfig.Settings.EnableAutoColor);
            }

            EnvTabsLog.Info("SSMS EnvTabs package initialized.");
            rdtEventManager = await RdtEventManager.CreateAndStartAsync(this, cancellationToken);
            
            // Register commands
            await OpenConfigCommand.InitializeAsync(this);
            await CheckUpdatesCommand.InitializeAsync(this);
            await EditActiveConnectionRuleCommand.InitializeAsync(this);

            UpdateChecker.ScheduleCheck(this, initialConfig?.Settings);
        }

    }
}
