using Aurora;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    public sealed partial class AxialSqlToolsPackage
    {
        private Plugin m_plugin;
        private CommandRegistry m_commandRegistry;
        private CommandBar m_commandBarQueryTemplates;
        private ImageList _templateIcons;
        private Task<bool> _templateRefreshTask;
        private Microsoft.VisualStudio.Threading.JoinableTask _templateStartupTask;
        private bool _templateRefreshRequested;

        internal string TemplateRefreshMessage { get; private set; }

        private void StartQueryTemplatesInitialization()
        {
            // Background package loading can precede command-bar restoration.
            // Yield to the shell, then retry failed registration without requiring
            // the user to restart SSMS. Manual refresh uses the same recovery path.
            _templateStartupTask = JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    await KnownUIContexts.ShellInitializedContext;
                    DisposalToken.ThrowIfCancellationRequested();
                    for (int attempt = 0; attempt < 5; attempt++)
                    {
                        await Task.Delay(250 * (1 << attempt), DisposalToken);
                        await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
                        if (await RefreshTemplatesListAsync()) return;
                    }
                }
                catch (OperationCanceledException) { }
            });
        }

        public async Task<bool> RefreshTemplatesListAsync()
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
            if (_templateRefreshTask != null && !_templateRefreshTask.IsCompleted)
            {
                _templateRefreshRequested = true;
                return await _templateRefreshTask;
            }
            _templateRefreshTask = RefreshRequestedTemplatesAsync();
            return await _templateRefreshTask;
        }

        private async Task<bool> RefreshRequestedTemplatesAsync()
        {
            bool succeeded;
            do
            {
                _templateRefreshRequested = false;
                succeeded = await RefreshTemplatesCoreAsync();
            }
            while (_templateRefreshRequested && !DisposalToken.IsCancellationRequested);
            return succeeded;
        }

        private async Task<bool> RefreshTemplatesCoreAsync()
        {
            try
            {
                string folder = SettingsManager.GetTemplatesFolder();
                // Complete discovery before touching the working menu. No file
                // contents are read until the user selects a script.
                var catalog = await Task.Run(() => QueryTemplateCatalog.Load(folder, DisposalToken), DisposalToken);
                await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
                if (!string.Equals(folder, SettingsManager.GetTemplatesFolder(), StringComparison.OrdinalIgnoreCase))
                {
                    _templateRefreshRequested = true;
                    return false;
                }

                await EnsureTemplatesMenuAsync();
                var warnings = new List<string>(catalog.Warnings);
                var activeCommands = new HashSet<string>(StringComparer.Ordinal);
                CollectTemplateCommands(catalog.Root, activeCommands);

                int failures = 0;
                // Remove persisted controls while their native identities are
                // still available. Tags and friendly captions do not identify them.
                QueryTemplateMenuControls.Clear(m_commandBarQueryTemplates, m_plugin.Commands,
                    new CommandID(RefreshTemplatesCommand.CommandSet, RefreshTemplatesCommand.CommandId),
                    new CommandID(OpenTemplatesFolderCommand.CommandSet, OpenTemplatesFolderCommand.CommandId),
                    ex => _logger?.Debug(ex, "Could not remove a detached templates submenu."));
                try
                {
                    m_commandRegistry.RemoveUnusedCommands(activeCommands);
                }
                catch (Exception ex)
                {
                    failures++;
                    warnings.Add("Could not remove an old template command: " + ex.Message);
                }
                int loaded = QueryTemplateMenuControls.Populate(catalog.Root, m_commandBarQueryTemplates,
                    m_plugin, m_commandRegistry,
                    file => new CommandProcessor(m_plugin, file.CommandName, file.Name, file.FullPath)
                    {
                        package = this,
                        FullFileName = file.FullPath
                    }, warnings, ref failures);
                var shell = GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
                if (shell != null) ErrorHandler.ThrowOnFailure(shell.UpdateCommandUI(0));
                foreach (string warning in warnings)
                    _logger?.Warn("Query templates: {0}", warning);

                TemplateRefreshMessage = "Loaded " + loaded + " of " + catalog.FileCount + " query templates."
                    + (warnings.Count == 0 ? string.Empty : " " + warnings.Count + " warning(s); see the extension log.");
                _logger?.Info("Query templates: {0} Folder: {1}", TemplateRefreshMessage, folder);
                SetTemplateStatus(TemplateRefreshMessage);
                return failures == 0;
            }
            catch (OperationCanceledException) when (DisposalToken.IsCancellationRequested)
            {
                return false;
            }
            catch (Exception ex)
            {
                await JoinableTaskFactory.SwitchToMainThreadAsync();
                TemplateRefreshMessage = "Could not refresh query templates: " + ex.Message;
                _logger?.Error(ex, "Query templates refresh failed.");
                SetTemplateStatus(TemplateRefreshMessage);
                return false;
            }
        }

        private async Task EnsureTemplatesMenuAsync()
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
            if (m_plugin == null)
            {
                var application = await GetServiceAsync(typeof(DTE)) as DTE2;
                var proffer = await GetServiceAsync(typeof(SVsProfferCommands)) as IVsProfferCommands3;
                var menuService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
                await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
                if (application == null || proffer == null || menuService == null)
                    throw new InvalidOperationException("The SSMS command services are not ready. Use Refresh Templates to retry.");
                _templateIcons = new ImageList();
                _templateIcons.Images.Add(Properties.Resources.script);
                m_plugin = new Plugin(application, proffer, _templateIcons, menuService, "AxialSqlTools", "Aurora.Connect",
                    (ex, message) => _logger?.Warn(ex, message));
            }

            // Resolve the VSCT-owned root on every attempt. A detached replacement
            // or an old COM reference cannot populate the visible toolbar menu.
            m_commandBarQueryTemplates = m_plugin.GetCommandBar(RefreshTemplatesCommand.CommandSet, 0x1100);
            if (m_commandRegistry == null)
                m_commandRegistry = new CommandRegistry(m_plugin, m_commandBarQueryTemplates,
                    new Guid(PackageGuidString), new Guid(PackageGuidGroup));
        }

        private static void CollectTemplateCommands(QueryTemplateFolder folder, HashSet<string> names)
        {
            foreach (var file in folder.Files) names.Add(file.CommandName);
            foreach (var child in folder.Folders) CollectTemplateCommands(child, names);
        }

        private static void SetTemplateStatus(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var statusBar = GetGlobalService(typeof(SVsStatusbar)) as IVsStatusbar;
                statusBar?.SetText(message.Replace('\r', ' ').Replace('\n', ' '));
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                _logger?.Warn(ex, "Could not update the SSMS status bar for query templates.");
            }
        }
    }
}
