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
        private readonly QueryTemplateRefreshCoordinator _templateRefreshCoordinator = new QueryTemplateRefreshCoordinator();
        private Microsoft.VisualStudio.Threading.JoinableTask _templateStartupTask;
        private Microsoft.VisualStudio.Threading.JoinableTask _templateRefreshWorker;
        private IVsShell _templateShell;

        internal string TemplateRefreshMessage { get; private set; }

        private void StartQueryTemplatesInitialization()
        {
            // ShellInitialized only means the message loop has started. A Connect
            // dialog may still be open. Idle priority must be explicit: the SDK's
            // default background priority is allowed to run inside modal dialogs.
            _templateStartupTask = JoinableTaskFactory.StartOnIdle(async () =>
            {
                try
                {
                    await RefreshTemplatesListAsync();
                }
                catch (OperationCanceledException) when (DisposalToken.IsCancellationRequested) { }
                catch (Exception ex) { _logger?.Error(ex, "Query templates startup failed. Use Refresh Templates to retry."); }
            }, VsTaskRunContext.UIThreadIdlePriority);
        }

        public async Task<bool> RefreshTemplatesListAsync()
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
            // Publish the shared task before the scheduled worker can re-enter
            // native UI code. Repeated requests share one run, without a rerun loop.
            try
            {
                return await _templateRefreshCoordinator.RequestAsync(RefreshTemplatesCoreAsync,
                    work => _templateRefreshWorker = JoinableTaskFactory.StartOnIdle(work, VsTaskRunContext.UIThreadIdlePriority));
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
                TemplateRefreshMessage = "Could not schedule query templates refresh: " + ex.Message;
                _logger?.Error(ex, "Query templates refresh could not be scheduled.");
                SetTemplateStatus(TemplateRefreshMessage);
                return false;
            }
        }

        private async Task WaitForTemplatesUiAccessAsync()
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
            if (_templateShell == null)
            {
                _templateShell = await GetServiceAsync(typeof(SVsShell)) as IVsShell;
                await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
                if (_templateShell == null)
                    throw new InvalidOperationException("The SSMS shell is not ready for query templates.");
            }
            await QueryTemplateUiGate.WaitAsync(YieldToTemplateUiAsync, CanAccessTemplatesUi,
                token => Task.Delay(250, token), DisposalToken);
        }

        private async Task YieldToTemplateUiAsync()
        {
            DisposalToken.ThrowIfCancellationRequested();
            // StartOnIdle retains idle priority when this continuation resumes.
            // Yield between short batches so input and painting are processed.
            await Task.Yield();
            await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
        }

        private bool CanAccessTemplatesUi()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ErrorHandler.ThrowOnFailure(_templateShell.GetProperty((int)__VSSPROPID4.VSSPROPID_ShellInitialized, out object initialized));
            ErrorHandler.ThrowOnFailure(_templateShell.GetProperty((int)__VSSPROPID4.VSSPROPID_IsModal, out object modal));
            if (!(initialized is bool ready) || !(modal is bool isModal))
                throw new InvalidOperationException("SSMS did not report its initialization and modal state.");
            return ready && !isModal;
        }

        private async Task<bool> RefreshTemplatesCoreAsync()
        {
            var elapsed = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                await WaitForTemplatesUiAccessAsync();
                string folder = SettingsManager.GetTemplatesFolder();
                _logger?.Info("Query templates: discovery started for {0}.", folder);
                // Complete discovery before touching the working menu. No file
                // contents are read until the user selects a script.
                var catalog = await Task.Run(() => QueryTemplateCatalog.Load(folder, DisposalToken), DisposalToken);
                // A Connect dialog can open while filesystem discovery runs.
                await WaitForTemplatesUiAccessAsync();
                _logger?.Info("Query templates: discovered {0} scripts in {1} ms; preparing menu.", catalog.FileCount, elapsed.ElapsedMilliseconds);
                if (!string.Equals(folder, SettingsManager.GetTemplatesFolder(), StringComparison.OrdinalIgnoreCase))
                {
                    TemplateRefreshMessage = "The templates folder changed during discovery. Use Refresh Templates to load the new folder.";
                    SetTemplateStatus(TemplateRefreshMessage);
                    return false;
                }

                await EnsureTemplatesMenuAsync();
                var uiBatcher = new QueryTemplateUiBatcher(CanAccessTemplatesUi,
                    token => WaitForTemplatesUiAccessAsync(), () => elapsed.ElapsedMilliseconds);
                Func<Task> beforeUiOperation = () => uiBatcher.CheckpointAsync(DisposalToken);
                var warnings = new List<string>(catalog.Warnings);
                var activeCommands = new HashSet<string>(StringComparer.Ordinal);
                CollectTemplateCommands(catalog.Root, activeCommands);

                int failures = 0;
                long menuStartedAt = elapsed.ElapsedMilliseconds;
                // Remove persisted controls while their native identities are
                // still available. Tags and friendly captions do not identify them.
                _logger?.Info("Query templates: removing old menu controls.");
                await QueryTemplateMenuControls.ClearAsync(m_commandBarQueryTemplates, m_plugin.Commands,
                    new CommandID(RefreshTemplatesCommand.CommandSet, RefreshTemplatesCommand.CommandId),
                    new CommandID(OpenTemplatesFolderCommand.CommandSet, OpenTemplatesFolderCommand.CommandId),
                    ex => _logger?.Warn(ex, "Could not remove a detached templates submenu."),
                    beforeUiOperation, DisposalToken);
                _logger?.Info("Query templates: old menu controls removed after {0} ms.", elapsed.ElapsedMilliseconds);
                try
                {
                    await m_commandRegistry.RemoveUnusedCommandsAsync(activeCommands,
                        beforeUiOperation, DisposalToken);
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    failures++;
                    warnings.Add("Could not remove an old template command: " + ex.Message);
                }
                var menuResult = await QueryTemplateMenuControls.PopulateAsync(catalog.Root, m_commandBarQueryTemplates,
                    m_plugin, m_commandRegistry,
                    file => new CommandProcessor(m_plugin, file.CommandName, file.Name, file.FullPath)
                    {
                        package = this,
                        FullFileName = file.FullPath
                    }, warnings, beforeUiOperation, DisposalToken);
                int loaded = menuResult.Loaded;
                failures += menuResult.Failures;
                await beforeUiOperation();
                var shell = GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
                if (shell != null) ErrorHandler.ThrowOnFailure(shell.UpdateCommandUI(0));
                _logger?.Info("Query templates: menu work took {0} ms, including {1} batch waits totaling {2} ms.",
                    elapsed.ElapsedMilliseconds - menuStartedAt, uiBatcher.YieldCount, uiBatcher.WaitMilliseconds);
                foreach (string warning in warnings)
                    _logger?.Warn("Query templates: {0}", warning);

                TemplateRefreshMessage = "Loaded " + loaded + " of " + catalog.FileCount + " query templates."
                    + (warnings.Count == 0 ? string.Empty : " " + warnings.Count + " warning(s); see the extension log.");
                _logger?.Info("Query templates: {0} Folder: {1}; elapsed {2} ms.", TemplateRefreshMessage, folder, elapsed.ElapsedMilliseconds);
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
                _logger?.Error(ex, "Query templates refresh failed after {0} ms.", elapsed.ElapsedMilliseconds);
                SetTemplateStatus(TemplateRefreshMessage);
                return false;
            }
        }

        private async Task EnsureTemplatesMenuAsync()
        {
            await WaitForTemplatesUiAccessAsync();
            if (m_plugin == null)
            {
                var application = await GetServiceAsync(typeof(DTE)) as DTE2;
                var proffer = await GetServiceAsync(typeof(SVsProfferCommands)) as IVsProfferCommands3;
                var menuService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
                await WaitForTemplatesUiAccessAsync();
                if (application == null || proffer == null || menuService == null)
                    throw new InvalidOperationException("The SSMS command services are not ready. Use Refresh Templates to retry.");
                _templateIcons = new ImageList();
                _templateIcons.Images.Add(Properties.Resources.script);
                m_plugin = new Plugin(application, proffer, _templateIcons, menuService, "AxialSqlTools", "Aurora.Connect",
                    (ex, message) => _logger?.Warn(ex, message));
            }

            // Resolve the VSCT-owned root on every attempt. A detached replacement
            // or an old COM reference cannot populate the visible toolbar menu.
            await WaitForTemplatesUiAccessAsync();
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
