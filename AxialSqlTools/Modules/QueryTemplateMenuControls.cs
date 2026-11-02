using EnvDTE;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    internal sealed class QueryTemplateMenuResult
    {
        internal QueryTemplateMenuResult(int loaded, int failures) { Loaded = loaded; Failures = failures; }
        internal int Loaded { get; }
        internal int Failures { get; }
    }

    internal static class QueryTemplateMenuControls
    {
        private const int MaximumMenuDepth = 64;
        private const int MaximumControlsPerMenu = 50000;

        private sealed class ChildMenu
        {
            internal ChildMenu(QueryTemplateFolder folder, CommandBarPopup popup, CommandBar bar)
            {
                Folder = folder;
                Popup = popup;
                Bar = bar;
            }

            internal QueryTemplateFolder Folder { get; }
            internal CommandBarPopup Popup { get; }
            internal CommandBar Bar { get; }
        }

        internal static bool TryGetCommandId(Commands commands, CommandBarControl control, out CommandID id)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            id = null;
            if (commands == null || control == null) return false;
            try
            {
                commands.CommandInfo(control, out string nativeGuid, out int nativeId);
                if (nativeId <= 0 || !Guid.TryParse(nativeGuid, out var group)) return false;
                id = new CommandID(group, nativeId);
                return true;
            }
            catch (COMException) { return false; }
            catch (ArgumentException) { return false; }
        }

        internal static async Task ClearAsync(CommandBar root, Commands commands, CommandID refreshCommand,
            CommandID openFolderCommand, Action<Exception> cleanupWarning, Func<Task> beforeUiOperation,
            CancellationToken cancellationToken)
        {
            if (root == null) throw new ArgumentNullException(nameof(root));
            if (commands == null) throw new ArgumentNullException(nameof(commands));
            if (refreshCommand == null) throw new ArgumentNullException(nameof(refreshCommand));
            if (openFolderCommand == null) throw new ArgumentNullException(nameof(openFolderCommand));
            if (refreshCommand.Equals(openFolderCommand))
                throw new ArgumentException("The two built-in template actions must have distinct command identities.");

            var snapshot = await SnapshotAsync(root, beforeUiOperation, cancellationToken);
            // Classify before any mutation, and retain that decision throughout
            // deletion even if a subsequent native identity query fails.
            var dynamicControls = await GetDynamicControlsAsync(snapshot, commands, refreshCommand,
                openFolderCommand, beforeUiOperation, cancellationToken);
            await CheckpointAsync(beforeUiOperation, cancellationToken);
            var ancestors = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { root.Name };
            for (int i = dynamicControls.Count - 1; i >= 0; i--)
                await DeleteControlAsync(dynamicControls[i], cleanupWarning, ancestors, 0,
                    beforeUiOperation, cancellationToken);

            var remaining = await SnapshotAsync(root, beforeUiOperation, cancellationToken);
            if ((await GetDynamicControlsAsync(remaining, commands, refreshCommand, openFolderCommand,
                beforeUiOperation, cancellationToken)).Count != 0)
                throw new InvalidOperationException("SSMS could not remove an existing template menu item. Refresh was stopped to avoid duplicate scripts.");
        }

        internal static Task<QueryTemplateMenuResult> PopulateAsync(QueryTemplateFolder folder, CommandBar bar,
            Aurora.Plugin plugin, Aurora.CommandRegistry registry, Func<QueryTemplateFile, Aurora.CommandBase> createHandler,
            List<string> warnings, Func<Task> beforeUiOperation, CancellationToken cancellationToken)
        {
            return PopulateFolderAsync(folder, bar, plugin, registry, createHandler, warnings,
                beforeUiOperation, cancellationToken, 0);
        }

        private static async Task<QueryTemplateMenuResult> PopulateFolderAsync(QueryTemplateFolder folder,
            CommandBar bar, Aurora.Plugin plugin, Aurora.CommandRegistry registry,
            Func<QueryTemplateFile, Aurora.CommandBase> createHandler, List<string> warnings,
            Func<Task> beforeUiOperation, CancellationToken cancellationToken, int depth)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (depth > MaximumMenuDepth)
            {
                warnings.Add("Templates menu nesting exceeds the supported depth: " + folder.FullPath);
                return new QueryTemplateMenuResult(0, 1);
            }
            int loaded = 0;
            int failures = 0;
            var childMenus = new List<ChildMenu>();
            foreach (var child in folder.Folders)
            {
                await CheckpointAsync(beforeUiOperation, cancellationToken);
                try
                {
                    var popup = plugin.AddCommandBarMenuPopup(child.CommandName, bar, appendOnly: true);
                    popup.Caption = child.Name;
                    popup.Visible = true;
                    popup.Enabled = false;
                    childMenus.Add(new ChildMenu(child, popup, popup.CommandBar));
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    failures++;
                    warnings.Add("Could not create templates menu '" + child.FullPath + "': " + ex.Message);
                }
            }

            foreach (var file in folder.Files)
            {
                await CheckpointAsync(beforeUiOperation, cancellationToken);
                try
                {
                    // ClearAsync has already removed all dynamic controls. Avoid
                    // rescanning the preceding buttons for every added script.
                    registry.RegisterCommand(false, createHandler(file), true, bar, appendOnly: true);
                    loaded++;
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    failures++;
                    warnings.Add("Could not register template '" + file.FullPath + "': " + ex.Message);
                }
            }
            foreach (var child in childMenus)
            {
                var result = await PopulateFolderAsync(child.Folder, child.Bar, plugin, registry, createHandler,
                    warnings, beforeUiOperation, cancellationToken, depth + 1);
                loaded += result.Loaded;
                failures += result.Failures;
                await CheckpointAsync(beforeUiOperation, cancellationToken);
                try
                {
                    // Publish this branch's state directly. Opening a submenu or
                    // a global command-status refresh must not activate its scripts.
                    child.Popup.Enabled = result.Loaded > 0;
                    child.Popup.Visible = true;
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    failures++;
                    warnings.Add("Could not enable templates menu '" + child.Folder.FullPath + "': " + ex.Message);
                }
            }
            return new QueryTemplateMenuResult(loaded, failures);
        }

        private static async Task<List<CommandBarControl>> GetDynamicControlsAsync(List<CommandBarControl> controls,
            Commands commands, CommandID refreshCommand, CommandID openFolderCommand,
            Func<Task> beforeUiOperation, CancellationToken cancellationToken)
        {
            bool hasRefresh = false;
            bool hasOpenFolder = false;
            var dynamicControls = new List<CommandBarControl>();
            foreach (var control in controls)
            {
                await CheckpointAsync(beforeUiOperation, cancellationToken);
                if (TryGetCommandId(commands, control, out var id) && refreshCommand.Equals(id))
                    hasRefresh = true;
                else if (id != null && openFolderCommand.Equals(id))
                    hasOpenFolder = true;
                else
                    dynamicControls.Add(control);
            }
            if (!hasRefresh || !hasOpenFolder)
                throw new InvalidOperationException("The Query Templates menu is not ready: its built-in actions could not be verified.");
            return dynamicControls;
        }

        private static async Task<List<CommandBarControl>> SnapshotAsync(CommandBar bar,
            Func<Task> beforeUiOperation, CancellationToken cancellationToken)
        {
            await CheckpointAsync(beforeUiOperation, cancellationToken);
            int count = bar.Controls.Count;
            if (count < 0 || count > MaximumControlsPerMenu)
                throw new InvalidOperationException("The templates menu contains an unexpected number of controls.");
            var controls = new List<CommandBarControl>(count);
            // Capture Count once. Never follow a growing native COM enumerator.
            for (int index = 1; index <= count; index++)
            {
                await CheckpointAsync(beforeUiOperation, cancellationToken);
                controls.Add(bar.Controls[index]);
            }
            return controls;
        }

        private static async Task DeleteControlAsync(CommandBarControl control, Action<Exception> cleanupWarning,
            HashSet<string> ancestors, int depth, Func<Task> beforeUiOperation, CancellationToken cancellationToken)
        {
            await CheckpointAsync(beforeUiOperation, cancellationToken);
            CommandBar child = null;
            string childName = null;
            bool enteredChild = false;
            if (control is CommandBarPopup popup)
            {
                try
                {
                    child = popup.CommandBar;
                    childName = child.Name;
                    if (depth >= MaximumMenuDepth || ancestors.Contains(childName))
                    {
                        // Remove only the cyclic/deep popup, never its ancestor
                        // bar (which may be the root containing Refresh).
                        child = null;
                        cleanupWarning?.Invoke(new InvalidOperationException("Skipped cyclic or excessively nested templates menu: " + childName));
                    }
                    else
                    {
                        enteredChild = ancestors.Add(childName);
                        var children = await SnapshotAsync(child, beforeUiOperation, cancellationToken);
                        for (int i = children.Count - 1; i >= 0; i--)
                            await DeleteControlAsync(children[i], cleanupWarning, ancestors, depth + 1,
                                beforeUiOperation, cancellationToken);
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) { cleanupWarning?.Invoke(ex); }
                finally { if (enteredChild) ancestors.Remove(childName); }
            }
            await CheckpointAsync(beforeUiOperation, cancellationToken);
            try { control.Delete(); }
            catch (Exception ex) { cleanupWarning?.Invoke(ex); }
            if (child != null)
            {
                await CheckpointAsync(beforeUiOperation, cancellationToken);
                try { child.Delete(); }
                catch (Exception ex) { cleanupWarning?.Invoke(ex); }
            }
        }

        private static async Task CheckpointAsync(Func<Task> beforeUiOperation, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await beforeUiOperation();
            cancellationToken.ThrowIfCancellationRequested();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
        }
    }
}
