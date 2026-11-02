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

    internal sealed class QueryTemplateReconcileResult
    {
        internal int Loaded { get; set; }
        internal int Added { get; set; }
        internal int Removed { get; set; }
        internal int Repaired { get; set; }
        internal int Retained { get; set; }
        internal int Failures { get; set; }
        internal int AttachedHandlers { get; set; }
        internal bool HasChanges => Added != 0 || Removed != 0 || Repaired != 0 || AttachedHandlers != 0;
        internal ISet<string> ActiveCommandNames { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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
            return TryReadCommandIdentity(commands, control, out id) && id != null;
        }

        private static bool TryReadCommandIdentity(Commands commands, CommandBarControl control, out CommandID id)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            id = null;
            if (commands == null || control == null) return false;
            try
            {
                commands.CommandInfo(control, out string nativeGuid, out int nativeId);
                if (nativeId > 0 && Guid.TryParse(nativeGuid, out var group)) id = new CommandID(group, nativeId);
                // A successful query with no binding is different from a COM
                // lookup failure: only the latter leaves ownership uncertain.
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

        private sealed class NativeTemplateItem
        {
            internal CommandBarControl Control;
            internal CommandBar ChildBar;
            internal string ChildName;
            internal CommandID Identity;
            internal bool IdentityUnavailable;
        }

        private sealed class DesiredTemplateItem
        {
            internal QueryTemplateFolder Folder;
            internal QueryTemplateFile File;
            internal Aurora.TemplateCommandRegistration Registration;
            internal NativeTemplateItem Existing;
            internal bool RegistrationFailed;
            internal CommandBarControl NextRetained;
        }

        // Native insertion positions are one-based. Track changes in a small
        // prefix-sum tree so a large batch inserted before a retained neighbor
        // never rescans native siblings or repeatedly shifts an index table.
        private sealed class MenuPositions
        {
            private readonly Dictionary<CommandBarControl, int> _original = new Dictionary<CommandBarControl, int>();
            private readonly int[] _counts;
            private readonly int _appendSlot;
            internal MenuPositions(List<NativeTemplateItem> items)
            {
                _appendSlot = items.Count + 1;
                _counts = new int[items.Count + 2];
                for (int i = 0; i < items.Count; i++) { _original.Add(items[i].Control, i + 1); Add(i + 1, 1); }
            }
            private void Add(int index, int amount)
            { for (; index < _counts.Length; index += index & -index) _counts[index] += amount; }
            private int Sum(int index)
            { int sum = 0; for (; index > 0; index -= index & -index) sum += _counts[index]; return sum; }
            internal int Before(CommandBarControl next)
            { return next == null ? Sum(_appendSlot) + 1 : Sum(_original[next]); }
            internal void InsertBefore(CommandBarControl next) { Add(next == null ? _appendSlot : _original[next], 1); }
            internal void Remove(CommandBarControl control) { Add(_original[control], -1); }
        }

        internal static async Task<QueryTemplateReconcileResult> ReconcileAsync(QueryTemplateFolder folder,
            CommandBar root, Aurora.Plugin plugin, Aurora.CommandRegistry registry,
            Func<QueryTemplateFile, Aurora.CommandBase> createHandler, CommandID refreshCommand,
            CommandID openFolderCommand, List<string> warnings, Func<Task> beforeUiOperation,
            CancellationToken cancellationToken)
        {
            if (folder == null) throw new ArgumentNullException(nameof(folder));
            if (root == null) throw new ArgumentNullException(nameof(root));
            if (refreshCommand == null || openFolderCommand == null || refreshCommand.Equals(openFolderCommand))
                throw new ArgumentException("The two built-in template actions must have distinct command identities.");
            var result = new QueryTemplateReconcileResult();
            var protectedBars = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            CollectDesiredState(folder, protectedBars, result.ActiveCommandNames, registry);
            await CheckpointAsync(beforeUiOperation, cancellationToken);
            var native = await CaptureItemsAsync(root, plugin.Commands, beforeUiOperation, cancellationToken);
            bool refreshFound = false, openFound = false;
            var protectedControls = new HashSet<CommandBarControl>();
            foreach (var item in native)
            {
                if (refreshCommand.Equals(item.Identity)) { refreshFound = true; protectedControls.Add(item.Control); }
                if (openFolderCommand.Equals(item.Identity)) { openFound = true; protectedControls.Add(item.Control); }
            }
            // Validate both built-ins before changing commands or native controls.
            if (!refreshFound || !openFound)
                throw new InvalidOperationException("The Query Templates menu is not ready: its built-in actions could not be verified.");
            var ancestors = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { root.Name };
            await ReconcileFolderAsync(folder, root, native, protectedControls, protectedBars, ancestors,
                plugin, registry, createHandler, warnings, result, beforeUiOperation, cancellationToken, 0);
            return result;
        }

        private static void CollectDesiredState(QueryTemplateFolder folder, HashSet<string> names,
            ISet<string> activeCommands, Aurora.CommandRegistry registry)
        {
            // Cleanup uses the complete filesystem snapshot, even if a native
            // subtree fails to open or reconciliation is depth-limited later.
            foreach (var file in folder.Files) activeCommands.Add(file.CommandName);
            if (!folder.IsAvailable)
                foreach (var name in registry.GetCommandNamesUnderFolder(folder.FullPath)) activeCommands.Add(name);
            foreach (var child in folder.Folders)
            {
                names.Add(child.CommandName);
                CollectDesiredState(child, names, activeCommands, registry);
            }
        }

        private static async Task<List<NativeTemplateItem>> CaptureItemsAsync(CommandBar bar, Commands commands,
            Func<Task> beforeUiOperation, CancellationToken cancellationToken)
        {
            var controls = await SnapshotAsync(bar, beforeUiOperation, cancellationToken);
            var items = new List<NativeTemplateItem>(controls.Count);
            foreach (var control in controls)
            {
                await CheckpointAsync(beforeUiOperation, cancellationToken);
                var item = new NativeTemplateItem { Control = control };
                if (control is CommandBarPopup popup)
                {
                    item.ChildBar = popup.CommandBar;
                    item.ChildName = item.ChildBar.Name;
                }
                else if (commands != null)
                {
                    item.IdentityUnavailable = !TryReadCommandIdentity(commands, control, out var identity);
                    item.Identity = identity;
                }
                items.Add(item);
            }
            return items;
        }

        private static async Task<int> ReconcileFolderAsync(QueryTemplateFolder folder, CommandBar bar,
            List<NativeTemplateItem> native, HashSet<CommandBarControl> protectedControls,
            HashSet<string> protectedBars, HashSet<string> ancestors, Aurora.Plugin plugin,
            Aurora.CommandRegistry registry, Func<QueryTemplateFile, Aurora.CommandBase> createHandler,
            List<string> warnings, QueryTemplateReconcileResult result, Func<Task> beforeUiOperation,
            CancellationToken cancellationToken, int depth)
        {
            if (depth > MaximumMenuDepth)
            {
                result.Failures++;
                warnings.Add("Templates menu nesting exceeds the supported depth: " + folder.FullPath);
                foreach (var name in registry.GetCommandNamesUnderFolder(folder.FullPath)) result.ActiveCommandNames.Add(name);
                return 0;
            }
            var popups = new Dictionary<string, NativeTemplateItem>(StringComparer.OrdinalIgnoreCase);
            var buttons = new Dictionary<CommandID, NativeTemplateItem>();
            foreach (var item in native)
            {
                if (item.ChildName != null && !popups.ContainsKey(item.ChildName)) popups.Add(item.ChildName, item);
                else if (item.Control is CommandBarButton && item.Identity != null && !buttons.ContainsKey(item.Identity))
                    buttons.Add(item.Identity, item);
            }
            var desired = new List<DesiredTemplateItem>();
            var retained = new HashSet<CommandBarControl>(protectedControls);
            bool preserveUnknownButtons = false;
            bool hasUnidentifiedButtons = false;
            foreach (var item in native)
                if (item.Control is CommandBarButton && item.IdentityUnavailable) hasUnidentifiedButtons = true;
            if (hasUnidentifiedButtons)
            {
                result.Failures++;
                warnings.Add("Could not verify all template button identities in '" + folder.FullPath + "'; unidentified buttons were retained for retry.");
            }
            foreach (var child in folder.Folders)
            {
                popups.TryGetValue(child.CommandName, out var existing);
                desired.Add(new DesiredTemplateItem { Folder = child, Existing = existing });
                if (existing != null) retained.Add(existing.Control);
                if (!child.IsAvailable)
                    foreach (var name in registry.GetCommandNamesUnderFolder(child.FullPath)) result.ActiveCommandNames.Add(name);
            }
            foreach (var file in folder.Files)
            {
                result.ActiveCommandNames.Add(file.CommandName);
                var item = new DesiredTemplateItem { File = file };
                desired.Add(item);
                await CheckpointAsync(beforeUiOperation, cancellationToken);
                try
                {
                    bool wasCached = registry.IsTemplateCommandCached(file.CommandName);
                    item.Registration = registry.EnsureTemplateCommand(file.CommandName, file.FullPath, () => createHandler(file));
                    if (!wasCached) result.AttachedHandlers++;
                    buttons.TryGetValue(item.Registration.Identity, out item.Existing);
                    // A missing/misbound cached control may indicate that the
                    // host replaced its command. Re-resolve only affected items.
                    if (wasCached && item.Existing == null)
                    {
                        item.Registration = registry.EnsureTemplateCommand(file.CommandName, file.FullPath,
                            () => createHandler(file), forceLookup: true);
                        buttons.TryGetValue(item.Registration.Identity, out item.Existing);
                    }
                    if (item.Existing != null) retained.Add(item.Existing.Control);
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    item.RegistrationFailed = true;
                    preserveUnknownButtons = true;
                    result.Failures++;
                    warnings.Add("Could not register template '" + file.FullPath + "': " + ex.Message);
                }
            }

            // Failed command lookup leaves ownership uncertain. Keep unknown
            // buttons for a later retry, rather than deleting a working control.
            var positions = new MenuPositions(native);
            for (int index = native.Count - 1; index >= 0; index--)
            {
                var item = native[index];
                if (retained.Contains(item.Control) || (preserveUnknownButtons && item.Control is CommandBarButton)
                    || (item.Control is CommandBarButton && item.IdentityUnavailable)) continue;
                if (await DeleteObsoleteItemAsync(item, protectedBars, ancestors, warnings, result,
                    beforeUiOperation, cancellationToken, depth)) positions.Remove(item.Control);
            }

            CommandBarControl nextRetained = null;
            for (int i = desired.Count - 1; i >= 0; i--)
            {
                desired[i].NextRetained = nextRetained;
                if (desired[i].Existing != null) nextRetained = desired[i].Existing.Control;
            }
            int loaded = 0;
            foreach (var item in desired)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (item.RegistrationFailed || (item.File != null && item.Existing == null && hasUnidentifiedButtons)) continue;
                if (item.Folder != null && !item.Folder.IsAvailable)
                {
                    if (item.Existing != null) result.Retained++;
                    continue;
                }
                await CheckpointAsync(beforeUiOperation, cancellationToken);
                bool isNew = item.Existing == null;
                try
                {
                    if (item.Folder != null)
                    {
                        var child = item.Folder;
                        CommandBarPopup popup;
                        if (isNew)
                        {
                            int position = positions.Before(item.NextRetained);
                            popup = plugin.InsertCommandBarMenuPopup(child.CommandName, bar, position);
                            positions.InsertBefore(item.NextRetained);
                            item.Existing = new NativeTemplateItem { Control = popup, ChildBar = popup.CommandBar, ChildName = child.CommandName };
                            result.Added++;
                            if (popup.Enabled) popup.Enabled = false;
                        }
                        else { popup = (CommandBarPopup)item.Existing.Control; result.Retained++; }
                        bool repaired = false;
                        if (!string.Equals(popup.Caption, child.Name, StringComparison.Ordinal)) { popup.Caption = child.Name; repaired = true; }
                        if (!popup.Visible) { popup.Visible = true; repaired = true; }
                        if (!ancestors.Add(item.Existing.ChildName))
                            throw new InvalidOperationException("Skipped cyclic templates submenu: " + child.FullPath);
                        int childLoaded;
                        int failuresBeforeChild = result.Failures;
                        try
                        {
                            var childNative = await CaptureItemsAsync(item.Existing.ChildBar, plugin.Commands,
                                beforeUiOperation, cancellationToken);
                            childLoaded = await ReconcileFolderAsync(child, item.Existing.ChildBar, childNative,
                                new HashSet<CommandBarControl>(), protectedBars, ancestors, plugin, registry, createHandler,
                                warnings, result, beforeUiOperation, cancellationToken, depth + 1);
                        }
                        finally { ancestors.Remove(item.Existing.ChildName); }
                        loaded += childLoaded;
                        await CheckpointAsync(beforeUiOperation, cancellationToken);
                        // A partial branch has unknown contents. Preserve its
                        // current usability until the filesystem/native lookup
                        // succeeds; newly created branches start disabled.
                        bool incomplete = result.Failures > failuresBeforeChild || HasUnavailableFolder(child);
                        bool enabled = childLoaded > 0 || (incomplete && popup.Enabled);
                        if (popup.Enabled != enabled) { popup.Enabled = enabled; repaired = true; }
                        if (repaired && !isNew) result.Repaired++;
                    }
                    else
                    {
                        if (isNew)
                        {
                            int position = positions.Before(item.NextRetained);
                            var button = registry.InsertTemplateControl(item.Registration, bar, position);
                            positions.InsertBefore(item.NextRetained);
                            item.Existing = new NativeTemplateItem { Control = button, Identity = item.Registration.Identity };
                            result.Added++;
                            item.Registration.Handler.RepairTemplateControl(button, item.Registration.Command);
                        }
                        else
                        {
                            result.Retained++;
                            if (item.Registration.Handler.RepairTemplateControl((CommandBarButton)item.Existing.Control,
                                item.Registration.Command)) result.Repaired++;
                        }
                        loaded++;
                        result.Loaded++;
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    result.Failures++;
                    if (item.Folder != null)
                        foreach (var name in registry.GetCommandNamesUnderFolder(item.Folder.FullPath)) result.ActiveCommandNames.Add(name);
                    warnings.Add("Could not reconcile template '" + (item.Folder?.FullPath ?? item.File.FullPath) + "': " + ex.Message);
                }
            }
            return loaded;
        }

        private static bool HasUnavailableFolder(QueryTemplateFolder folder)
        {
            if (!folder.IsAvailable) return true;
            foreach (var child in folder.Folders) if (HasUnavailableFolder(child)) return true;
            return false;
        }

        private static async Task<bool> DeleteObsoleteItemAsync(NativeTemplateItem item,
            HashSet<string> protectedBars, HashSet<string> ancestors, List<string> warnings,
            QueryTemplateReconcileResult result, Func<Task> beforeUiOperation,
            CancellationToken cancellationToken, int depth)
        {
            CommandBar child = item.ChildBar;
            bool deleteChild = child != null && !protectedBars.Contains(item.ChildName)
                && !ancestors.Contains(item.ChildName) && depth < MaximumMenuDepth;
            if (deleteChild)
            {
                ancestors.Add(item.ChildName);
                try
                {
                    var children = await CaptureItemsAsync(child, null, beforeUiOperation, cancellationToken);
                    for (int i = children.Count - 1; i >= 0; i--)
                        await DeleteObsoleteItemAsync(children[i], protectedBars, ancestors, warnings,
                            result, beforeUiOperation, cancellationToken, depth + 1);
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) { result.Failures++; warnings.Add("Could not inspect obsolete templates submenu: " + ex.Message); }
                finally { ancestors.Remove(item.ChildName); }
            }
            await CheckpointAsync(beforeUiOperation, cancellationToken);
            try { item.Control.Delete(); result.Removed++; }
            catch (Exception ex)
            {
                result.Failures++;
                warnings.Add("Could not remove obsolete template menu item: " + ex.Message);
                return false;
            }
            if (deleteChild)
            {
                await CheckpointAsync(beforeUiOperation, cancellationToken);
                try { child.Delete(); }
                catch (Exception ex) { result.Failures++; warnings.Add("Could not remove obsolete templates submenu: " + ex.Message); }
            }
            return true;
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
