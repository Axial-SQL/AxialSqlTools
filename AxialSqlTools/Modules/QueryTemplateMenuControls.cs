using EnvDTE;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;

namespace AxialSqlTools
{
    internal static class QueryTemplateMenuControls
    {
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

        internal static void Clear(CommandBar root, Commands commands, CommandID refreshCommand,
            CommandID openFolderCommand, Action<Exception> cleanupWarning)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (root == null) throw new ArgumentNullException(nameof(root));
            if (commands == null) throw new ArgumentNullException(nameof(commands));
            if (refreshCommand == null) throw new ArgumentNullException(nameof(refreshCommand));
            if (openFolderCommand == null) throw new ArgumentNullException(nameof(openFolderCommand));
            if (refreshCommand.Equals(openFolderCommand))
                throw new ArgumentException("The two built-in template actions must have distinct command identities.");

            // Verify this is the expected VSCT menu before removing anything.
            // Tag is not supported by Visual Studio command bars and captions
            // may have changed or been localized since the previous session.
            // Classify once before mutation. Re-querying a verified built-in
            // during deletion could mistake a transient CommandInfo failure for
            // an unowned control and delete Refresh/Open Templates Folder.
            var dynamicControls = GetDynamicControls(Snapshot(root), commands, refreshCommand, openFolderCommand);
            for (int i = dynamicControls.Count - 1; i >= 0; i--)
                DeleteControl(dynamicControls[i], cleanupWarning);

            if (GetDynamicControls(Snapshot(root), commands, refreshCommand, openFolderCommand).Count != 0)
                throw new InvalidOperationException("SSMS could not remove an existing template menu item. Refresh was stopped to avoid duplicate scripts.");
        }

        internal static int Populate(QueryTemplateFolder folder, CommandBar bar, Aurora.Plugin plugin,
            Aurora.CommandRegistry registry, Func<QueryTemplateFile, Aurora.CommandBase> createHandler,
            List<string> warnings, ref int failures)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            int loaded = 0;
            var childMenus = new List<KeyValuePair<QueryTemplateFolder, CommandBar>>();
            foreach (var child in folder.Folders)
            {
                try
                {
                    var childBar = plugin.AddCommandBarMenu(child.CommandName, MsoBarPosition.msoBarMenuBar, bar);
                    // Locate the actual popup by its native bar name. Visual
                    // Studio does not use Tag and restores friendly captions.
                    bool found = false;
                    foreach (CommandBarControl control in bar.Controls)
                    {
                        if (!(control is CommandBarPopup popup)
                            || !string.Equals(popup.CommandBar.Name, childBar.Name, StringComparison.Ordinal)) continue;
                        control.Caption = child.Name;
                        found = true;
                        break;
                    }
                    if (!found) throw new InvalidOperationException("The submenu was not attached to its parent.");
                    childMenus.Add(new KeyValuePair<QueryTemplateFolder, CommandBar>(child, childBar));
                }
                catch (Exception ex)
                {
                    failures++;
                    warnings.Add("Could not create templates menu '" + child.FullPath + "': " + ex.Message);
                }
            }

            // Publish this folder's scripts before visiting descendants. A failed
            // child must never prevent the root scripts from being registered.
            foreach (var file in folder.Files)
            {
                try
                {
                    var handler = createHandler(file);
                    registry.RegisterCommand(false, handler, true, bar);
                    loaded++;
                }
                catch (Exception ex)
                {
                    failures++;
                    warnings.Add("Could not register template '" + file.FullPath + "': " + ex.Message);
                }
            }
            foreach (var child in childMenus)
                loaded += Populate(child.Key, child.Value, plugin, registry, createHandler, warnings, ref failures);
            return loaded;
        }

        private static List<CommandBarControl> GetDynamicControls(List<CommandBarControl> controls, Commands commands,
            CommandID refreshCommand, CommandID openFolderCommand)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool hasRefresh = false;
            bool hasOpenFolder = false;
            var dynamicControls = new List<CommandBarControl>();
            foreach (CommandBarControl control in controls)
            {
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

        private static List<CommandBarControl> Snapshot(CommandBar bar)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var controls = new List<CommandBarControl>();
            foreach (CommandBarControl control in bar.Controls) controls.Add(control);
            return controls;
        }

        private static void DeleteControl(CommandBarControl control, Action<Exception> cleanupWarning)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CommandBar child = null;
            if (control is CommandBarPopup popup)
            {
                try
                {
                    child = popup.CommandBar;
                    var children = Snapshot(child);
                    for (int i = children.Count - 1; i >= 0; i--)
                        DeleteControl(children[i], cleanupWarning);
                }
                catch (Exception ex) { cleanupWarning?.Invoke(ex); }
            }

            // Removing the popup and its named bar are independent operations.
            // A detached bar cleanup failure is cosmetic; a surviving root item
            // is caught by Clear's postcondition before any scripts are appended.
            try { control.Delete(); }
            catch (Exception ex) { cleanupWarning?.Invoke(ex); }
            if (child != null)
            {
                try { child.Delete(); }
                catch (Exception ex) { cleanupWarning?.Invoke(ex); }
            }
        }
    }
}
