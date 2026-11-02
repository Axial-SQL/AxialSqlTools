// AddinCore.cs
using System;
using System.IO;
using System.Xml.Serialization;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Windows.Forms;
using stdole;
using System.ComponentModel.Design;
using System.Collections.Generic;
using Microsoft.VisualStudio;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace Aurora
{
   
    //---- Simple singleton holder
    public sealed class Singleton<T> where T : class, new()
    {
        private Singleton() { }
        public static T Instance = new T();
    }

    //---- Thin wrap around VS command proffering
    public class Plugin
    {
        private readonly DTE2 _app;
        private readonly IVsProfferCommands3 _proffer;
        private readonly OleMenuCommandService _menuService;
        private readonly ImageList _icons;
        private readonly string _panelName;
        private readonly string _connectPath;
        private readonly Action<Exception, string> _warning;

        public DTE2 App => _app;
        public Commands Commands
        {
            get
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                return _app.Commands;
            }
        }
        public IVsProfferCommands3 ProfferCommands => _proffer;
        public OleMenuCommandService MenuCommandService => _menuService;
        public ImageList Icons => _icons;


        // If you need a prefix for command names, expose it here:
        public string Prefix => _panelName;

        public Plugin(
            DTE2 application,
            IVsProfferCommands3 profferCommands,
            ImageList icons,
            OleMenuCommandService menuService,
            string panelName,
            string connectPath,
            Action<Exception, string> warning = null)
        {
            _app = application;
            _proffer = profferCommands;
            _icons = icons;
            _menuService = menuService;
            _panelName = panelName;
            _connectPath = connectPath;
            _warning = warning;
        }

        internal void ReportWarning(Exception exception, string message) => _warning?.Invoke(exception, message);

        public CommandBar GetCommandBar(Guid commandGroup, int commandId)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            object result;
            ErrorHandler.ThrowOnFailure(_proffer.FindCommandBar(null, ref commandGroup, (uint)commandId, out result));
            return result as CommandBar
                ?? throw new InvalidOperationException("SSMS could not resolve the Query Templates menu by its command identity.");
        }

        public CommandBar AddCommandBar(string name, MsoBarPosition pos)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var bars = (CommandBars)_app.CommandBars;
            try { return bars[name]; }
            catch { }

            object newBar;
            ErrorHandler.ThrowOnFailure(_proffer.AddCommandBar(name, (uint)vsCommandBarType.vsCommandBarTypeToolbar, null, 0, out newBar));
            return (CommandBar)newBar;
        }

        public CommandBar AddCommandBarMenu(string name, MsoBarPosition pos, CommandBar parent)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parent == null)
                return ((CommandBars)_app.CommandBars)[name]; // The root is supplied by VSCT. Never create a detached replacement.

            return AddCommandBarMenuPopup(name, parent).CommandBar;
        }

        internal CommandBarPopup AddCommandBarMenuPopup(string name, CommandBar parent, bool appendOnly = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parent == null) throw new ArgumentNullException(nameof(parent));
            if (!appendOnly)
            {
                int existingCount = GetMenuControlCount(parent);
                for (int index = 1; index <= existingCount; index++)
                {
                    var existing = FindMatchingPopup(parent, index, name);
                    if (existing != null) return existing;
                }
            }

            // A failed/previous refresh may leave a named bar without a parent
            // control. Never reuse a global match under the wrong parent.
            CommandBar orphan = null;
            try { orphan = ((CommandBars)_app.CommandBars)[name]; }
            catch (System.Runtime.InteropServices.COMException) { }
            catch (ArgumentException) { }
            if (orphan != null) orphan.Delete();

            int insertionIndex = GetMenuControlCount(parent);
            object newBar;
            ErrorHandler.ThrowOnFailure(_proffer.AddCommandBar(name, (uint)vsCommandBarType.vsCommandBarTypeMenu, parent,
                                  (uint)insertionIndex,
                                  out newBar));
            var result = newBar as CommandBar;
            if (result == null)
                throw new InvalidOperationException("SSMS did not create the templates submenu: " + name);
            bool ownsReturnedBar = false;
            try
            {
                if (!string.Equals(result.Name, name, StringComparison.Ordinal))
                    throw new InvalidOperationException("SSMS returned a different templates submenu: " + name);
                ownsReturnedBar = true;
                if (GetMenuControlCount(parent) != insertionIndex + 1)
                    throw new InvalidOperationException("SSMS did not attach exactly one templates submenu: " + name);

                // Normal creation appends one popup. Verify the native child-bar
                // name before using the indexed control. The adjacent-slot probe
                // accommodates a host's insertion convention without scanning all
                // previously created siblings on every folder.
                var popup = FindMatchingPopup(parent, insertionIndex + 1, name);
                if (popup == null && insertionIndex > 0)
                    popup = FindMatchingPopup(parent, insertionIndex, name);
                if (popup == null)
                    throw new InvalidOperationException("SSMS attached the templates submenu at an unexpected location: " + name);
                return popup;
            }
            catch
            {
                if (ownsReturnedBar)
                {
                    try { result.Delete(); }
                    catch (Exception ex) { ReportWarning(ex, "Could not remove an unattached templates submenu '" + name + "'."); }
                }
                throw;
            }
        }

        internal CommandBarPopup InsertCommandBarMenuPopup(string name, CommandBar parent, int position)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            int count = GetMenuControlCount(parent);
            if (position < 1 || position > count + 1) throw new ArgumentOutOfRangeException(nameof(position));
            CommandBar orphan = null;
            try { orphan = ((CommandBars)_app.CommandBars)[name]; }
            catch (System.Runtime.InteropServices.COMException) { }
            catch (ArgumentException) { }
            if (orphan != null) orphan.Delete();

            var created = _app.Commands.AddCommandBar(name, vsCommandBarType.vsCommandBarTypeMenu,
                parent, position) as CommandBar;
            if (created == null || !string.Equals(created.Name, name, StringComparison.Ordinal))
                throw new InvalidOperationException("SSMS did not create the requested templates submenu: " + name);
            try
            {
                if (GetMenuControlCount(parent) != count + 1)
                    throw new InvalidOperationException("SSMS did not attach exactly one templates submenu: " + name);
                var popup = FindMatchingPopup(parent, position, name);
                if (popup == null)
                    throw new InvalidOperationException("SSMS attached the templates submenu at an unexpected location: " + name);
                return popup;
            }
            catch
            {
                try { created.Delete(); }
                catch (Exception ex) { ReportWarning(ex, "Could not remove an unattached templates submenu '" + name + "'."); }
                throw;
            }
        }

        private static int GetMenuControlCount(CommandBar parent)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            int count = parent.Controls.Count;
            if (count < 0 || count > 50000)
                throw new InvalidOperationException("The templates menu contains an unexpected number of controls.");
            return count;
        }

        private static CommandBarPopup FindMatchingPopup(CommandBar parent, int index, string name)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var popup = parent.Controls[index] as CommandBarPopup;
            return popup != null && string.Equals(popup.CommandBar.Name, name, StringComparison.Ordinal) ? popup : null;
        }
    }

    //---- Base for each command handler
    public abstract class CommandBase
    {
        private readonly Plugin _plugin;
        private readonly string _name, _canonical, _tooltip;
        private Command _resolvedGuiCommand;
        private CommandID _resolvedGuiCommandId;

        protected CommandBase(string name, string canonicalName, Plugin plugin, string tooltip)
        {
            _name = name;
            _canonical = canonicalName;
            _plugin = plugin;
            _tooltip = tooltip;
        }

        public string Name => _name;
        public string CanonicalName => _canonical;
        public string AbsName => _plugin.Prefix + "." + _canonical;
        public string Tooltip => _tooltip;

        public virtual int IconIndex => 0;
        public virtual int ScriptIconIndex => 0;

        // Called when registering GUI controls
        public virtual bool RegisterGUI(OleMenuCommand vsCommand, CommandBar bar, bool toolBarOnly, bool appendOnly = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (IconIndex >= 0 && toolBarOnly)
                RegisterGUIBar(vsCommand, bar, appendOnly);
            return true;
        }

        internal bool RegisterGUIWithResolvedCommand(OleMenuCommand vsCommand, CommandBar bar,
            bool toolBarOnly, bool appendOnly, Command registered)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var previousCommand = _resolvedGuiCommand;
            var previousId = _resolvedGuiCommandId;
            _resolvedGuiCommand = registered;
            _resolvedGuiCommandId = vsCommand.CommandID;
            try
            {
                // Keep virtual dispatch for existing command subclasses. The
                // resolved command is scoped to this call, never cached across
                // refreshes or retained if an override throws or reenters.
                return RegisterGUI(vsCommand, bar, toolBarOnly, appendOnly);
            }
            finally
            {
                _resolvedGuiCommand = previousCommand;
                _resolvedGuiCommandId = previousId;
            }
        }

        protected void RegisterGUIBar(OleMenuCommand vsCmd, CommandBar bar, bool appendOnly)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool hasResolvedCommand = _resolvedGuiCommand != null && vsCmd.CommandID.Equals(_resolvedGuiCommandId);
            var registered = hasResolvedCommand ? _resolvedGuiCommand : _plugin.Commands.Item(CanonicalName, -1);
            if (!hasResolvedCommand && (!Guid.TryParse(registered.Guid, out var registeredGroup)
                || registeredGroup != vsCmd.CommandID.Guid || registered.ID != vsCmd.CommandID.ID))
                throw new InvalidOperationException("The template command identity changed before its button was created: " + Name);

            CommandBarButton btn = null;
            // After the caller has cleared this menu, every script is new.
            // Scanning every earlier button would make a rebuild quadratic and
            // perform thousands of synchronous COM identity queries on the UI thread.
            if (!appendOnly)
            {
                foreach (CommandBarControl control in bar.Controls)
                {
                    if (control is CommandBarButton candidate
                        && AxialSqlTools.QueryTemplateMenuControls.TryGetCommandId(_plugin.Commands, candidate, out var identity)
                        && identity.Equals(vsCmd.CommandID))
                    {
                        btn = candidate;
                        break;
                    }
                }
            }

            if (btn == null)
            {
                // Bind through the shell's actual Command object. AddControl has
                // an explicit one-based position; no name/type-based routing.
                var created = registered.AddControl(bar, bar.Controls.Count + 1) as CommandBarControl;
                btn = created as CommandBarButton;
                if (btn == null)
                {
                    created?.Delete();
                    throw new InvalidOperationException("SSMS did not create the template button: " + Name);
                }
            }

            if (!AxialSqlTools.QueryTemplateMenuControls.TryGetCommandId(_plugin.Commands, btn, out var buttonId)
                || !buttonId.Equals(vsCmd.CommandID))
            {
                btn.Delete();
                throw new InvalidOperationException("SSMS attached the template button to the wrong command: " + Name);
            }

            // Tag is not used by Visual Studio and must never identify controls.
            btn.Caption = Name;
            btn.TooltipText = Tooltip;
            btn.Visible = true;
            btn.Enabled = vsCmd.Enabled;
            // Restored controls can lose their image. Reapply it on each refresh,
            // but a cosmetic COM failure must not disable a valid script command.
            try { AssignIcon(btn, ScriptIconIndex); }
            catch (Exception ex) { _plugin.ReportWarning(ex, "Could not set the icon for template '" + Name + "'."); }
        }

        internal bool RepairTemplateControl(CommandBarButton button, OleMenuCommand command)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool repaired = false;
            // Native controls can survive an SSMS restart. Preserve their COM
            // identities and leave correct properties entirely untouched.
            if (!string.Equals(button.Caption, Name, StringComparison.Ordinal)) { button.Caption = Name; repaired = true; }
            if (!string.Equals(button.TooltipText, Tooltip, StringComparison.Ordinal)) { button.TooltipText = Tooltip; repaired = true; }
            if (!button.Visible) { button.Visible = true; repaired = true; }
            if (button.Enabled != command.Enabled) { button.Enabled = command.Enabled; repaired = true; }
            try
            {
                var picture = button.Picture;
                if (picture == null || picture.Type <= 0)
                {
                    button.Picture = ImageConverter.LoadPictureFromImage(_plugin.Icons.Images[ScriptIconIndex]);
                    repaired = true;
                }
                if (button.Style != MsoButtonStyle.msoButtonIconAndCaption)
                {
                    button.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    repaired = true;
                }
            }
            catch (Exception ex) { _plugin.ReportWarning(ex, "Could not restore the icon for template '" + Name + "'."); }
            return repaired;
        }

        protected void AssignIcon(CommandBarButton target, int idx)
        {
            target.Picture = ImageConverter.LoadPictureFromImage(_plugin.Icons.Images[idx]);
            target.Style = MsoButtonStyle.msoButtonIconAndCaption;
        }

        public abstract bool OnCommand();
        public abstract bool IsEnabled();

        // Optionally override if you need keyboard binding
        public virtual void BindToKeyboard(Command vsCmd) { }
    }

    //---- Helper for converting Image to IPictureDisp
    class ImageConverter : AxHost
    {
        private ImageConverter() : base("{63109182-966B-4e3c-A8B2-8BC4A88D221C}") { }
        public static StdPicture LoadPictureFromImage(System.Drawing.Image i)
            => (StdPicture)GetIPictureDispFromPicture(i);
    }

    internal sealed class TemplateCommandRegistration
    {
        internal TemplateCommandRegistration(Command nativeCommand, OleMenuCommand command, CommandBase handler, string fullPath)
        { NativeCommand = nativeCommand; Command = command; Handler = handler; FullPath = fullPath; }
        internal Command NativeCommand { get; }
        internal OleMenuCommand Command { get; }
        internal CommandBase Handler { get; }
        internal string FullPath { get; }
        internal CommandID Identity => Command.CommandID;
    }

    //---- Registers commands and (re)adds GUI controls
    public class CommandRegistry
    {
        private readonly Plugin _plugin;
        private readonly CommandBar _toolbar;
        private readonly Guid _pkgGuid, _cmdSet;
        private readonly Dictionary<uint, CommandBase> _byId = new Dictionary<uint, CommandBase>();
        private readonly Dictionary<uint, OleMenuCommand> _commands = new Dictionary<uint, OleMenuCommand>();
        private readonly Dictionary<string, TemplateCommandRegistration> _registrations =
            new Dictionary<string, TemplateCommandRegistration>(StringComparer.OrdinalIgnoreCase);

        public CommandRegistry(Plugin plugin, CommandBar toolbar, Guid packageGuid, Guid cmdGroupGuid)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _plugin = plugin;
            _toolbar = toolbar;
            _pkgGuid = packageGuid;
            _cmdSet = cmdGroupGuid;
        }

        public void RegisterCommand(
            bool doBindings,
            CommandBase handler,
            bool onlyToolbar,
            CommandBar menuParent = null,
            bool appendOnly = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var registration = RegisterNativeCommand(handler, null);
            var registered = registration.NativeCommand;
            var ole = registration.Command;

            // 3) Always (re-)register the GUI control on the specified bar
            var targetBar = menuParent ?? _toolbar;
            handler.RegisterGUIWithResolvedCommand(ole, targetBar, onlyToolbar, appendOnly, registered);

            // 4) Optional keyboard bindings
            if (doBindings)
            {
                try
                {
                    handler.BindToKeyboard(registered);
                }
                catch { }
            }

        }

        internal bool IsTemplateCommandCached(string canonicalName) => _registrations.ContainsKey(canonicalName);

        internal TemplateCommandRegistration EnsureTemplateCommand(string canonicalName, string fullPath,
            Func<CommandBase> createHandler, bool forceLookup = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!forceLookup && _registrations.TryGetValue(canonicalName, out var cached))
            {
                if (!string.Equals(cached.FullPath, fullPath, StringComparison.Ordinal))
                {
                    // Path hashes are case insensitive, but a case-only rename
                    // still updates the visible name and the path used on click.
                    var replacement = createHandler();
                    cached = new TemplateCommandRegistration(cached.NativeCommand, cached.Command, replacement, fullPath);
                    _registrations[canonicalName] = cached;
                    _byId[(uint)cached.Identity.ID] = replacement;
                }
                UpdateStatus((uint)cached.Identity.ID, cached.Command);
                return cached;
            }
            var handler = _registrations.TryGetValue(canonicalName, out cached)
                && string.Equals(cached.FullPath, fullPath, StringComparison.Ordinal) ? cached.Handler : createHandler();
            return RegisterNativeCommand(handler, fullPath);
        }

        internal IEnumerable<string> GetCommandNamesUnderFolder(string fullPath)
        {
            string prefix = fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            foreach (var entry in _registrations)
                if (entry.Value.FullPath != null && entry.Value.FullPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    yield return entry.Key;
        }

        internal CommandBarButton InsertTemplateControl(TemplateCommandRegistration registration, CommandBar bar, int position)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var created = registration.NativeCommand.AddControl(bar, position) as CommandBarControl;
            var button = created as CommandBarButton;
            if (button == null || !AxialSqlTools.QueryTemplateMenuControls.TryGetCommandId(_plugin.Commands, button, out var id)
                || !registration.Identity.Equals(id))
            {
                created?.Delete();
                throw new InvalidOperationException("SSMS attached the template button to the wrong command: " + registration.Handler.Name);
            }
            return button;
        }

        private TemplateCommandRegistration RegisterNativeCommand(CommandBase handler, string fullPath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // 1) Find or proffer the named command, then use the identity the
            // shell exposes instead of assuming a cached/generated ID is current.
            Command registered = null;
            try
            {
                registered = _plugin.Commands.Item(handler.CanonicalName, -1);
            }
            catch (System.Runtime.InteropServices.COMException ex) when (
                ex.ErrorCode == unchecked((int)0x8002000B) || ex.ErrorCode == unchecked((int)0x80020003)) { }
            catch (ArgumentException) { }

            if (registered == null)
            {
                uint profferedId;
                ErrorHandler.ThrowOnFailure(_plugin.ProfferCommands.AddNamedCommand(
                    _pkgGuid,
                    _cmdSet,
                    handler.CanonicalName,
                    out profferedId,
                    handler.CanonicalName,
                    handler.Name,
                    handler.Tooltip,
                    null,
                    0,
                    (uint)handler.ScriptIconIndex,
                    0,
                    0,
                    null));
                if (profferedId == 0)
                    throw new InvalidOperationException("SSMS did not register template command " + handler.CanonicalName);
                registered = _plugin.Commands.Item(handler.CanonicalName, -1);
            }

            if (!Guid.TryParse(registered.Guid, out var commandGroup) || commandGroup != _cmdSet)
                throw new InvalidOperationException("A different command group owns command " + handler.CanonicalName);
            int registeredId = registered.ID;
            if (registeredId <= 0)
                throw new InvalidOperationException("SSMS did not register template command " + handler.CanonicalName);
            uint cmdId = (uint)registeredId;

            if (_registrations.TryGetValue(handler.CanonicalName, out var previous)
                && previous.Identity.ID != registeredId)
            {
                _plugin.MenuCommandService.RemoveCommand(previous.Command);
                _commands.Remove((uint)previous.Identity.ID);
                _byId.Remove((uint)previous.Identity.ID);
            }

            // 2) Get or create the OleMenuCommand
            var menuSvc = _plugin.MenuCommandService;
            var cmdID = new CommandID(commandGroup, registeredId);
            var ole = menuSvc.FindCommand(cmdID) as OleMenuCommand;
            if (ole != null && (!_commands.TryGetValue(cmdId, out var ownedOle) || !ReferenceEquals(ole, ownedOle)))
            {
                // A new registry must own its callbacks, even when the host
                // retained a previous package instance's command service entry.
                menuSvc.RemoveCommand(ole);
                ole = null;
            }
            // SSMS can query status while AddCommand or Command.AddControl runs.
            // Publish the handler before making its control visible.
            _byId[cmdId] = handler;
            if (ole == null)
            {
                ole = new OleMenuCommand((s, e) => Invoke(cmdId), cmdID);
                ole.BeforeQueryStatus += (s, e) => UpdateStatus(cmdId, ole);
                UpdateStatus(cmdId, ole);
                menuSvc.AddCommand(ole);
            }
            UpdateStatus(cmdId, ole);
            _commands[cmdId] = ole;
            var registration = new TemplateCommandRegistration(registered, ole, handler, fullPath);
            _registrations[handler.CanonicalName] = registration;
            return registration;

        }

        public async Task RemoveUnusedCommandsAsync(ISet<string> activeNames, Func<Task> beforeUiOperation,
            CancellationToken cancellationToken)
        {
            if (activeNames == null) throw new ArgumentNullException(nameof(activeNames));
            if (beforeUiOperation == null) throw new ArgumentNullException(nameof(beforeUiOperation));
            cancellationToken.ThrowIfCancellationRequested();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // Never enumerate DTE.Commands: it includes every SSMS command and
            // can trigger expensive shell/package work during startup. Cleanup
            // is bounded by the commands this registry actually registered.
            // Untracked persisted names are harmless: menu cleanup removes their
            // visible controls, and a later load can reuse an exact named command.
            Exception cleanupError = null;
            foreach (var entry in new List<KeyValuePair<uint, CommandBase>>(_byId))
            {
                uint id = entry.Key;
                string name = entry.Value.CanonicalName;
                if (activeNames.Contains(name)) continue;

                // Await outside the cleanup try/finally: cancellation must not
                // detach a command whose cleanup has not started yet.
                cancellationToken.ThrowIfCancellationRequested();
                await beforeUiOperation();
                cancellationToken.ThrowIfCancellationRequested();
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

                try
                {
                    Command command = null;
                    try { command = _plugin.Commands.Item(name, -1); }
                    catch (System.Runtime.InteropServices.COMException ex) when (
                        ex.ErrorCode == unchecked((int)0x8002000B) // DISP_E_BADINDEX: no such named command.
                        || ex.ErrorCode == unchecked((int)0x80020003)) // DISP_E_MEMBERNOTFOUND.
                    {
                        // The shell has already removed this command.
                    }

                    if (command != null)
                    {
                        if (!Guid.TryParse(command.Guid, out var group) || group != _cmdSet
                            || command.ID <= 0 || (uint)command.ID != id)
                            throw new InvalidOperationException("The identity of obsolete template command '" + name + "' changed; it was not deleted.");
                        command.Delete();
                    }
                }
                catch (Exception ex) { if (cleanupError == null) cleanupError = ex; }
                finally
                {
                    // Even a shell lookup/delete failure must detach the old
                    // callback, so a stale control cannot invoke an obsolete file.
                    if (_commands.TryGetValue(id, out var ole))
                    {
                        try { _plugin.MenuCommandService.RemoveCommand(ole); }
                        catch (Exception ex) { if (cleanupError == null) cleanupError = ex; }
                    }
                    _commands.Remove(id);
                    _byId.Remove(id);
                    _registrations.Remove(name);
                }
            }
            if (cleanupError != null)
                throw new InvalidOperationException("SSMS could not remove an obsolete template command.", cleanupError);
        }

        private void UpdateStatus(uint id, OleMenuCommand cmd)
        {
            if (_byId.TryGetValue(id, out var h))
            {
                cmd.Enabled = h.IsEnabled();
                cmd.Visible = true;
                cmd.Supported = true;
            }
            else
            {
                cmd.Enabled = false;
                cmd.Visible = false;
                cmd.Supported = false;
            }
        }

        private void Invoke(uint id)
        {
            if (_byId.TryGetValue(id, out var h))
                h.OnCommand();
        }
    }
}
