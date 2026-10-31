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

            foreach (CommandBarControl control in parent.Controls)
            {
                if (control is CommandBarPopup popup && string.Equals(popup.CommandBar.Name, name, StringComparison.Ordinal))
                    return popup.CommandBar;
            }

            // A failed/previous refresh may leave a named bar without a parent
            // control. Never reuse a global match under the wrong parent.
            CommandBar orphan = null;
            try { orphan = ((CommandBars)_app.CommandBars)[name]; }
            catch (System.Runtime.InteropServices.COMException) { }
            catch (ArgumentException) { }
            if (orphan != null) orphan.Delete();

            object newBar;
            ErrorHandler.ThrowOnFailure(_proffer.AddCommandBar(name, (uint)vsCommandBarType.vsCommandBarTypeMenu, parent,
                                  parent == null ? 0u : (uint)parent.Controls.Count,
                                  out newBar));
            var result = newBar as CommandBar;
            if (result == null)
                throw new InvalidOperationException("SSMS did not create the templates submenu: " + name);
            return result;
        }
    }

    //---- Base for each command handler
    public abstract class CommandBase
    {
        private readonly Plugin _plugin;
        private readonly string _name, _canonical, _tooltip;

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
        public virtual bool RegisterGUI(OleMenuCommand vsCommand, CommandBar bar, bool toolBarOnly)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (IconIndex >= 0 && toolBarOnly)
                RegisterGUIBar(vsCommand, bar);
            return true;
        }

        protected void RegisterGUIBar(OleMenuCommand vsCmd, CommandBar bar)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var registered = _plugin.Commands.Item(CanonicalName, -1);
            if (!Guid.TryParse(registered.Guid, out var registeredGroup)
                || registeredGroup != vsCmd.CommandID.Guid || registered.ID != vsCmd.CommandID.ID)
                throw new InvalidOperationException("The template command identity changed before its button was created: " + Name);

            CommandBarButton btn = null;
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

    //---- Registers commands and (re)adds GUI controls
    public class CommandRegistry
    {
        private readonly Plugin _plugin;
        private readonly CommandBar _toolbar;
        private readonly Guid _pkgGuid, _cmdSet;
        private readonly Dictionary<uint, CommandBase> _byId = new Dictionary<uint, CommandBase>();
        private readonly Dictionary<uint, OleMenuCommand> _commands = new Dictionary<uint, OleMenuCommand>();

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
            CommandBar menuParent = null)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // 1) Find or proffer the named command, then use the identity the
            // shell exposes instead of assuming a cached/generated ID is current.
            Command registered = null;
            try
            {
                registered = _plugin.Commands.Item(handler.CanonicalName, -1);
            }
            catch (System.Runtime.InteropServices.COMException) { }
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
            if (registered.ID <= 0)
                throw new InvalidOperationException("SSMS did not register template command " + handler.CanonicalName);
            uint cmdId = (uint)registered.ID;

            // 2) Get or create the OleMenuCommand
            var menuSvc = _plugin.MenuCommandService;
            var cmdID = new CommandID(commandGroup, registered.ID);
            var ole = menuSvc.FindCommand(cmdID) as OleMenuCommand;
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

            // 3) Always (re-)register the GUI control on the specified bar
            var targetBar = menuParent ?? _toolbar;
            handler.RegisterGUI(ole, targetBar, onlyToolbar);

            // 4) Optional keyboard bindings
            if (doBindings)
            {
                try
                {
                    var vsCmd = _plugin.Commands.Item(handler.CanonicalName, -1);
                    handler.BindToKeyboard(vsCmd);
                }
                catch { }
            }

        }

        public void RemoveUnusedCommands(ISet<string> activeNames)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Sweep persisted commands too. A previous SSMS session may have
            // registered a file that has since disappeared, or old Template_N IDs.
            var obsolete = new List<Command>();
            foreach (Command command in _plugin.Commands)
            {
                string name = command.Name;
                bool legacy = name != null && name.StartsWith("Template_", StringComparison.Ordinal)
                    && int.TryParse(name.Substring(9), out _);
                bool owned = name != null && name.StartsWith("AxialSqlTools_Template_", StringComparison.Ordinal);
                if ((owned || legacy) && !activeNames.Contains(name)
                    && Guid.TryParse(command.Guid, out var group) && group == _cmdSet)
                    obsolete.Add(command);
            }
            Exception cleanupError = null;
            foreach (var command in obsolete)
            {
                try { command.Delete(); }
                catch (Exception ex) { if (cleanupError == null) cleanupError = ex; }
            }

            foreach (uint id in new List<uint>(_byId.Keys))
            {
                if (activeNames.Contains(_byId[id].CanonicalName)) continue;
                if (_commands.TryGetValue(id, out var ole))
                    _plugin.MenuCommandService.RemoveCommand(ole);
                _commands.Remove(id);
                _byId.Remove(id);
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
