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

        public DTE2 App => _app;
        public Commands Commands => _app.Commands;
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
            string connectPath)
        {
            _app = application;
            _proffer = profferCommands;
            _icons = icons;
            _menuService = menuService;
            _panelName = panelName;
            _connectPath = connectPath;
        }

        public CommandBar AddCommandBar(string name, MsoBarPosition pos)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var bars = (CommandBars)_app.CommandBars;
            try { return bars[name]; }
            catch { }

            object newBar;
            _proffer.AddCommandBar(name, (uint)vsCommandBarType.vsCommandBarTypeToolbar, null, 0, out newBar);
            return (CommandBar)newBar;
        }

        public CommandBar AddCommandBarMenu(string name, MsoBarPosition pos, CommandBar parent)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var bars = (CommandBars)_app.CommandBars;
            try { return bars[name]; }
            catch { }

            object newBar;
            _proffer.AddCommandBar(name, (uint)vsCommandBarType.vsCommandBarTypeMenu, parent,
                                  parent == null ? 0u : (uint)parent.Controls.Count,
                                  out newBar);
            return (CommandBar)newBar;
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
            if (IconIndex >= 0 && toolBarOnly)
                RegisterGUIBar(vsCommand, bar);
            return true;
        }

        protected void RegisterGUIBar(OleMenuCommand vsCmd, CommandBar bar)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CommandBarButton btn = null;
            try
            {
                var existing = bar.Controls[CanonicalName];
                btn = existing as CommandBarButton;
            }
            catch { }

            if (btn == null)
            {
                object newControl;
                _plugin.ProfferCommands.AddCommandBarControl(
                    CanonicalName,
                    bar,
                    (uint)bar.Controls.Count,
                    (uint)vsCommandBarType.vsCommandBarTypeToolbar,
                    out newControl);
                btn = newControl as CommandBarButton;
            }

            if (btn != null)
                AssignIcon(btn, ScriptIconIndex);
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

            // 1) Find or proffer named command
            uint cmdId = 0;
            try
            {
                var existing = _plugin.Commands.Item(handler.CanonicalName, -1);
                cmdId = (uint)existing.ID;
            }
            catch { }

            if (cmdId == 0)
            {
                _plugin.ProfferCommands.AddNamedCommand(
                    _pkgGuid,
                    _cmdSet,
                    handler.CanonicalName,
                    out cmdId,
                    handler.CanonicalName,
                    handler.Name,
                    handler.Tooltip,
                    null,
                    0,
                    (uint)handler.ScriptIconIndex,
                    0,
                    0,
                    null);
            }

            if (cmdId == 0)
                return;

            // 2) Get or create the OleMenuCommand
            var menuSvc = _plugin.MenuCommandService;
            var cmdID = new CommandID(_cmdSet, (int)cmdId);
            var ole = menuSvc.FindCommand(cmdID) as OleMenuCommand;
            if (ole == null)
            {
                ole = new OleMenuCommand((s, e) => Invoke(cmdId), cmdID);
                ole.BeforeQueryStatus += (s, e) => UpdateStatus(cmdId, ole);
                menuSvc.AddCommand(ole);
            }

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

            // 5) Update our lookup
            _byId[cmdId] = handler;
        }

        private void UpdateStatus(uint id, OleMenuCommand cmd)
        {
            if (_byId.TryGetValue(id, out var h))
            {
                cmd.Enabled = h.IsEnabled();
                cmd.Visible = true;
                cmd.Supported = true;
            }
        }

        private void Invoke(uint id)
        {
            if (_byId.TryGetValue(id, out var h))
                h.OnCommand();
        }
    }
}
