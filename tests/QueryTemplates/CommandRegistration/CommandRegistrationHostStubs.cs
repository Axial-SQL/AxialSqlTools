// Behavioral stand-ins for documented VS automation semantics. In particular,
// Tag is not stored, native control indexes are one-based, and button clicks
// dispatch by the GUID/ID installed by Command.AddControl, not by their caption.
// These tests do not validate COM interop or replace an SSMS integration test.
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;

namespace EnvDTE
{
    public enum vsCommandBarType { vsCommandBarTypeMenu = 0, vsCommandBarTypeToolbar = 1 }

    public sealed class Command
    {
        public string Name { get; set; }
        public string Guid { get; set; }
        public int ID { get; set; }
        internal Commands Owner { get; set; }
        internal Action DeleteAction { get; set; }
        public void Delete() => DeleteAction();
        public object AddControl(object owner, int position) => Owner.AddControl(this, (CommandBar)owner, position);
    }

    public sealed class Commands : IEnumerable<Command>
    {
        private readonly Dictionary<string, Command> _commands = new Dictionary<string, Command>();
        public int Count => _commands.Count;
        public int DeletedCount { get; private set; }
        public string FailDeleteName { get; set; }
        public OleMenuCommandService MenuService { get; set; }
        public int ControlCalls { get; private set; }
        public int ControlResult { get; set; }
        public bool ReturnNullControl { get; set; }
        public bool ReturnWrongBinding { get; set; }
        public bool FailButtonPicture { get; set; }
        public Action<string> WhileAddingControl { get; set; }
        public int EnumerationCalls { get; private set; }
        public int ItemCalls { get; set; }
        public int CommandInfoCallCount { get; set; }
        public string FailLookupName { get; set; }
        public int LookupFailureHResult { get; set; } = unchecked((int)0x80004005);
        public IEnumerator<Command> GetEnumerator()
        {
            EnumerationCalls++;
            throw new InvalidOperationException("Enumerating the global DTE.Commands collection is forbidden during template operations.");
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public Command Item(string name, int id)
        {
            ItemCalls++;
            if (name == FailLookupName) throw new COMException("Simulated command lookup failure.", LookupFailureHResult);
            if (!_commands.TryGetValue(name, out var command))
                throw new COMException("Command not found.", unchecked((int)0x8002000B));
            return command;
        }

        // Test-only inspection does not expose the shell's global enumerator.
        internal Command GetByIdentity(string guid, int id)
        {
            foreach (var command in _commands.Values)
                if (command.Guid == guid && command.ID == id) return command;
            throw new InvalidOperationException("No command has this native identity.");
        }

        public int AllocateId()
        {
            int result = 100;
            foreach (var command in _commands.Values) result = Math.Max(result, command.ID + 1);
            return result;
        }

        public void Add(string name, Guid group, int id)
        {
            _commands.Add(name, new Command
            {
                Name = name,
                Guid = group.ToString("B"),
                ID = id,
                Owner = this,
                DeleteAction = () =>
                {
                    if (name == FailDeleteName) throw new COMException("Simulated command deletion failure.");
                    _commands.Remove(name);
                    DeletedCount++;
                }
            });
        }

        internal object AddControl(Command command, CommandBar parent, int position)
        {
            ControlCalls++;
            WhileAddingControl?.Invoke(command.Name);
            if (ControlResult < 0) throw new COMException("Simulated native AddControl failure.", ControlResult);
            if (ReturnNullControl) return null;
            var button = new CommandBarButton
            {
                NativeGuid = command.Guid,
                NativeId = ReturnWrongBinding ? command.ID + 10000 : command.ID,
                FailPicture = FailButtonPicture,
                Dispatch = control =>
                {
                    var id = new CommandID(System.Guid.Parse(control.NativeGuid), control.NativeId);
                    var target = MenuService.FindCommand(id) as OleMenuCommand;
                    if (target == null) throw new InvalidOperationException("Button has no registered native command handler.");
                    target.QueryStatus();
                    if (target.Enabled && target.Visible && target.Supported) target.Invoke();
                }
            };
            parent.Controls.Insert(position, button);
            return button;
        }

        public void CommandInfo(object control, out string guid, out int id)
        {
            CommandInfoCallCount++;
            var native = control as CommandBarControl;
            if (native != null) native.CommandInfoCalls++;
            if (native == null || native.FailCommandInfo || string.IsNullOrEmpty(native.NativeGuid)
                || (native.CommandInfoSuccessLimit > 0 && native.CommandInfoCalls > native.CommandInfoSuccessLimit))
                throw new COMException("Control has no native command identity.");
            guid = native.NativeGuid;
            id = native.NativeId;
        }
    }
}

namespace EnvDTE80
{
    public sealed class DTE2
    {
        public EnvDTE.Commands Commands { get; } = new EnvDTE.Commands();
        public object CommandBars { get; } = new Microsoft.VisualStudio.CommandBars.CommandBars();
    }
}

namespace Microsoft.VisualStudio.CommandBars
{
    public enum MsoBarPosition { msoBarTop, msoBarMenuBar }
    public enum MsoButtonStyle { msoButtonAutomatic, msoButtonIconAndCaption }

    public class CommandBarControl
    {
        // Documented Visual Studio behavior: this Office compatibility property
        // is not used. Persisting it here would conceal the real regression.
        public string Tag { get => string.Empty; set { if (ThrowOnTagWrite) throw new COMException("Tag is unsupported."); } }
        public bool ThrowOnTagWrite { get; set; }
        public string Caption { get; set; }
        public string TooltipText { get; set; }
        public bool Visible { get; set; }
        public bool Enabled { get; set; }
        public string NativeGuid { get; set; }
        public int NativeId { get; set; }
        public bool FailCommandInfo { get; set; }
        public int CommandInfoCalls { get; set; }
        public int CommandInfoSuccessLimit { get; set; }
        public bool FailDelete { get; set; }
        public bool Deleted { get; private set; }
        internal CommandBarControls Owner { get; set; }
        public void Delete(bool temporary = false)
        {
            if (FailDelete) throw new COMException("Simulated control deletion failure.");
            Owner?.Remove(this);
            Deleted = true;
        }
    }

    public sealed class CommandBarButton : CommandBarControl
    {
        private object _picture;
        public bool FailPicture { get; set; }
        public object Picture
        {
            get => _picture;
            set { if (FailPicture) throw new COMException("Simulated picture assignment failure."); _picture = value; }
        }
        public MsoButtonStyle Style { get; set; }
        public Action<CommandBarButton> Dispatch { get; set; }
        public void Click()
        {
            if (Deleted) throw new InvalidOperationException("Cannot click a deleted control.");
            Dispatch?.Invoke(this);
        }
    }

    public sealed class CommandBarPopup : CommandBarControl
    {
        public CommandBar CommandBar { get; set; }
    }

    public sealed class CommandBarControls : IEnumerable<CommandBarControl>
    {
        private readonly List<CommandBarControl> _items = new List<CommandBarControl>();
        public bool ThrowOnEnumeration { get; set; }
        public int EnumerationCalls { get; private set; }
        public int Count => _items.Count;
        public CommandBarControl this[int index]
        {
            get
            {
                if (index < 1 || index > _items.Count) throw new COMException("Native control index is one-based.");
                return _items[index - 1];
            }
        }
        public void Add(CommandBarControl control) => Insert(Count + 1, control);
        public void Insert(int position, CommandBarControl control)
        {
            if (position < 1 || position > _items.Count + 1) throw new COMException("Native insertion position is one-based.");
            _items.Insert(position - 1, control);
            control.Owner = this;
        }
        internal void Remove(CommandBarControl control) => _items.Remove(control);
        public IEnumerator<CommandBarControl> GetEnumerator()
        {
            EnumerationCalls++;
            if (ThrowOnEnumeration) throw new InvalidOperationException("Native COM control enumeration is forbidden in this test.");
            return _items.GetEnumerator();
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    public sealed class CommandBar
    {
        public string Name { get; set; }
        public CommandBarControls Controls { get; } = new CommandBarControls();
        internal Action DeleteAction { get; set; }
        public bool Deleted { get; private set; }
        public bool FailDelete { get; set; }
        public void Delete()
        {
            if (FailDelete) throw new COMException("Simulated named bar deletion failure.");
            DeleteAction?.Invoke();
            Deleted = true;
        }
    }

    public sealed class CommandBars
    {
        private readonly Dictionary<string, CommandBar> _bars = new Dictionary<string, CommandBar>();
        public CommandBar this[string name]
        {
            get
            {
                if (!_bars.TryGetValue(name, out var bar))
                    throw new COMException("Command bar not found.", unchecked((int)0x8002000B));
                return bar;
            }
        }
        public void Add(string name, CommandBar bar)
        {
            _bars.Add(name, bar);
            bar.Name = name;
            bar.DeleteAction = () => _bars.Remove(name);
        }
    }
}

namespace Microsoft.VisualStudio
{
    public static class ErrorHandler
    {
        public static int ThrowOnFailure(int result)
        {
            if (result < 0) throw new COMException("Simulated host failure.", result);
            return result;
        }
    }
}

namespace Microsoft.VisualStudio.Shell
{
    public static class ThreadHelper
    {
        public static TestJoinableTaskFactory JoinableTaskFactory { get; } = new TestJoinableTaskFactory();
        public static void ThrowIfNotOnUIThread() { }
    }

    public sealed class TestJoinableTaskFactory
    {
        public System.Threading.Tasks.Task SwitchToMainThreadAsync(System.Threading.CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return System.Threading.Tasks.Task.CompletedTask;
        }
    }

    public sealed class OleMenuCommand : MenuCommand
    {
        public OleMenuCommand(EventHandler invoke, CommandID id) : base(invoke, id) { }
        public event EventHandler BeforeQueryStatus;
        public void QueryStatus() => BeforeQueryStatus?.Invoke(this, EventArgs.Empty);
    }

    public sealed class OleMenuCommandService
    {
        private readonly Dictionary<CommandID, MenuCommand> _commands = new Dictionary<CommandID, MenuCommand>();
        public Action<OleMenuCommand> WhileAdding { get; set; }
        public int Count => _commands.Count;
        public MenuCommand FindCommand(CommandID id) { _commands.TryGetValue(id, out var command); return command; }
        public void AddCommand(MenuCommand command)
        {
            _commands.Add(command.CommandID, command);
            WhileAdding?.Invoke((OleMenuCommand)command);
        }
        public void RemoveCommand(MenuCommand command) => _commands.Remove(command.CommandID);
    }
}

namespace Microsoft.VisualStudio.Shell.Interop
{
    [ComImport]
    [Guid("3A83904D-4540-4C51-95A7-618B32A9A9C0")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IVsProfferCommands3
    {
        int FindCommandBar(object owner, ref Guid group, uint menuId, out object bar);
        int AddCommandBar(string name, uint type, object parent, uint index, out object bar);
        int AddCommandBarControl(string name, object parent, uint index, uint type, out object control);
        int AddNamedCommand(ref Guid package, ref Guid group, string name, out uint id, string canonicalName,
            string buttonText, string tooltip, string binding, uint bitmapResource, uint bitmap,
            uint flags, uint contextCount, Guid[] contexts);
    }
}

namespace stdole { public sealed class StdPicture { } }
namespace System.Drawing { public class Image { } }
namespace System.Windows.Forms
{
    public class AxHost
    {
        protected AxHost(string id) { }
        protected static object GetIPictureDispFromPicture(System.Drawing.Image image) => new stdole.StdPicture();
    }
    public sealed class ImageList { public List<System.Drawing.Image> Images { get; } = new List<System.Drawing.Image>(); }
}
