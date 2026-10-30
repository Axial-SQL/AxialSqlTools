// These intentionally small API stand-ins exercise registration ordering and
// failures. They do not validate COM interop or replace an SSMS integration test.
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;

namespace EnvDTE
{
    public enum vsCommandBarType { vsCommandBarTypeMenu = 0, vsCommandBarTypeToolbar = 1 }

    public sealed class Command
    {
        public string Name { get; set; }
        public string Guid { get; set; }
        public int ID { get; set; }
        internal Action DeleteAction { get; set; }
        public void Delete() => DeleteAction();
    }

    public sealed class Commands
    {
        private readonly Dictionary<string, Command> _commands = new Dictionary<string, Command>();
        public int Count => _commands.Count;
        public int DeletedCount { get; private set; }
        public string FailDeleteName { get; set; }
        public IEnumerator<Command> GetEnumerator() => _commands.Values.GetEnumerator();

        public Command Item(string name, int id)
        {
            if (!_commands.TryGetValue(name, out var command))
                throw new COMException("Command not found.", unchecked((int)0x8002000B));
            return command;
        }

        public void Add(string name, Guid group, int id)
        {
            _commands.Add(name, new Command
            {
                Name = name,
                Guid = group.ToString("B"),
                ID = id,
                DeleteAction = () =>
                {
                    if (name == FailDeleteName) throw new COMException("Simulated command deletion failure.");
                    _commands.Remove(name);
                    DeletedCount++;
                }
            });
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
    public enum MsoButtonStyle { msoButtonIconAndCaption }

    public class CommandBarControl
    {
        public string Tag { get; set; }
        public string Caption { get; set; }
        public string TooltipText { get; set; }
        public bool Visible { get; set; }
        public bool Enabled { get; set; }
    }

    public sealed class CommandBarButton : CommandBarControl
    {
        public object Picture { get; set; }
        public MsoButtonStyle Style { get; set; }
    }

    public sealed class CommandBarPopup : CommandBarControl
    {
        public CommandBar CommandBar { get; set; }
    }

    public sealed class CommandBar
    {
        public List<CommandBarControl> Controls { get; } = new List<CommandBarControl>();
        internal Action DeleteAction { get; set; }
        public bool Deleted { get; private set; }
        public void Delete() { DeleteAction?.Invoke(); Deleted = true; }
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
        public static void ThrowIfNotOnUIThread() { }
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

        public MenuCommand FindCommand(CommandID id)
        {
            _commands.TryGetValue(id, out var command);
            return command;
        }

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
        int AddCommandBar(string name, uint type, object parent, uint index, out object bar);
        int AddCommandBarControl(string name, object parent, uint index, uint type, out object control);
        int AddNamedCommand(ref Guid package, ref Guid group, string name, out uint id, string canonicalName,
            string buttonText, string tooltip, string binding, uint bitmapResource, uint bitmap,
            uint flags, uint contextCount, Guid[] contexts);
    }
}

namespace stdole
{
    public sealed class StdPicture { }
}

namespace System.Drawing
{
    public class Image { }
}

namespace System.Windows.Forms
{
    public class AxHost
    {
        protected AxHost(string id) { }
        protected static object GetIPictureDispFromPicture(System.Drawing.Image image) => new stdole.StdPicture();
    }

    public sealed class ImageList
    {
        public List<System.Drawing.Image> Images { get; } = new List<System.Drawing.Image>();
    }
}
