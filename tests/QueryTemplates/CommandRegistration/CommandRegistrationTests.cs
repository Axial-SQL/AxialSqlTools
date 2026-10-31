using Aurora;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Windows.Forms;

internal static partial class CommandRegistrationTests
{
    private static int _passed;

    private static int Main()
    {
        try
        {
            Run("Status can be queried synchronously during command/control publication", SynchronousStatus);
            Run("Initial command and button state is visible and supported", InitialVisibility);
            Run("Failed named-command HRESULT is reported", NamedCommandFailure);
            Run("A success HRESULT with zero command ID is rejected", ZeroCommandId);
            Run("Failed control HRESULT is reported and a retry succeeds", ControlFailureAndRetry);
            Run("A success HRESULT with no control is rejected", NullControl);
            Run("A command from another group is never reused", ForeignCommandCollision);
            Run("Refresh replaces handlers without duplicate controls or callbacks", HandlerReplacement);
            Run("Removing unused commands deletes bindings and disables stale callbacks", RemovedHandlers);
            Run("Cleanup leaves unrelated and untracked persisted commands untouched", PersistedCleanup);
            Run("Cleanup reports deletion failure while disabling obsolete handlers", CleanupFailure);
            Run("A missing VSCT root is not replaced by a detached menu", MissingRoot);
            Run("A submenu is reused only under the requested parent", ParentScopedSubmenu);
            Run("Submenu creation failures are reported", SubmenuFailures);
            RunMenuLifecycleTests();
            RunBoundedRegistrationTests();
            RunAsyncMenuTests();
            Console.WriteLine("Passed " + _passed + " command-registration behavioral tests (stub host).");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static void SynchronousStatus()
    {
        var fixture = new Fixture();
        var handler = fixture.Handler("script", false);
        int earlyQueries = 0;
        Action<OleMenuCommand> inspect = command =>
        {
            int previous = handler.StatusCalls;
            command.QueryStatus();
            Assert(handler.StatusCalls > previous, "Handler was not available during synchronous status query.");
            Assert(!command.Enabled && command.Visible && command.Supported,
                "Synchronous status did not reflect the actual handler.");
            earlyQueries++;
        };
        fixture.MenuService.WhileAdding = inspect;
        fixture.Proffer.WhileAddingControl = name => inspect(fixture.Command(name));
        fixture.Registry.RegisterCommand(false, handler, true);
        Equal(2, earlyQueries);
    }

    private static void InitialVisibility()
    {
        var fixture = new Fixture();
        fixture.Registry.RegisterCommand(false, fixture.Handler("script"), true);
        var command = fixture.Command("script");
        Assert(command.Enabled && command.Visible && command.Supported, "Command started hidden or unsupported.");
        var button = (CommandBarButton)fixture.Bar.Controls[1];
        Assert(button.Visible && button.Enabled, "Button started hidden or disabled.");
        Equal(string.Empty, button.Tag);
        Equal(command.CommandID.ID, button.NativeId);
        Equal(command.CommandID.Guid, Guid.Parse(button.NativeGuid));
        Equal("Caption script", button.Caption);
        Equal("Tooltip script", button.TooltipText);
    }

    private static void NamedCommandFailure()
    {
        var fixture = new Fixture();
        fixture.Proffer.NamedResult = FakeProffer.Failure;
        Throws<COMException>(() => fixture.Registry.RegisterCommand(false, fixture.Handler("script"), true));
        Equal(0, fixture.MenuService.Count);
        Equal(0, fixture.Proffer.ControlCalls);
    }

    private static void ZeroCommandId()
    {
        var fixture = new Fixture();
        fixture.Proffer.ReturnZeroId = true;
        Throws<InvalidOperationException>(() => fixture.Registry.RegisterCommand(false, fixture.Handler("script"), true));
        Equal(0, fixture.MenuService.Count);
        Equal(0, fixture.Proffer.ControlCalls);
    }

    private static void ControlFailureAndRetry()
    {
        var fixture = new Fixture();
        fixture.Proffer.ControlResult = FakeProffer.Failure;
        var failedHandler = fixture.Handler("script");
        Throws<COMException>(() => fixture.Registry.RegisterCommand(false, failedHandler, true));
        fixture.Proffer.ControlResult = 0;
        var retryHandler = fixture.Handler("script");
        fixture.Registry.RegisterCommand(false, retryHandler, true);
        fixture.Command("script").Invoke();
        Equal(0, failedHandler.Invocations);
        Equal(1, retryHandler.Invocations);
        Equal(1, fixture.App.Commands.Count);
        Equal(1, fixture.MenuService.Count);
        Equal(1, fixture.Bar.Controls.Count);
    }

    private static void NullControl()
    {
        var fixture = new Fixture();
        fixture.Proffer.ReturnNullControl = true;
        Throws<InvalidOperationException>(() => fixture.Registry.RegisterCommand(false, fixture.Handler("script"), true));
        Equal(0, fixture.Bar.Controls.Count);
    }

    private static void ForeignCommandCollision()
    {
        var fixture = new Fixture();
        fixture.App.Commands.Add("script", Guid.NewGuid(), 123);
        Throws<InvalidOperationException>(() => fixture.Registry.RegisterCommand(false, fixture.Handler("script"), true));
        Equal(0, fixture.Proffer.NamedCalls);
        Equal(0, fixture.Proffer.ControlCalls);
        Equal(0, fixture.MenuService.Count);
        Equal(1, fixture.App.Commands.Count);
    }

    private static void HandlerReplacement()
    {
        var fixture = new Fixture();
        var first = fixture.Handler("script");
        fixture.Registry.RegisterCommand(false, first, true);
        var command = fixture.Command("script");
        var replacement = fixture.Handler("script", false);
        fixture.Registry.RegisterCommand(false, replacement, true);
        Assert(ReferenceEquals(command, fixture.Command("script")), "Refresh unnecessarily replaced the Ole command.");
        command.QueryStatus();
        Assert(!command.Enabled, "Refresh kept the old handler's status.");
        command.Invoke();
        Equal(0, first.Invocations);
        Equal(1, replacement.Invocations);
        Equal(1, fixture.Proffer.NamedCalls);
        Equal(1, fixture.Proffer.ControlCalls);
        Equal(1, fixture.MenuService.Count);
    }

    private static void RemovedHandlers()
    {
        var fixture = new Fixture();
        const string removedName = "AxialSqlTools_Template_removed";
        const string keptName = "AxialSqlTools_Template_kept";
        var removed = fixture.Handler(removedName);
        var kept = fixture.Handler(keptName);
        fixture.Registry.RegisterCommand(false, removed, true);
        fixture.Registry.RegisterCommand(false, kept, true);
        var stale = fixture.Command(removedName);
        RemoveUnused(fixture, new HashSet<string> { keptName });
        Equal(1, fixture.App.Commands.DeletedCount);
        Equal(1, fixture.App.Commands.Count);
        Equal(1, fixture.MenuService.Count);
        stale.QueryStatus();
        Assert(!stale.Enabled && !stale.Visible && !stale.Supported, "Removed command stayed enabled.");
        stale.Invoke();
        Equal(0, removed.Invocations);
        fixture.Command(keptName).Invoke();
        Equal(1, kept.Invocations);
    }

    private static void PersistedCleanup()
    {
        var fixture = new Fixture();
        fixture.App.Commands.Add("AxialSqlTools_Template_obsolete", fixture.Group, 200);
        fixture.App.Commands.Add("Template_99", fixture.Group, 201);
        fixture.App.Commands.Add("AxialSqlTools_Template_foreign", Guid.NewGuid(), 202);
        fixture.App.Commands.Add("Template_helper", fixture.Group, 203);
        const string activeName = "AxialSqlTools_Template_active";
        fixture.Registry.RegisterCommand(false, fixture.Handler(activeName), true);
        fixture.App.Commands.ItemCalls = 0;
        RemoveUnused(fixture, new HashSet<string> { activeName });
        Equal(0, fixture.App.Commands.DeletedCount);
        Equal(5, fixture.App.Commands.Count);
        Equal(0, fixture.App.Commands.ItemCalls);
        Equal(0, fixture.App.Commands.EnumerationCalls);
        Equal(200, fixture.App.Commands.Item("AxialSqlTools_Template_obsolete", -1).ID);
        Equal(201, fixture.App.Commands.Item("Template_99", -1).ID);
        Equal(202, fixture.App.Commands.Item("AxialSqlTools_Template_foreign", -1).ID);
        Equal(203, fixture.App.Commands.Item("Template_helper", -1).ID);
        Equal(1, fixture.MenuService.Count);
    }

    private static void MissingRoot()
    {
        var fixture = new Fixture();
        Throws<COMException>(() => fixture.Plugin.AddCommandBarMenu("Query Templates", MsoBarPosition.msoBarTop, null));
        Equal(0, fixture.Proffer.BarCalls);
        var expected = new CommandBar();
        ((CommandBars)fixture.App.CommandBars).Add("Query Templates", expected);
        Assert(ReferenceEquals(expected, fixture.Plugin.AddCommandBarMenu("Query Templates", MsoBarPosition.msoBarTop, null)),
            "The registered root menu was not reused.");
    }

    private static void ParentScopedSubmenu()
    {
        var fixture = new Fixture();
        var unrelated = new CommandBar();
        ((CommandBars)fixture.App.CommandBars).Add("folder", unrelated);
        var ownChild = new CommandBar { Name = "folder" };
        fixture.Bar.Controls.Add(new CommandBarPopup { Tag = "folder", CommandBar = ownChild });
        Assert(ReferenceEquals(ownChild, fixture.Plugin.AddCommandBarMenu("folder", MsoBarPosition.msoBarMenuBar, fixture.Bar)),
            "Submenu lookup used a globally named bar instead of the requested parent.");
        Equal(0, fixture.Proffer.BarCalls);
        var otherParent = new CommandBar();
        var created = fixture.Plugin.AddCommandBarMenu("folder", MsoBarPosition.msoBarMenuBar, otherParent);
        Assert(!ReferenceEquals(created, ownChild) && !ReferenceEquals(created, unrelated), "Submenu was reused across parents.");
        Assert(unrelated.Deleted, "Orphan bar was not deleted before recreating it under the right parent.");
        Equal(1, fixture.Proffer.BarCalls);
    }

    private static void SubmenuFailures()
    {
        var fixture = new Fixture();
        fixture.Proffer.BarResult = FakeProffer.Failure;
        Throws<COMException>(() => fixture.Plugin.AddCommandBarMenu("folder", MsoBarPosition.msoBarMenuBar, fixture.Bar));
        fixture.Proffer.BarResult = 0;
        fixture.Proffer.ReturnNullBar = true;
        Throws<InvalidOperationException>(() => fixture.Plugin.AddCommandBarMenu("folder", MsoBarPosition.msoBarMenuBar, fixture.Bar));
    }

    private static void CleanupFailure()
    {
        var fixture = new Fixture();
        const string failed = "AxialSqlTools_Template_failed";
        const string removed = "AxialSqlTools_Template_removed";
        var handler = fixture.Handler(failed);
        fixture.Registry.RegisterCommand(false, handler, true);
        fixture.Registry.RegisterCommand(false, fixture.Handler(removed), true);
        var stale = fixture.Command(failed);
        fixture.App.Commands.FailDeleteName = failed;
        Throws<InvalidOperationException>(() => RemoveUnused(fixture, new HashSet<string>()));
        Equal(1, fixture.App.Commands.Count);
        Equal(1, fixture.App.Commands.DeletedCount);
        Equal(0, fixture.MenuService.Count);
        stale.QueryStatus();
        Assert(!stale.Enabled && !stale.Visible && !stale.Supported, "Failed shell deletion retained a callable handler.");
        stale.Invoke();
        Equal(0, handler.Invocations);
    }

    private static void Run(string name, Action action)
    {
        action();
        _passed++;
        Console.WriteLine("PASS: " + name);
    }

    private static void Equal<T>(T expected, T actual) => Assert(object.Equals(expected, actual), "Expected '" + expected + "', got '" + actual + "'.");
    private static void Assert(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
    private static void Throws<T>(Action action) where T : Exception
    {
        try { action(); }
        catch (T) { return; }
        throw new InvalidOperationException("Expected " + typeof(T).Name + ".");
    }

    private sealed class Fixture
    {
        public readonly Guid Group;
        public readonly DTE2 App;
        public readonly CommandBar Bar;
        public readonly OleMenuCommandService MenuService = new OleMenuCommandService();
        public readonly FakeProffer Proffer;
        public readonly Plugin Plugin;
        public readonly CommandRegistry Registry;
        public readonly List<Exception> Warnings = new List<Exception>();

        public Fixture(DTE2 app = null, CommandBar bar = null, Guid? group = null)
        {
            App = app ?? new DTE2();
            Bar = bar ?? new CommandBar();
            Group = group ?? Guid.NewGuid();
            App.Commands.MenuService = MenuService;
            Proffer = new FakeProffer(App);
            var icons = new ImageList();
            icons.Images.Add(new System.Drawing.Image());
            Plugin = new Plugin(App, Proffer, icons, MenuService, "AxialSqlTools", "Test", (ex, message) => Warnings.Add(ex));
            Registry = new CommandRegistry(Plugin, Bar, Guid.NewGuid(), Group);
        }

        public Handler Handler(string name, bool enabled = true) => new Handler(Plugin, name, enabled);

        public OleMenuCommand Command(string name)
        {
            var named = App.Commands.Item(name, -1);
            return (OleMenuCommand)MenuService.FindCommand(new CommandID(Group, named.ID));
        }
    }

    private sealed class Handler : CommandBase
    {
        private readonly bool _enabled;
        public int Invocations { get; private set; }
        public int StatusCalls { get; private set; }

        public Handler(Plugin plugin, string name, bool enabled)
            : base("Caption " + name, name, plugin, "Tooltip " + name) { _enabled = enabled; }

        public override bool OnCommand() { Invocations++; return true; }
        public override bool IsEnabled() { StatusCalls++; return _enabled; }
    }

    private sealed class FakeProffer : IVsProfferCommands3
    {
        public const int Failure = unchecked((int)0x80004005);
        private readonly DTE2 _app;
        public int NamedResult;
        public int ControlResult { get => _app.Commands.ControlResult; set => _app.Commands.ControlResult = value; }
        public int BarResult;
        public bool ReturnZeroId;
        public bool ReturnNullControl { get => _app.Commands.ReturnNullControl; set => _app.Commands.ReturnNullControl = value; }
        public bool ReturnNullBar;
        public int NamedCalls;
        public int ControlCalls => _app.Commands.ControlCalls;
        public int BarCalls;
        public Action<string> WhileAddingControl { get => _app.Commands.WhileAddingControl; set => _app.Commands.WhileAddingControl = value; }
        public int ReturnedIdOffset;
        public readonly Dictionary<CommandID, CommandBar> BarsByIdentity = new Dictionary<CommandID, CommandBar>();

        public FakeProffer(DTE2 app) { _app = app; }

        public int AddNamedCommand(ref Guid package, ref Guid group, string name, out uint id, string canonicalName,
            string buttonText, string tooltip, string binding, uint bitmapResource, uint bitmap,
            uint flags, uint contextCount, Guid[] contexts)
        {
            NamedCalls++;
            id = 0;
            if (NamedResult < 0 || ReturnZeroId) return NamedResult;
            int nativeId = _app.Commands.AllocateId();
            id = (uint)(nativeId + ReturnedIdOffset);
            _app.Commands.Add(name, group, nativeId);
            return NamedResult;
        }

        public int AddCommandBarControl(string name, object parent, uint index, uint type, out object control)
        {
            control = null;
            throw new InvalidOperationException("Template buttons must bind through the native Command.AddControl API.");
        }

        public int FindCommandBar(object owner, ref Guid group, uint menuId, out object bar)
        {
            BarsByIdentity.TryGetValue(new CommandID(group, (int)menuId), out var found);
            bar = found;
            return 0;
        }

        public int AddCommandBar(string name, uint type, object parent, uint index, out object bar)
        {
            BarCalls++;
            bar = null;
            if (BarResult < 0 || ReturnNullBar) return BarResult;
            var child = new CommandBar();
            if (parent is CommandBar parentBar)
                parentBar.Controls.Add(new CommandBarPopup { Caption = name, CommandBar = child });
            ((CommandBars)_app.CommandBars).Add(name, child);
            bar = child;
            return BarResult;
        }
    }
}
