using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

// Only the package host is substituted. The linked Templates partial, discovery,
// coordinator, UI gate, batcher, menu builder, and command registry are production code.
internal sealed class PackageTestHost : IDisposable
{
    private readonly int _uiThreadId = Thread.CurrentThread.ManagedThreadId;
    internal readonly NormalOnlyScheduler Scheduler = new NormalOnlyScheduler();
    internal readonly CancellationTokenSource Disposal = new CancellationTokenSource();
    internal readonly DTE2 App = new DTE2();
    internal readonly OleMenuCommandService MenuService = new OleMenuCommandService();
    internal readonly CommandBar Root = new CommandBar { Name = "Query Templates" };
    internal readonly List<OleMenuCommand> PublishedCommands = new List<OleMenuCommand>();
    internal readonly List<string> ExecutedFiles = new List<string>();
    internal readonly List<int> UpdateRequests = new List<int>();
    internal readonly List<string> Logs = new List<string>();
    internal readonly PackageShell Shell;
    internal readonly AxialSqlTools.AxialSqlToolsPackage Package;
    internal readonly Dictionary<Type, object> Services = new Dictionary<Type, object>();
    internal int NativeMutations;
    internal int UnsafeNativeMutations;
    internal int StatusQueries;
    internal int MenuOpenEvents { get; private set; }
    internal string LastStatus;

    internal PackageTestHost(string folder)
    {
        Shell = new PackageShell(this);
        App.Commands.MenuService = MenuService;
        App.Commands.Add("Test.Refresh", AxialSqlTools.RefreshTemplatesCommand.CommandSet,
            AxialSqlTools.RefreshTemplatesCommand.CommandId);
        App.Commands.Add("Test.OpenFolder", AxialSqlTools.OpenTemplatesFolderCommand.CommandSet,
            AxialSqlTools.OpenTemplatesFolderCommand.CommandId);
        App.Commands.Item("Test.Refresh", -1).AddControl(Root, 1);
        App.Commands.Item("Test.OpenFolder", -1).AddControl(Root, 2);
        App.Commands.WhileAddingControl = name => RecordMutation();
        MenuService.WhileAdding = command => { RecordMutation(); PublishedCommands.Add(command); };
        Services[typeof(EnvDTE.DTE)] = App;
        Services[typeof(SVsShell)] = Shell;
        Services[typeof(SVsUIShell)] = Shell;
        Services[typeof(SVsStatusbar)] = Shell;
        Services[typeof(SVsProfferCommands)] = new PackageProffer(this);
        Services[typeof(IMenuCommandService)] = MenuService;
        ThreadHelper.JoinableTaskFactory.Scheduler = Scheduler;
        AxialSqlTools.SettingsManager.Folder = folder;
        Package = new AxialSqlTools.AxialSqlToolsPackage(this);
    }

    internal void RecordMutation()
    {
        NativeMutations++;
        if (Thread.CurrentThread.ManagedThreadId != _uiThreadId)
        {
            UnsafeNativeMutations++;
            throw new InvalidOperationException("A native menu mutation occurred outside the host UI thread.");
        }
        if (!Shell.Initialized || Shell.Modal)
        {
            UnsafeNativeMutations++;
            throw new InvalidOperationException("A native menu mutation occurred while shell access was unsafe.");
        }
    }

    public void Dispose()
    {
        Disposal.Cancel();
        Disposal.Dispose();
        Scheduler.Dispose();
    }
}

internal sealed class PackageShell : IVsShell, IVsUIShell, IVsStatusbar
{
    private readonly PackageTestHost _host;
    internal bool Initialized = true;
    internal bool Modal;
    internal int PropertyReads;
    internal PackageShell(PackageTestHost host) { _host = host; }

    public int GetProperty(int property, out object value)
    {
        PropertyReads++;
        value = property == (int)__VSSPROPID4.VSSPROPID_ShellInitialized ? Initialized : Modal;
        return 0;
    }

    public int UpdateCommandUI(int immediate)
    {
        _host.UpdateRequests.Add(immediate);
        Action update = () =>
        {
            foreach (var command in _host.PublishedCommands)
            {
                command.QueryStatus();
                _host.StatusQueries++;
            }
        };
        if (immediate != 0) update();
        else _host.Scheduler.Post(update, VsTaskRunContext.UIThreadIdlePriority);
        return 0;
    }

    public int SetText(string text) { _host.LastStatus = text; return 0; }
}

internal sealed class PackageProffer : IVsProfferCommands3
{
    private readonly PackageTestHost _host;
    internal PackageProffer(PackageTestHost host) { _host = host; }

    public int FindCommandBar(object owner, ref Guid group, uint id, out object bar)
    {
        bar = group == AxialSqlTools.RefreshTemplatesCommand.CommandSet && id == 0x1100 ? _host.Root : null;
        return 0;
    }

    public int AddNamedCommand(ref Guid package, ref Guid group, string name, out uint id, string canonicalName,
        string text, string tooltip, string binding, uint resource, uint bitmap, uint flags, uint contextCount, Guid[] contexts)
    {
        _host.RecordMutation();
        id = (uint)_host.App.Commands.AllocateId();
        _host.App.Commands.Add(name, group, (int)id);
        return 0;
    }

    public int AddCommandBar(string name, uint type, object parent, uint index, out object bar)
    {
        _host.RecordMutation();
        var child = new CommandBar();
        ((CommandBars)_host.App.CommandBars).Add(name, child);
        if (parent is CommandBar parentBar)
            parentBar.Controls.Insert((int)index + 1, new CommandBarPopup { CommandBar = child, Caption = name });
        bar = child;
        return 0;
    }

    public int AddCommandBarControl(string name, object parent, uint index, uint type, out object control)
    {
        control = null;
        throw new InvalidOperationException("Unexpected legacy control registration.");
    }
}

internal sealed class NormalOnlyScheduler : IDisposable
{
    private readonly ConcurrentQueue<Action> _normal = new ConcurrentQueue<Action>();
    private readonly ConcurrentQueue<Action> _idle = new ConcurrentQueue<Action>();
    private readonly AutoResetEvent _posted = new AutoResetEvent(false);
    internal int NormalScheduled;
    internal int IdleScheduled;
    internal int NormalWorkers;
    internal int IdleWorkers;
    internal int NormalPumped;
    internal int IdlePumped => 0;

    internal void Post(Action action, VsTaskRunContext priority)
    {
        if (priority == VsTaskRunContext.UIThreadNormalPriority)
        {
            Interlocked.Increment(ref NormalScheduled);
            _normal.Enqueue(action);
        }
        else
        {
            Interlocked.Increment(ref IdleScheduled);
            _idle.Enqueue(action);
        }
        _posted.Set();
    }

    internal IDisposable EnterNormalContext()
    {
        var previous = SynchronizationContext.Current;
        SynchronizationContext.SetSynchronizationContext(new PriorityContext(this, VsTaskRunContext.UIThreadNormalPriority));
        return new RestoreContext(previous);
    }

    internal void PumpNormalUntil(Func<bool> completed, int timeoutMilliseconds = 5000)
    {
        var elapsed = Stopwatch.StartNew();
        while (!completed())
        {
            if (elapsed.ElapsedMilliseconds > timeoutMilliseconds)
                throw new InvalidOperationException("Normal-priority work did not complete; idle queue was intentionally never pumped.");
            if (_normal.TryDequeue(out var action))
            {
                using (EnterNormalContext()) action();
                NormalPumped++;
            }
            else _posted.WaitOne(5);
        }
    }

    public void Dispose() => _posted.Dispose();

    private sealed class PriorityContext : SynchronizationContext
    {
        private readonly NormalOnlyScheduler _scheduler;
        private readonly VsTaskRunContext _priority;
        internal PriorityContext(NormalOnlyScheduler scheduler, VsTaskRunContext priority)
        { _scheduler = scheduler; _priority = priority; }
        public override void Post(SendOrPostCallback callback, object state) => _scheduler.Post(() => callback(state), _priority);
        public override SynchronizationContext CreateCopy() => new PriorityContext(_scheduler, _priority);
    }

    private sealed class RestoreContext : IDisposable
    {
        private readonly SynchronizationContext _previous;
        internal RestoreContext(SynchronizationContext previous) { _previous = previous; }
        public void Dispose() => SynchronizationContext.SetSynchronizationContext(_previous);
    }
}

namespace Microsoft.VisualStudio.Threading
{
    public sealed class JoinableTask
    {
        public Task Task { get; }
        internal JoinableTask(Task task) { Task = task; }
    }
}

namespace Microsoft.VisualStudio.Shell
{
    public enum VsTaskRunContext { UIThreadIdlePriority = 3, UIThreadNormalPriority = 6 }

    public sealed partial class TestJoinableTaskFactory
    {
        internal NormalOnlyScheduler Scheduler { get; set; }
        public Task SwitchToMainThreadAsync() => SwitchToMainThreadAsync(CancellationToken.None);
        public Microsoft.VisualStudio.Threading.JoinableTask StartOnIdle(Func<Task> work, VsTaskRunContext priority)
        {
            if (priority == VsTaskRunContext.UIThreadNormalPriority) Scheduler.NormalWorkers++;
            else Scheduler.IdleWorkers++;
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            Scheduler.Post(() =>
            {
                Task running;
                try { running = work(); }
                catch (Exception ex) { completion.TrySetException(ex); return; }
                running.ContinueWith(task =>
                {
                    if (task.IsCanceled) completion.TrySetCanceled();
                    else if (task.IsFaulted) completion.TrySetException(task.Exception.InnerExceptions);
                    else completion.TrySetResult(true);
                }, CancellationToken.None, TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
            }, priority);
            return new Microsoft.VisualStudio.Threading.JoinableTask(completion.Task);
        }
    }
}

namespace Microsoft.VisualStudio.Shell.Interop
{
    public enum __VSSPROPID4 { VSSPROPID_ShellInitialized = 1, VSSPROPID_IsModal = 2 }
    public sealed class SVsShell { }
    public sealed class SVsUIShell { }
    public sealed class SVsStatusbar { }
    public sealed class SVsProfferCommands { }
    public interface IVsShell { int GetProperty(int property, out object value); }
    public interface IVsUIShell { int UpdateCommandUI(int immediate); }
    public interface IVsStatusbar { int SetText(string text); }
}

namespace EnvDTE { public sealed class DTE { } }

namespace AxialSqlTools
{
    public sealed partial class AxialSqlToolsPackage
    {
        private static PackageTestHost _testHost;
        internal static PackageLogger _logger;
        public const string PackageGuidString = "82ff597d-c4bc-469f-b990-637219074984";
        public const string PackageGuidGroup = "d8ef26a8-e88c-4ad1-85fd-ddc48a207530";
        internal AxialSqlToolsPackage(PackageTestHost host) { _testHost = host; _logger = new PackageLogger(host); }
        internal CancellationToken DisposalToken => _testHost.Disposal.Token;
        internal TestJoinableTaskFactory JoinableTaskFactory => ThreadHelper.JoinableTaskFactory;
        internal Task<object> GetServiceAsync(Type type) => Task.FromResult(GetGlobalService(type));
        internal static object GetGlobalService(Type type) => _testHost.Services.TryGetValue(type, out var value) ? value : null;
        internal void StartTemplatesForTest() => StartQueryTemplatesInitialization();
        internal Task StartupCompletionForTest => _templateStartupTask;
        internal void RecordScriptExecution(string file) => _testHost.ExecutedFiles.Add(file);
    }

    internal sealed class PackageLogger
    {
        private readonly PackageTestHost _host;
        internal PackageLogger(PackageTestHost host) { _host = host; }
        public void Info(string message, params object[] values) => _host.Logs.Add(string.Format(message, values));
        public void Warn(string message, params object[] values) => _host.Logs.Add(string.Format(message, values));
        public void Warn(Exception exception, string message) => _host.Logs.Add(message + " " + exception.Message);
        public void Error(Exception exception, string message, params object[] values) => _host.Logs.Add(string.Format(message, values) + " " + exception.Message);
    }

    internal static class SettingsManager { internal static string Folder; internal static string GetTemplatesFolder() => Folder; }
    internal static class RefreshTemplatesCommand
    {
        internal static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");
        internal const int CommandId = 4132;
    }
    internal static class OpenTemplatesFolderCommand
    {
        internal static readonly Guid CommandSet = RefreshTemplatesCommand.CommandSet;
        internal const int CommandId = 4145;
    }
}

namespace AxialSqlTools.Properties
{
    internal static class Resources { internal static System.Drawing.Image script => new System.Drawing.Image(); }
}

namespace Aurora
{
    internal sealed class CommandProcessor : CommandBase
    {
        internal AxialSqlTools.AxialSqlToolsPackage package;
        internal string FullFileName;
        internal CommandProcessor(Plugin plugin, string canonicalName, string text, string tooltip)
            : base(text, canonicalName, plugin, tooltip) { }
        public override bool IsEnabled() => true;
        public override bool OnCommand() { package.RecordScriptExecution(FullFileName); return true; }
    }
}
