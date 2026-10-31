using AxialSqlTools;
using Microsoft.VisualStudio.CommandBars;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

internal static partial class CommandRegistrationTests
{
    private static void RunAsyncMenuTests()
    {
        Run("Populate awaits the UI gate and checkpoints each script with linear native queries", PopulateAwaitsPerScriptCheckpoint);
        Run("Cancellation stops a large population before the next native registration", CancellationStopsPopulate);
        Run("Cancellation interrupts control cleanup without swallowing the request", CancellationStopsControlCleanup);
        Run("Cyclic native popups are bounded and never delete the root actions", CyclicPopupCleanupIsBounded);
        Run("Tracked-command cleanup awaits checkpoints and preserves unprocessed callbacks on cancellation", CleanupAwaitsAndCancels);
    }

    private static void RemoveUnused(Fixture fixture, ISet<string> activeNames)
    {
        fixture.Registry.RemoveUnusedCommandsAsync(activeNames, () => Task.CompletedTask,
            CancellationToken.None).GetAwaiter().GetResult();
    }

    private static QueryTemplateFolder FlatTree(int scriptCount)
    {
        string root = Path.Combine(Path.GetTempPath(), "AxialSqlTools-async-rendering-test");
        return new QueryTemplateFolder(root, new List<QueryTemplateFolder>(),
            Enumerable.Range(0, scriptCount)
                .Select(i => new QueryTemplateFile(Path.Combine(root, "script" + i + ".sql"))).ToList());
    }

    private static void PopulateAwaitsPerScriptCheckpoint()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        const int count = 512;
        var gate = new TaskCompletionSource<bool>();
        var warnings = new List<string>();
        int checkpoints = 0;
        fixture.App.Commands.CommandInfoCallCount = 0;
        var pending = QueryTemplateMenuControls.PopulateAsync(FlatTree(count), fixture.Bar,
            fixture.Plugin, fixture.Registry, file => fixture.Handler(file.CommandName), warnings,
            () => ++checkpoints == 1 ? gate.Task : Task.CompletedTask, CancellationToken.None);
        Assert(!pending.IsCompleted, "Populate did not await the incomplete UI checkpoint.");
        Equal(2, fixture.Bar.Controls.Count);
        Equal(0, fixture.MenuService.Count);
        Equal(0, fixture.App.Commands.CommandInfoCallCount);
        gate.SetResult(true);
        var result = pending.GetAwaiter().GetResult();
        Equal(count, checkpoints);
        Equal(count, result.Loaded);
        Equal(0, result.Failures);
        Equal(count, fixture.App.Commands.CommandInfoCallCount);
        Equal(count + 2, fixture.Bar.Controls.Count);
        Equal(0, warnings.Count);
        Equal(0, fixture.App.Commands.EnumerationCalls);
    }

    private static void CancellationStopsPopulate()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        var warnings = new List<string>();
        int checkpoints = 0;
        using (var cancellation = new CancellationTokenSource())
        {
            Throws<OperationCanceledException>(() => QueryTemplateMenuControls.PopulateAsync(FlatTree(512),
                fixture.Bar, fixture.Plugin, fixture.Registry, file => fixture.Handler(file.CommandName), warnings,
                () =>
                {
                    if (++checkpoints == 27) cancellation.Cancel();
                    return Task.CompletedTask;
                }, cancellation.Token).GetAwaiter().GetResult());
        }
        Equal(27, checkpoints);
        Equal(26, fixture.MenuService.Count);
        Equal(28, fixture.Bar.Controls.Count);
        Equal(0, warnings.Count);
    }

    private static void CancellationStopsControlCleanup()
    {
        var fixture = new Fixture();
        AddBuiltIns(fixture);
        for (int i = 0; i < 10; i++)
            fixture.Registry.RegisterCommand(false, fixture.Handler("cleanup" + i), true, appendOnly: true);
        var refresh = fixture.Bar.Controls[1];
        var open = fixture.Bar.Controls[2];
        var dynamicControls = fixture.Bar.Controls.Skip(2).ToList();
        var warnings = new List<Exception>();
        using (var cancellation = new CancellationTokenSource())
        {
            Throws<OperationCanceledException>(() => QueryTemplateMenuControls.ClearAsync(fixture.Bar,
                fixture.App.Commands, RefreshId, OpenFolderId, warnings.Add,
                () =>
                {
                    if (dynamicControls.Any(control => control.Deleted)) cancellation.Cancel();
                    return Task.CompletedTask;
                }, cancellation.Token).GetAwaiter().GetResult());
        }
        Equal(1, dynamicControls.Count(control => control.Deleted));
        Equal(11, fixture.Bar.Controls.Count);
        Equal(0, warnings.Count);
        Assert(!refresh.Deleted && !open.Deleted, "Cancellation damaged a built-in action.");
        ClearMenu(fixture);
        Equal(2, fixture.Bar.Controls.Count);
    }

    private static void CyclicPopupCleanupIsBounded()
    {
        var fixture = new Fixture();
        fixture.Bar.Name = "Query Templates";
        AddBuiltIns(fixture);
        var refresh = fixture.Bar.Controls[1];
        var open = fixture.Bar.Controls[2];
        var child = fixture.Plugin.AddCommandBarMenu("Child", MsoBarPosition.msoBarMenuBar, fixture.Bar);
        var rootPopup = fixture.Bar.Controls[3];
        var loop = new CommandBarPopup { Caption = "Back to root", CommandBar = fixture.Bar };
        child.Controls.Add(loop);
        int rootEnumerations = fixture.Bar.Controls.EnumerationCalls;
        int childEnumerations = child.Controls.EnumerationCalls;
        fixture.Bar.Controls.ThrowOnEnumeration = true;
        child.Controls.ThrowOnEnumeration = true;
        int checkpoints = 0;
        var warnings = new List<Exception>();
        QueryTemplateMenuControls.ClearAsync(fixture.Bar, fixture.App.Commands, RefreshId, OpenFolderId,
            warnings.Add, () =>
            {
                if (++checkpoints > 64) throw new InvalidOperationException("Cyclic native cleanup exceeded its finite checkpoint budget.");
                return Task.CompletedTask;
            }, CancellationToken.None).GetAwaiter().GetResult();
        Equal(1, warnings.Count);
        Assert(warnings[0].Message.Contains("cyclic"), "Cycle was not identified in the cleanup warning.");
        Assert(loop.Deleted && rootPopup.Deleted && child.Deleted, "The cyclic popup branch was not removed.");
        Assert(!fixture.Bar.Deleted && !refresh.Deleted && !open.Deleted, "Cycle cleanup deleted the ancestor root or its actions.");
        Equal(2, fixture.Bar.Controls.Count);
        Equal(rootEnumerations, fixture.Bar.Controls.EnumerationCalls);
        Equal(childEnumerations, child.Controls.EnumerationCalls);
    }

    private static void CleanupAwaitsAndCancels()
    {
        var fixture = new Fixture();
        var handlers = new List<Handler>();
        var commands = new List<Microsoft.VisualStudio.Shell.OleMenuCommand>();
        for (int i = 0; i < 4; i++)
        {
            string name = "AxialSqlTools_Template_cancel_cleanup_" + i;
            var handler = fixture.Handler(name);
            handlers.Add(handler);
            fixture.Registry.RegisterCommand(false, handler, true, appendOnly: true);
            commands.Add(fixture.Command(name));
        }
        var gate = new TaskCompletionSource<bool>();
        int checkpoints = 0;
        fixture.App.Commands.ItemCalls = 0;
        using (var cancellation = new CancellationTokenSource())
        {
            var pending = fixture.Registry.RemoveUnusedCommandsAsync(new HashSet<string>(),
                () =>
                {
                    checkpoints++;
                    if (checkpoints == 1) return gate.Task;
                    if (checkpoints == 3) cancellation.Cancel();
                    return Task.CompletedTask;
                }, cancellation.Token);
            Assert(!pending.IsCompleted, "Tracked cleanup did not await its first UI gate.");
            Equal(4, fixture.MenuService.Count);
            Equal(0, fixture.App.Commands.ItemCalls);
            Equal(0, fixture.App.Commands.DeletedCount);
            gate.SetResult(true);
            Throws<OperationCanceledException>(() => pending.GetAwaiter().GetResult());
        }
        Equal(3, checkpoints);
        Equal(2, fixture.App.Commands.ItemCalls);
        Equal(2, fixture.App.Commands.DeletedCount);
        Equal(2, fixture.App.Commands.Count);
        Equal(2, fixture.MenuService.Count);
        foreach (var command in commands)
        {
            command.QueryStatus();
            command.Invoke();
        }
        Equal(2, commands.Count(command => command.Enabled));
        Equal(2, handlers.Sum(handler => handler.Invocations));
        Equal(0, fixture.App.Commands.EnumerationCalls);
        RemoveUnused(fixture, new HashSet<string>());
        Equal(0, fixture.MenuService.Count);
        Equal(0, fixture.App.Commands.Count);
    }
}
