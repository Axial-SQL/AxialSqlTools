using AxialSqlTools;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

internal static class PackageOrchestrationTests
{
    private static int _passed;

    private static int Main()
    {
        try
        {
            Run("The fake host pumps normal posts and never pumps idle-priority tasks", SchedulerModelsIdleStarvation);
            Run("Actual package startup eagerly loads root and nested scripts without a menu-open event", StartupWithoutOpeningMenus);
            Run("Actual startup makes no native mutations while a modal dialog is open", ModalStartupWaitsSafely);
            Run("Disposal before the queued startup worker runs completes without native changes", DisposeBeforeWorkerStarts);
            Run("Disposal interrupts a modal startup wait without native changes", DisposeDuringModalWait);
            Console.WriteLine("Passed " + _passed + " actual-package orchestration tests (normal-only fake host).");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static void Run(string name, Action test)
    {
        test();
        _passed++;
        Console.WriteLine("PASS: " + name);
    }

    private static void SchedulerModelsIdleStarvation()
    {
        using (var scheduler = new NormalOnlyScheduler())
        using (scheduler.EnterNormalContext())
        {
            ThreadHelper.JoinableTaskFactory.Scheduler = scheduler;
            bool idleRan = false;
            var idle = ThreadHelper.JoinableTaskFactory.StartOnIdle(() =>
            {
                idleRan = true;
                return Task.CompletedTask;
            }, VsTaskRunContext.UIThreadIdlePriority);
            var normal = ThreadHelper.JoinableTaskFactory.StartOnIdle(async () =>
            {
                await Task.Yield(); // Continuations must retain normal priority.
                await Task.Delay(1);
            }, VsTaskRunContext.UIThreadNormalPriority);
            scheduler.PumpNormalUntil(() => normal.Task.IsCompleted);
            normal.Task.GetAwaiter().GetResult();
            Assert(!idleRan && !idle.Task.IsCompleted, "An idle task was accidentally pumped.");
            Equal(1, scheduler.IdleScheduled);
            Equal(0, scheduler.IdlePumped);
            Assert(scheduler.NormalPumped >= 3, "Normal-priority continuations were not exercised.");
        }
    }

    private static void StartupWithoutOpeningMenus()
    {
        WithCatalog(folder =>
        {
            using (var host = new PackageTestHost(folder))
            using (host.Scheduler.EnterNormalContext())
            {
                host.Package.StartTemplatesForTest();
                // The previous production code fails here before any queue is
                // pumped: it puts startup behind an idle-priority worker.
                Assert(host.Scheduler.IdleWorkers == 0, "Startup scheduled an idle-priority worker instead of normal dispatcher work.");
                host.Scheduler.PumpNormalUntil(() => host.Package.StartupCompletionForTest.IsCompleted);
                host.Package.StartupCompletionForTest.GetAwaiter().GetResult();
                VerifyCompletedCatalog(host, folder);
            }
        });
    }

    private static void ModalStartupWaitsSafely()
    {
        WithCatalog(folder =>
        {
            using (var host = new PackageTestHost(folder))
            using (host.Scheduler.EnterNormalContext())
            {
                host.Shell.Modal = true;
                host.Package.StartTemplatesForTest();
                host.Scheduler.PumpNormalUntil(() => host.Shell.PropertyReads >= 2);
                Equal(0, host.NativeMutations);
                Equal(0, host.UnsafeNativeMutations);
                Equal(2, host.Root.Controls.Count);
                Assert(!host.Package.StartupCompletionForTest.IsCompleted, "A modal startup wait completed prematurely.");
                host.Shell.Modal = false;
                host.Scheduler.PumpNormalUntil(() => host.Package.StartupCompletionForTest.IsCompleted);
                host.Package.StartupCompletionForTest.GetAwaiter().GetResult();
                VerifyCompletedCatalog(host, folder);
            }
        });
    }

    private static void DisposeBeforeWorkerStarts()
    {
        WithCatalog(folder =>
        {
            using (var host = new PackageTestHost(folder))
            using (host.Scheduler.EnterNormalContext())
            {
                host.Package.StartTemplatesForTest();
                host.Disposal.Cancel();
                host.Scheduler.PumpNormalUntil(() => host.Package.StartupCompletionForTest.IsCompleted);
                host.Package.StartupCompletionForTest.GetAwaiter().GetResult();
                VerifyCanceledStartup(host);
            }
        });
    }

    private static void DisposeDuringModalWait()
    {
        WithCatalog(folder =>
        {
            using (var host = new PackageTestHost(folder))
            using (host.Scheduler.EnterNormalContext())
            {
                host.Shell.Modal = true;
                host.Package.StartTemplatesForTest();
                host.Scheduler.PumpNormalUntil(() => host.Shell.PropertyReads >= 2);
                host.Disposal.Cancel();
                host.Scheduler.PumpNormalUntil(() => host.Package.StartupCompletionForTest.IsCompleted);
                host.Package.StartupCompletionForTest.GetAwaiter().GetResult();
                VerifyCanceledStartup(host);
            }
        });
    }

    private static void VerifyCanceledStartup(PackageTestHost host)
    {
        Equal(0, host.NativeMutations);
        Equal(0, host.UnsafeNativeMutations);
        Equal(2, host.Root.Controls.Count);
        Equal(0, host.UpdateRequests.Count);
        Equal(0, host.Scheduler.IdleScheduled);
        Equal(0, host.Scheduler.IdlePumped);
    }

    private static void VerifyCompletedCatalog(PackageTestHost host, string folder)
    {
        string[] files = Directory.GetFiles(folder, "*", SearchOption.AllDirectories);
        var buttons = ScriptButtons(host.Root).ToList();
        Equal(files.Length, buttons.Count);
        Equal(files.Length, host.PublishedCommands.Count);
        Assert(host.UpdateRequests.SequenceEqual(new[] { 0 }), "Final command status refresh was not deferred and exactly once.");
        Equal(0, host.StatusQueries); // Deferred global invalidation is deliberately never pumped.
        Assert(host.Package.TemplateRefreshMessage.StartsWith("Loaded " + files.Length + " of " + files.Length + " query templates.", StringComparison.Ordinal),
            "Unexpected startup result: " + host.Package.TemplateRefreshMessage + "\n" + string.Join("\n", host.Logs));
        Equal(host.Package.TemplateRefreshMessage, host.LastStatus);
        foreach (var button in buttons)
        {
            Assert(button.Visible && button.Enabled && button.Picture != null, "A script was not ready before any menu was opened.");
            button.Click();
        }
        VerifyPopupState(host.Root);
        Assert(files.OrderBy(path => path, StringComparer.Ordinal).SequenceEqual(
            host.ExecutedFiles.OrderBy(path => path, StringComparer.Ordinal)), "Native root/nested clicks did not map to the discovered files.");
        Equal(0, host.MenuOpenEvents);
        Equal(0, host.UnsafeNativeMutations);
        Equal(0, host.Scheduler.IdleWorkers);
        Equal(1, host.Scheduler.NormalWorkers);
        Equal(1, host.Scheduler.IdleScheduled); // Only the optional global UI invalidation remains pending.
        Equal(0, host.Scheduler.IdlePumped);
        Equal(0, host.App.Commands.EnumerationCalls);
    }

    private static void VerifyPopupState(CommandBar bar)
    {
        foreach (CommandBarControl control in bar.Controls)
        {
            if (!(control is CommandBarPopup popup)) continue;
            Assert(popup.Visible && popup.Enabled, "A nested popup was not usable before a deferred global status update.");
            VerifyPopupState(popup.CommandBar);
        }
    }

    private static IEnumerable<CommandBarButton> ScriptButtons(CommandBar bar)
    {
        foreach (CommandBarControl control in bar.Controls)
        {
            if (control is CommandBarPopup popup)
            {
                foreach (var child in ScriptButtons(popup.CommandBar)) yield return child;
            }
            else if (control is CommandBarButton button
                && button.NativeGuid != RefreshTemplatesCommand.CommandSet.ToString("B"))
                yield return button;
        }
    }

    private static void WithCatalog(Action<string> test)
    {
        string folder = Path.Combine(Path.GetTempPath(), "AxialSqlTools-startup-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(folder, "Alpha", "Nested"));
        Directory.CreateDirectory(Path.Combine(folder, "Other"));
        try
        {
            for (int i = 0; i < 65; i++) File.WriteAllText(Path.Combine(folder, "root-" + i + ".sql"), "select " + i + ";");
            File.WriteAllText(Path.Combine(folder, "Alpha", "child.sql"), "select 100;");
            File.WriteAllText(Path.Combine(folder, "Alpha", "Nested", "deep.sql"), "select 101;");
            File.WriteAllText(Path.Combine(folder, "Other", "other.sql"), "select 102;");
            test(folder);
        }
        finally { Directory.Delete(folder, true); }
    }

    private static void Equal<T>(T expected, T actual) => Assert(object.Equals(expected, actual), "Expected '" + expected + "', got '" + actual + "'.");
    private static void Assert(bool value, string message) { if (!value) throw new InvalidOperationException(message); }
}
