using AxialSqlTools.JobQuickView;
using System;
using System.Collections.Generic;
using System.Linq;

internal static class JobQuickViewExecutionHistoryTests
{
    private static int passed;
    private static readonly DateTime Start = new DateTime(2026, 11, 8, 10, 42, 24, 637);
    private const string Server = @"SQL01\INSTANCE";

    private static int Main()
    {
        try
        {
            Run("Active execution is pinned above completed summaries", ActiveFirst);
            Run("First ever execution is visible without retained history", FirstRun);
            Run("Opening and polling never auto-select an execution", NoAutomaticSelection);
            Run("Stable rows and historical selection survive polling", StableRows);
            Run("Explicit active selection follows a completed execution", CompleteSelection);
            Run("Active selection survives a delayed summary", DelayedSummary);
            Run("User history selection cancels pending completion selection", SelectOtherDuringLag);
            Run("Unknown start never guesses a completed execution", UnknownStart);
            Run("Missing timing becoming available keeps the live row", TimingAppears);
            Run("A new start produces a distinct live row", NewStart);
            Run("New Agent session produces a distinct live row", NewSession);
            Run("History failure preserves retained rows and live status", HistoryFailure);
            Run("Idle removes live row even when history cannot load", IdleAfterHistoryFailure);
            Run("Unknown state freezes the last known active run", UnknownState);
            Run("Confirmed completion retires a stale active row when status is unknown", UnknownStateWithConfirmedCompletion);
            Run("Transport failure freezes animation and elapsed time", StaleSample);
            Run("Fresh activity resumes stale row and elapsed time", RecoverSample);
            Run("Elapsed time uses supplied server duration and monotonic delta", ElapsedTime);
            Run("Unknown elapsed time remains unavailable", UnknownDuration);
            Run("Negative duration and elapsed increments do not become negative", NegativeDuration);
            Run("Elapsed time cannot overflow", DurationOverflow);
            Run("Purged historical selection clears without choosing another run", PurgedSelection);
            Run("Ambiguous same-second summaries do not steal selection", AmbiguousCompletion);
            Run("A different execution server cannot claim the active selection", DifferentServer);
            Run("Old summaries cannot claim the active selection", HistoryBoundary);
            Run("Completed detail rows keep their step labels and durations", CompletedDetails);
            Run("Connection aliases use the actual execution server", ExecutionServer);
            Run("Activity history identity can confirm an aliased server", ActivityIdentity);
            Run("Activity identity alone does not guess a missing start time", ActivityIdentityNeedsStart);
            Run("Completing displays one live row until final status", Completing);
            Run("New run can coexist with a same-second previous summary", SameSecondNewRun);
            Run("Completing never hides an unrelated retained summary", CompletingOtherSummary);
            Run("A rapid restart cannot reuse the preceding active row", RapidRestart);
            Run("A different server does not make same-server completion ambiguous", MultiServerHistory);
            Console.WriteLine("Passed " + passed + " Quick Manage execution history tests.");
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
        passed++;
        Console.WriteLine("PASS: " + name);
    }

    private static JobQuickViewSnapshot Sample(bool? running, params JobQuickViewHistoryRun[] history)
        => new JobQuickViewSnapshot
        {
            IsRunning = running,
            ExecutionStatus = running == true ? "Running" : running == false ? "Idle" : "Unknown",
            RunningSince = running == true ? Start : (DateTime?)null,
            RunningDuration = running == true ? TimeSpan.FromSeconds(12) : (TimeSpan?)null,
            AgentSessionId = 7,
            History = history
        };

    private static JobQuickViewHistoryRun Completed(int id = 100, DateTime? start = null, string server = Server)
        => new JobQuickViewHistoryRun
        {
            InstanceId = id, Server = server, StartedAt = start ?? Start.AddDays(-1),
            RunStatus = 1, Outcome = "Succeeded", Duration = TimeSpan.FromSeconds(10)
        };

    private static JobQuickViewHistoryRun Finished(int id = 101, string server = Server)
        => Completed(id, new DateTime(Start.Ticks - Start.Ticks % TimeSpan.TicksPerSecond), server);

    private static JobQuickViewExecutionHistory StartHistory()
    {
        var history = new JobQuickViewExecutionHistory();
        history.Apply(Sample(true, Completed()), null, Server);
        return history;
    }

    private static void ActiveFirst()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true, Completed());
        var selection = history.Apply(sample, null, Server);
        Equal(2, history.Rows.Count);
        Same(history.ActiveRun, history.Rows[0]);
        True(history.Rows[0].IsAnimating);
        Equal(4, history.Rows[0].Item.RunStatus);
        Equal(0, history.Rows[0].Item.InstanceId);
        Equal("Succeeded", history.Rows[1].Item.Outcome);
        Equal(1, sample.History.Count);
        True(selection == null);
    }

    private static void FirstRun()
    {
        var history = new JobQuickViewExecutionHistory();
        history.Apply(Sample(true), null, Server);
        Equal(1, history.Rows.Count);
        True(history.ActiveRun.Title.StartsWith("Running now"));
        True(history.ActiveRun.Summary.Contains("Running"));
        True(history.ActiveRun.Summary.Contains("0:00:12 elapsed"));
    }

    private static void NoAutomaticSelection()
    {
        var history = StartHistory();
        True(history.Apply(Sample(true, Completed()), null, Server) == null);
        True(history.Apply(Sample(false, Finished(), Completed()), null, Server) == null);
    }

    private static void StableRows()
    {
        var history = StartHistory();
        var selected = history.Rows[1];
        var active = history.ActiveRun;
        int changes = 0;
        int propertyChanges = 0;
        selected.PropertyChanged += (sender, args) => propertyChanges++;
        history.Rows.CollectionChanged += (sender, args) => changes++;
        Same(selected, history.Apply(Sample(true, Completed()), selected, Server));
        Same(selected, history.Rows[1]);
        Same(active, history.ActiveRun);
        Equal(0, changes);
        Equal(0, propertyChanges);
        Same(selected, history.Apply(Sample(false, Finished(), Completed()), selected, Server));
        Same(selected, history.Rows[1]);
    }

    private static void CompleteSelection()
    {
        var history = StartHistory();
        var selected = history.ActiveRun;
        var next = history.Apply(Sample(false, Finished(server: @"sql01\instance"), Completed()), selected, Server + ",1132");
        Equal(101, next.Item.InstanceId);
        True(!next.IsActive);
        True(history.ActiveRun == null);
        Equal(2, history.Rows.Count);
    }

    private static void DelayedSummary()
    {
        var history = StartHistory();
        var selected = history.Apply(Sample(false, Completed()), history.ActiveRun, Server);
        True(selected == null);
        True(history.ActiveRun == null);
        selected = history.Apply(Sample(false, Finished(), Completed()), selected, Server);
        Equal(101, selected.Item.InstanceId);
        True(!selected.IsActive);
    }

    private static void SelectOtherDuringLag()
    {
        var history = StartHistory();
        True(history.Apply(Sample(false, Completed()), history.ActiveRun, Server) == null);
        var historical = history.Rows[0];
        Same(historical, history.Apply(Sample(false, Finished(), Completed()), historical, Server));
    }

    private static void UnknownStart()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true, Completed());
        sample.RunningSince = null;
        history.Apply(sample, null, Server);
        True(history.Apply(Sample(false, Finished(), Completed()), history.ActiveRun, Server) == null);
        True(history.Apply(Sample(false, Finished(), Completed()), null, Server) == null);
    }

    private static void TimingAppears()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true);
        sample.RunningSince = null;
        history.Apply(sample, null, Server);
        var selected = history.ActiveRun;
        Same(selected, history.Apply(Sample(true), selected, Server));
        Equal(Start, selected.Item.StartedAt.Value);
    }

    private static void NewStart()
    {
        var history = StartHistory();
        var previous = history.ActiveRun;
        var next = Sample(true, Finished(), Completed());
        next.RunningSince = Start.AddMinutes(1);
        var selected = history.Apply(next, previous, Server);
        True(!ReferenceEquals(previous, history.ActiveRun));
        Equal(101, selected.Item.InstanceId);
        Same(history.ActiveRun, history.Rows[0]);
    }

    private static void NewSession()
    {
        var history = StartHistory();
        var previous = history.ActiveRun;
        var next = Sample(true, Completed());
        next.AgentSessionId = 8;
        history.Apply(next, previous, Server);
        True(!ReferenceEquals(previous, history.ActiveRun));
    }

    private static void HistoryFailure()
    {
        var history = StartHistory();
        var selected = history.Rows[1];
        var failed = Sample(true);
        failed.HistoryLoaded = false;
        failed.RunningDuration = TimeSpan.FromSeconds(20);
        Same(selected, history.Apply(failed, selected, Server));
        Equal(2, history.Rows.Count);
        Equal("0:00:20", history.ActiveRun.DurationText);
    }

    private static void IdleAfterHistoryFailure()
    {
        var history = StartHistory();
        var failed = Sample(false);
        failed.HistoryLoaded = false;
        True(history.Apply(failed, history.ActiveRun, Server) == null);
        Equal(1, history.Rows.Count);
        True(history.ActiveRun == null);
        Equal(100, history.Rows[0].Item.InstanceId);
    }

    private static void UnknownState()
    {
        var history = StartHistory();
        var selected = history.ActiveRun;
        Same(selected, history.Apply(Sample(null, Completed()), selected, Server));
        True(selected.IsStale);
        True(!selected.IsAnimating);
        True(selected.Summary.StartsWith("Last known: Running"));
        selected.UpdateElapsed(TimeSpan.FromSeconds(20));
        Equal("0:00:12", selected.DurationText);
        var empty = new JobQuickViewExecutionHistory();
        empty.Apply(Sample(null), null, Server);
        Equal(0, empty.Rows.Count);
    }

    private static void UnknownStateWithConfirmedCompletion()
    {
        var history = StartHistory();
        var previous = history.ActiveRun;
        var sample = Sample(null, Finished(), Completed());
        // A timestamp alone is not enough to discard last known activity.
        Same(previous, history.Apply(sample, previous, Server));
        True(previous.IsStale);
        sample.ActivityStartedAt = Start;
        sample.ActivityStoppedAt = Start.AddSeconds(12);
        sample.LastCompletedHistoryId = 101;
        var selected = history.Apply(sample, previous, Server);
        True(history.ActiveRun == null);
        Equal(2, history.Rows.Count);
        Equal(101, selected.Item.InstanceId);
        True(!selected.IsActive);
    }

    private static void StaleSample()
    {
        var history = StartHistory();
        var active = history.ActiveRun;
        active.UpdateElapsed(TimeSpan.FromSeconds(2));
        history.MarkActivityStale();
        active.UpdateElapsed(TimeSpan.FromHours(1));
        Equal("0:00:14", active.DurationText);
        True(active.IsStale);
        True(!active.IsAnimating);
        True(!active.Title.StartsWith("Running now"));
    }

    private static void RecoverSample()
    {
        var history = StartHistory();
        var selected = history.ActiveRun;
        history.MarkActivityStale();
        var fresh = Sample(true, Completed());
        fresh.RunningDuration = TimeSpan.FromSeconds(30);
        Same(selected, history.Apply(fresh, selected, Server));
        selected.UpdateElapsed(TimeSpan.FromSeconds(2));
        Equal("0:00:32", selected.DurationText);
        True(selected.IsAnimating);
        True(!selected.IsStale);
    }

    private static void ElapsedTime()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true);
        sample.RunningSince = new DateTime(2001, 1, 1); // Deliberately unrelated to this workstation's clock.
        sample.RunningDuration = TimeSpan.FromHours(27) + TimeSpan.FromSeconds(58);
        history.Apply(sample, null, Server);
        var active = history.ActiveRun;
        var changed = new List<string>();
        active.PropertyChanged += (sender, args) => changed.Add(args.PropertyName);
        active.UpdateElapsed(TimeSpan.FromSeconds(4));
        Equal("27:01:02", active.DurationText);
        True(changed.Contains("DurationText") && changed.Contains("Summary"));
        Equal(TimeSpan.FromHours(27) + TimeSpan.FromSeconds(58), sample.RunningDuration.Value);
    }

    private static void UnknownDuration()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true);
        sample.RunningDuration = null;
        history.Apply(sample, null, Server);
        history.ActiveRun.UpdateElapsed(TimeSpan.FromMinutes(1));
        Equal("Elapsed time unavailable", history.ActiveRun.DurationText);
    }

    private static void NegativeDuration()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true);
        sample.RunningDuration = TimeSpan.FromSeconds(-10);
        history.Apply(sample, null, Server);
        history.ActiveRun.UpdateElapsed(TimeSpan.FromSeconds(-1));
        Equal("0:00:00", history.ActiveRun.DurationText);
    }

    private static void DurationOverflow()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true);
        sample.RunningDuration = TimeSpan.MaxValue - TimeSpan.FromSeconds(1);
        history.Apply(sample, null, Server);
        history.ActiveRun.UpdateElapsed(TimeSpan.FromSeconds(3));
        True(history.ActiveRun.DurationText.StartsWith(((long)TimeSpan.MaxValue.TotalHours).ToString()));
    }

    private static void PurgedSelection()
    {
        var history = StartHistory();
        var selected = history.Rows[1];
        True(history.Apply(Sample(true), selected, Server) == null);
        Equal(1, history.Rows.Count);
    }

    private static void AmbiguousCompletion()
    {
        var history = StartHistory();
        True(history.Apply(Sample(false, Finished(102), Finished(), Completed()), history.ActiveRun, Server) == null);
    }

    private static void DifferentServer()
    {
        var history = StartHistory();
        True(history.Apply(Sample(false, Finished(server: @"SQL02\INSTANCE"), Completed()), history.ActiveRun, Server) == null);
    }

    private static void HistoryBoundary()
    {
        var history = new JobQuickViewExecutionHistory();
        history.Apply(Sample(true, Finished(100)), null, Server);
        True(history.Apply(Sample(false, Finished(100)), history.ActiveRun, Server) == null);
    }

    private static void CompletedDetails()
    {
        var item = Completed();
        item.StepId = 3;
        item.StepName = "Collect data";
        var row = new JobQuickViewExecutionRow(item);
        Equal("Step 3: Collect data", row.EntryLabel);
        Equal("0:00:10", row.DurationText);
        True(!row.IsActive && !row.IsAnimating && !row.IsStale);
        row.UpdateElapsed(TimeSpan.FromHours(1));
        Equal("0:00:10", row.DurationText);
    }


    private static void ExecutionServer()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true, Completed());
        sample.ExecutionServer = Server;
        history.Apply(sample, null, "tcp:db-alias.example.com,1132");
        var finished = Sample(false, Finished(), Completed());
        finished.ExecutionServer = Server;
        Equal(101, history.Apply(finished, history.ActiveRun, "tcp:db-alias.example.com,1132").Item.InstanceId);
    }

    private static void ActivityIdentity()
    {
        var history = new JobQuickViewExecutionHistory();
        history.Apply(Sample(true, Completed()), null, "db-alias.example.com");
        var finished = Sample(false, Finished(), Completed());
        finished.LastCompletedHistoryId = 101;
        Equal(101, history.Apply(finished, history.ActiveRun, "db-alias.example.com").Item.InstanceId);
    }

    private static void ActivityIdentityNeedsStart()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true, Completed());
        sample.RunningSince = null;
        history.Apply(sample, null, "db-alias.example.com");
        var finished = Sample(false, Finished(), Completed());
        finished.LastCompletedHistoryId = 101;
        True(history.Apply(finished, history.ActiveRun, "db-alias.example.com") == null);
    }

    private static void Completing()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true, Finished(), Completed());
        sample.ExecutionStatus = "Completing";
        sample.LastCompletedHistoryId = 101;
        history.Apply(sample, null, Server);
        Equal(2, history.Rows.Count);
        True(history.ActiveRun.Summary.StartsWith("Completing"));
        True(history.Rows.All(row => row.Item.InstanceId != 101));
        Equal(2, sample.History.Count);
        var selected = history.Apply(Sample(false, Finished(), Completed()), history.ActiveRun, Server);
        Equal(101, selected.Item.InstanceId);
    }

    private static void SameSecondNewRun()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true, Finished(), Completed());
        sample.LastCompletedHistoryId = 101;
        history.Apply(sample, null, Server);
        Equal(3, history.Rows.Count);
        True(history.ActiveRun.IsAnimating);
    }

    private static void CompletingOtherSummary()
    {
        var history = new JobQuickViewExecutionHistory();
        var sample = Sample(true, Completed());
        sample.ExecutionStatus = "Completing";
        sample.LastCompletedHistoryId = 100;
        history.Apply(sample, null, Server);
        Equal(2, history.Rows.Count);
    }

    private static void RapidRestart()
    {
        var history = StartHistory();
        var previous = history.ActiveRun;
        var sample = Sample(true, Finished(), Completed());
        sample.ActivityStoppedAt = Start.AddSeconds(12);
        sample.ActivityStartedAt = Start;
        sample.RunningSince = null;
        sample.RunningDuration = null;
        sample.LastCompletedHistoryId = 101;
        var selection = history.Apply(sample, previous, Server);
        True(!ReferenceEquals(previous, history.ActiveRun));
        True(!history.ActiveRun.Item.StartedAt.HasValue);
        Equal(101, selection.Item.InstanceId);
        var queued = history.ActiveRun;
        sample.ExecutionStatus = "Waiting for thread";
        Same(queued, history.Apply(sample, queued, Server));
        Same(queued, history.ActiveRun);
        Equal(3, history.Rows.Count);
    }

    private static void MultiServerHistory()
    {
        var history = StartHistory();
        var selected = history.Apply(Sample(false, Finished(102, @"SQL02\INSTANCE"), Finished(), Completed()), history.ActiveRun, Server);
        Equal(101, selected.Item.InstanceId);
    }

    private static void True(bool condition)
    {
        if (!condition) throw new InvalidOperationException("Assertion failed.");
    }
    private static void Same(object expected, object actual)
    {
        if (!ReferenceEquals(expected, actual)) throw new InvalidOperationException("Expected the same row instance.");
    }
    private static void Equal<T>(T expected, T actual)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new InvalidOperationException("Expected " + expected + "; actual " + actual + ".");
    }
}
