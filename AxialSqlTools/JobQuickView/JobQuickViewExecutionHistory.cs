using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;

namespace AxialSqlTools.JobQuickView
{
    /// <summary>
    /// Reconciles retained summaries with the live execution without rebuilding unchanged rows.
    /// The live row is presentation state and never becomes part of retained SQL Agent history.
    /// </summary>
    internal sealed class JobQuickViewExecutionHistory
    {
        public ObservableCollection<JobQuickViewExecutionRow> Rows { get; } = new ObservableCollection<JobQuickViewExecutionRow>();
        public JobQuickViewExecutionRow ActiveRun { get; private set; }
        private JobQuickViewExecutionRow pendingCompletionSelection;

        public JobQuickViewExecutionRow Apply(JobQuickViewSnapshot snapshot,
            JobQuickViewExecutionRow selection, string serverName)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            var previousActive = ActiveRun;
            string executionServer = string.IsNullOrWhiteSpace(snapshot.ExecutionServer) ? serverName : snapshot.ExecutionServer;
            var retained = Rows.Where(row => !row.IsActive).ToDictionary(row => row.Item.InstanceId);
            var completed = new List<JobQuickViewExecutionRow>();
            if (snapshot.HistoryLoaded)
            {
                foreach (var item in snapshot.History)
                {
                    JobQuickViewExecutionRow row;
                    if (!retained.TryGetValue(item.InstanceId, out row)) row = new JobQuickViewExecutionRow(item);
                    else row.UpdateItem(item);
                    completed.Add(row);
                }
            }
            else
            {
                // Losing access to history must not hide previously loaded executions.
                completed.AddRange(Rows.Where(row => !row.IsActive));
            }

            if (snapshot.IsRunning == true)
            {
                // Agent may have written its final summary while it still reports Completing.
                // Only the activity row's exact history ID plus matching start can identify that
                // duplicate. A new run's start may share the previous run's whole-second value.
                if (string.Equals(snapshot.ExecutionStatus, "Completing", StringComparison.Ordinal) &&
                    snapshot.LastCompletedHistoryId.HasValue && snapshot.RunningSince.HasValue)
                {
                    completed.RemoveAll(row => row.Item.InstanceId == snapshot.LastCompletedHistoryId &&
                        SameStartSecond(row.Item.StartedAt, snapshot.RunningSince));
                }
                bool sameRun = previousActive != null &&
                    (!snapshot.ActivityStoppedAt.HasValue ||
                     string.Equals(snapshot.ExecutionStatus, "Completing", StringComparison.Ordinal) ||
                     (!previousActive.Item.StartedAt.HasValue && previousActive.MatchesStoppedActivity(snapshot))) &&
                    (!snapshot.AgentSessionId.HasValue || !previousActive.AgentSessionId.HasValue ||
                     snapshot.AgentSessionId.Value == previousActive.AgentSessionId.Value) &&
                    (!snapshot.RunningSince.HasValue || !previousActive.Item.StartedAt.HasValue ||
                     snapshot.RunningSince.Value == previousActive.Item.StartedAt.Value);
                if (!sameRun)
                {
                    ActiveRun = JobQuickViewExecutionRow.CreateActive(snapshot, executionServer,
                        completed.Count == 0 ? 0 : completed.Max(row => row.Item.InstanceId));
                }
                else
                {
                    ActiveRun.UpdateActive(snapshot, executionServer);
                }
            }
            else if (snapshot.IsRunning == false)
            {
                ActiveRun = null;
            }
            else
            {
                // Unknown alone is not idle. Retire a stale execution only when the same
                // Agent activity and its exact final history row independently prove completion.
                bool completedActivity = ActiveRun != null && snapshot.HistoryLoaded &&
                    ActiveRun.Item.StartedAt.HasValue && snapshot.ActivityStartedAt == ActiveRun.Item.StartedAt &&
                    snapshot.ActivityStoppedAt.HasValue && snapshot.AgentSessionId.HasValue &&
                    snapshot.AgentSessionId == ActiveRun.AgentSessionId && snapshot.LastCompletedHistoryId.HasValue &&
                    completed.Count(row => row.Item.InstanceId == snapshot.LastCompletedHistoryId &&
                        row.Item.InstanceId > ActiveRun.HistoryBoundary &&
                        SameStartSecond(row.Item.StartedAt, ActiveRun.Item.StartedAt) &&
                        SameServer(row.Item.Server, ActiveRun.Item.Server, executionServer)) == 1;
                if (completedActivity) ActiveRun = null;
                else ActiveRun?.MarkActivityStale();
            }

            var desired = new List<JobQuickViewExecutionRow>();
            if (ActiveRun != null) desired.Add(ActiveRun);
            desired.AddRange(completed);
            for (int index = 0; index < desired.Count; index++)
            {
                var row = desired[index];
                if (index < Rows.Count && ReferenceEquals(Rows[index], row)) continue;
                int existingIndex = Rows.IndexOf(row);
                if (existingIndex >= 0) Rows.Move(existingIndex, index);
                else Rows.Insert(index, row);
            }
            while (Rows.Count > desired.Count) Rows.RemoveAt(Rows.Count - 1);

            if (selection != null && Rows.Contains(selection))
            {
                pendingCompletionSelection = null;
                return selection;
            }
            if (selection != null && !selection.IsActive)
            {
                pendingCompletionSelection = null;
                return null;
            }
            var finishedSelection = selection ?? pendingCompletionSelection;
            // Never guess which execution completed when the active start time was unavailable.
            if (finishedSelection?.Item.StartedAt == null) return null;
            var matches = completed.Where(row =>
                row.Item.InstanceId > finishedSelection.HistoryBoundary &&
                SameStartSecond(row.Item.StartedAt, finishedSelection.Item.StartedAt) &&
                (SameServer(row.Item.Server, finishedSelection.Item.Server, executionServer) ||
                 (snapshot.AgentSessionId.HasValue && snapshot.AgentSessionId == finishedSelection.AgentSessionId &&
                  snapshot.LastCompletedHistoryId == row.Item.InstanceId))).Take(2).ToArray();
            if (matches.Length == 1)
            {
                pendingCompletionSelection = null;
                return matches[0];
            }
            // Agent can report idle just before its summary becomes visible. Remember an
            // explicitly selected active execution until a later summary sample can identify it.
            pendingCompletionSelection = matches.Length == 0 ? finishedSelection : null;
            return null;
        }

        public void MarkActivityStale() => ActiveRun?.MarkActivityStale();

        private static bool SameStartSecond(DateTime? left, DateTime? right)
            => left.HasValue && right.HasValue &&
               left.Value.Ticks / TimeSpan.TicksPerSecond == right.Value.Ticks / TimeSpan.TicksPerSecond;

        private static bool SameServer(string left, string right, string fallback)
            => string.Equals(NormalizeServer(string.IsNullOrWhiteSpace(left) ? fallback : left),
                NormalizeServer(string.IsNullOrWhiteSpace(right) ? fallback : right), StringComparison.OrdinalIgnoreCase);

        private static string NormalizeServer(string server)
        {
            server = (server ?? string.Empty).Trim();
            // A SQL connection's optional TCP port is not part of sysjobhistory.server.
            int separator = server.LastIndexOf(',');
            int port;
            if (separator >= 0 && int.TryParse(server.Substring(separator + 1), NumberStyles.None,
                CultureInfo.InvariantCulture, out port)) server = server.Substring(0, separator).TrimEnd();
            return server;
        }
    }

    internal sealed class JobQuickViewExecutionRow : INotifyPropertyChanged
    {
        private TimeSpan? observedDuration;
        private TimeSpan? elapsed;
        private DateTime? activityStartedAt;
        private DateTime? activityStoppedAt;
        private int? completedHistoryId;
        public JobQuickViewHistoryItem Item { get; private set; }
        public bool IsActive { get; }
        public bool IsStale { get; private set; }
        public bool IsAnimating => IsActive && !IsStale;
        internal int HistoryBoundary { get; }
        internal int? AgentSessionId { get; private set; }
        public string Title => IsActive
            ? (IsStale ? "Last known active run" : "Running now") +
                (Item.StartedAt.HasValue ? " - " + Date(Item.StartedAt) : string.Empty)
            : Date(Item.StartedAt);
        public string EntryLabel => IsActive ? "Current execution" : Item.StepId == 0
            ? "Job outcome" : "Step " + Item.StepId + ": " + Value(Item.StepName, "Unnamed step");
        public string Summary => (IsStale ? "Last known: " : string.Empty) + Value(Item.Outcome, "Unknown") +
            "  |  " + DurationText + (IsActive && elapsed.HasValue ? " elapsed" : string.Empty);
        public string DurationText => FormatDuration(IsActive ? elapsed : Item.Duration,
            IsActive ? "Elapsed time unavailable" : "Duration unavailable");
        public event PropertyChangedEventHandler PropertyChanged;

        public JobQuickViewExecutionRow(JobQuickViewHistoryItem item)
        {
            Item = item ?? throw new ArgumentNullException(nameof(item));
        }

        private JobQuickViewExecutionRow(JobQuickViewSnapshot snapshot, string serverName, int historyBoundary)
        {
            IsActive = true;
            HistoryBoundary = historyBoundary;
            Item = new JobQuickViewHistoryRun { RunStatus = 4, Server = serverName };
            UpdateActive(snapshot, serverName);
        }

        internal static JobQuickViewExecutionRow CreateActive(JobQuickViewSnapshot snapshot, string serverName, int historyBoundary)
            => new JobQuickViewExecutionRow(snapshot, serverName, historyBoundary);

        internal bool MatchesStoppedActivity(JobQuickViewSnapshot snapshot)
            => snapshot.ActivityStoppedAt.HasValue && snapshot.ActivityStartedAt == activityStartedAt &&
               snapshot.ActivityStoppedAt == activityStoppedAt && snapshot.LastCompletedHistoryId == completedHistoryId;

        internal void UpdateActive(JobQuickViewSnapshot snapshot, string serverName)
        {
            activityStartedAt = snapshot.ActivityStartedAt;
            activityStoppedAt = snapshot.ActivityStoppedAt;
            completedHistoryId = snapshot.LastCompletedHistoryId;
            AgentSessionId = snapshot.AgentSessionId ?? AgentSessionId;
            Item.StartedAt = snapshot.RunningSince ?? Item.StartedAt;
            Item.Server = serverName;
            Item.Outcome = Value(snapshot.ExecutionStatus, "Running");
            observedDuration = snapshot.RunningDuration;
            if (observedDuration < TimeSpan.Zero) observedDuration = TimeSpan.Zero;
            elapsed = observedDuration;
            Item.Duration = observedDuration;
            IsStale = false;
            Notify();
        }

        internal void UpdateItem(JobQuickViewHistoryItem item)
        {
            if (ReferenceEquals(Item, item)) return;
            bool unchanged = Item.InstanceId == item.InstanceId && Item.StepId == item.StepId &&
                Item.StartedAt == item.StartedAt && Item.Duration == item.Duration && Item.RunStatus == item.RunStatus &&
                Item.Outcome == item.Outcome && Item.StepName == item.StepName && Item.Server == item.Server &&
                Item.Message == item.Message && Item.SqlMessageId == item.SqlMessageId && Item.SqlSeverity == item.SqlSeverity;
            Item = item;
            if (!unchanged) Notify();
        }

        public void UpdateElapsed(TimeSpan elapsedSinceSample)
        {
            if (!IsActive || IsStale || !observedDuration.HasValue) return;
            long increment = Math.Max(0, elapsedSinceSample.Ticks);
            long initial = observedDuration.Value.Ticks;
            var next = TimeSpan.FromTicks(increment > TimeSpan.MaxValue.Ticks - initial
                ? TimeSpan.MaxValue.Ticks : initial + increment);
            if (next == elapsed) return;
            elapsed = next;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DurationText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Summary)));
        }

        internal void MarkActivityStale()
        {
            if (!IsActive || IsStale) return;
            IsStale = true;
            Notify();
        }

        private void Notify()
        {
            // The active item's individual fields are not observable, so also refresh Item bindings.
            foreach (string property in new[] { nameof(Item), nameof(Title), nameof(Summary), nameof(DurationText),
                nameof(EntryLabel), nameof(IsStale), nameof(IsAnimating) })
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
        }

        private static string Date(DateTime? date)
            => date.HasValue ? date.Value.ToString("G", CultureInfo.CurrentCulture) : "Unknown start time";
        private static string Value(string text, string fallback) => string.IsNullOrWhiteSpace(text) ? fallback : text;
        private static string FormatDuration(TimeSpan? duration, string fallback)
            => duration.HasValue ? ((long)duration.Value.TotalHours).ToString(CultureInfo.InvariantCulture) +
                duration.Value.ToString(@"\:mm\:ss", CultureInfo.InvariantCulture) : fallback;
    }
}
