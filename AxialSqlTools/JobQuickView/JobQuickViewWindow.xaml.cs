using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Search;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace AxialSqlTools.JobQuickView
{
    public partial class JobQuickViewWindow : Window
    {
        private readonly JobQuickViewService service;
        private readonly string serverName;
        private readonly string requestedJobName;
        private readonly CancellationTokenSource lifetime = new CancellationTokenSource();
        private readonly CancellationToken token;
        private readonly DispatcherTimer refreshTimer;
        private readonly ToolWindowThemeController theme;
        private readonly SearchPanel searchPanel;
        private readonly ObservableCollection<StepDraft> drafts = new ObservableCollection<StepDraft>();
        private JobQuickViewSnapshot snapshot;
        private StepDraft selectedDraft;
        private bool loaded;
        private bool closed;
        private bool refreshing;
        private bool busy;
        private bool saving;
        private bool mutating;
        private bool applyingSnapshot;
        private bool refreshMessage;

        internal JobQuickViewWindow(JobQuickViewService service, string serverName, string jobName)
        {
            this.service = service ?? throw new ArgumentNullException(nameof(service));
            this.serverName = serverName ?? string.Empty;
            requestedJobName = jobName ?? throw new ArgumentNullException(nameof(jobName));
            token = lifetime.Token;
            InitializeComponent();
            JobNameText.Text = requestedJobName;
            ServerText.Text = this.serverName;
            Title = requestedJobName + " - Job Quick View";
            StepsList.ItemsSource = drafts;
            CommandEditor.Options.ConvertTabsToSpaces = false;
            CommandEditor.Options.IndentationSize = 4;
            CommandEditor.TextArea.Caret.PositionChanged += Caret_PositionChanged;
            searchPanel = SearchPanel.Install(CommandEditor);
            searchPanel.SetResourceReference(Control.BackgroundProperty, "AxialThemeHeaderBackgroundBrush");
            searchPanel.SetResourceReference(Control.ForegroundProperty, "AxialThemeForegroundBrush");
            theme = new ToolWindowThemeController(this, ApplyTheme);
            refreshTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10) };
            refreshTimer.Tick += RefreshTimer_Tick;
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (loaded) return;
            loaded = true;
            await RefreshAsync(false);
            if (!closed) refreshTimer.Start();
        }

        private async void RefreshTimer_Tick(object sender, EventArgs e)
        {
            if (IsVisible && !busy && !refreshing) await RefreshAsync(true);
        }

        private async Task<bool> RefreshAsync(bool automatic)
        {
            if (closed || busy || refreshing) return false;
            refreshing = true;
            RefreshText.Text = snapshot == null ? "Loading job information..." : "Refreshing...";
            UpdateActions();
            try
            {
                var result = snapshot == null
                    ? await service.LoadAsync(requestedJobName, token)
                    : await service.LoadAsync(snapshot.JobId, token);
                if (closed) return false;
                ApplySnapshot(result);
                RefreshText.Text = "Updated " + DateTime.Now.ToString("T", CultureInfo.CurrentCulture);
                if (!string.IsNullOrWhiteSpace(result.ActivityWarning))
                {
                    ShowMessage(result.ActivityWarning, false);
                    refreshMessage = true;
                }
                else if (refreshMessage || !automatic)
                {
                    MessageBanner.Visibility = Visibility.Collapsed;
                    refreshMessage = false;
                }
                return true;
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested) { return false; }
            catch (Exception ex)
            {
                if (!closed)
                {
                    RefreshText.Text = "Refresh failed";
                    ShowMessage("Could not refresh this job. " + ex.Message +
                        (snapshot == null ? " Use Refresh to retry." : " Displayed information may be out of date; drafts are preserved."), true);
                    refreshMessage = true;
                    if (snapshot == null)
                    {
                        ScheduleSummaryText.Text = "Unavailable";
                        LastRunSummaryText.Text = "Unavailable";
                        NextRunText.Text = "Unavailable";
                        EnabledText.Text = "Unknown";
                        RunningText.Text = "Status unknown";
                        EditorEmptyText.Text = "Job steps could not be loaded. Use Refresh to retry.";
                    }
                }
                return false;
            }
            finally
            {
                refreshing = false;
                if (!closed) UpdateActions();
            }
        }

        private void ApplySnapshot(JobQuickViewSnapshot result)
        {
            snapshot = result;
            JobNameText.Text = result.Name;
            JobNameText.ToolTip = result.Name;
            ServerText.Text = serverName + "  |  Owner: " + EmptyValue(result.Owner);
            EnabledText.Text = result.IsEnabled ? "Enabled" : "Disabled";
            RunningText.Text = EmptyValue(result.ExecutionStatus, "Status unknown");
            RunningText.ToolTip = result.RunningSince.HasValue ? "Started " + ServerDate(result.RunningSince) + " (server local time)" : null;
            ScheduleSummaryText.Text = !result.SchedulesLoaded ? "Unavailable" : result.Schedules.Count == 0 ? "On demand" :
                result.Schedules.Count(s => s.IsEnabled) + " enabled / " + result.Schedules.Count + " total";
            ScheduleDetailText.Text = !result.SchedulesLoaded ? "Schedule information could not be loaded" : result.Schedules.Count == 0 ? "No schedules attached" : string.Join(", ", result.Schedules.Select(s => s.Name));
            ScheduleDetailText.ToolTip = ScheduleDetailText.Text;
            LastRunSummaryText.Text = EmptyValue(result.LastRunOutcome, "No recorded execution");
            LastRunSummaryText.SetResourceReference(TextBlock.ForegroundProperty,
                string.Equals(result.LastRunOutcome, "Failed", StringComparison.OrdinalIgnoreCase)
                    ? "AxialThemeStatusErrorBrush" : "AxialThemeForegroundBrush");
            LastRunDetailText.Text = result.LastRunStarted.HasValue
                ? ServerDate(result.LastRunStarted) + "  |  " + Duration(result.LastRunDuration)
                : result.HistoryLoaded ? "No completed job history is available" : "Execution details could not be loaded";
            NextRunText.Text = !result.IsEnabled ? "Job disabled" : ServerDate(result.NextRun, "No scheduled time");
            NextRunDetailText.Text = !result.IsEnabled ? "Enable the job to allow scheduled runs" : "SQL Server local time; Agent cache may lag";
            JobDetailsText.Text = "Job: " + result.Name + "\r\nServer: " + serverName + "\r\nOwner: " + EmptyValue(result.Owner) +
                "\r\nStarting step: " + result.StartStepId + "\r\nJob ID: " + result.JobId + "\r\n\r\n" + EmptyValue(result.Description, "No description.");
            var scheduleDetails = new StringBuilder();
            foreach (var schedule in result.Schedules)
            {
                if (scheduleDetails.Length > 0) scheduleDetails.AppendLine().AppendLine();
                scheduleDetails.Append(schedule.Name).Append(schedule.IsEnabled ? " (enabled)" : " (disabled)")
                    .AppendLine().Append(EmptyValue(schedule.Description))
                    .AppendLine().Append("Next run: ").Append(ServerDate(schedule.NextRun, schedule.IsEnabled ? "No scheduled time" : "Schedule disabled"));
            }
            SchedulesText.Text = (!result.SchedulesLoaded ? "Schedules could not be loaded. Use Refresh to retry." : scheduleDetails.Length == 0 ? "No schedules attached. This job can be started manually." : scheduleDetails.ToString()) +
                "\r\n\r\n" + JobQuickViewSnapshot.NextRunNote;
            LastRunMessageText.Text = "Outcome: " + EmptyValue(result.LastRunOutcome, "No recorded execution") +
                "\r\nStarted: " + ServerDate(result.LastRunStarted) + "\r\nDuration: " + Duration(result.LastRunDuration) +
                "\r\n\r\n" + EmptyValue(result.LastRunMessage, result.HistoryLoaded ? "No execution message is available." : "Execution details could not be loaded. Use Refresh to retry.");

            // Keep the original server command as the concurrency baseline for every draft.
            // A server refresh may update clean documents, but never replaces unsaved work.
            var previousSelection = selectedDraft;
            applyingSnapshot = true;
            try
            {
                var existing = drafts.ToDictionary(d => d.Original.StepUid);
                var next = new System.Collections.Generic.List<StepDraft>();
                foreach (var step in result.Steps)
                {
                    if (existing.TryGetValue(step.StepUid, out var draft)) draft.Merge(step);
                    else
                    {
                        draft = new StepDraft(step);
                        draft.PropertyChanged += Draft_Changed;
                    }
                    next.Add(draft);
                    existing.Remove(step.StepUid);
                }
                foreach (var removed in existing.Values)
                {
                    if (removed.IsDirty)
                    {
                        removed.MarkRemoved();
                        next.Add(removed);
                    }
                    else removed.PropertyChanged -= Draft_Changed;
                }
                if (!drafts.SequenceEqual(next))
                {
                    drafts.Clear();
                    foreach (var draft in next) drafts.Add(draft);
                }
                StepsList.SelectedItem = previousSelection != null && drafts.Contains(previousSelection)
                    ? previousSelection : drafts.FirstOrDefault(d => d.Original.StepId == result.StartStepId) ?? drafts.FirstOrDefault();
            }
            finally { applyingSnapshot = false; }
            SelectDraft(StepsList.SelectedItem as StepDraft);
            StepsCountText.Text = "JOB STEPS (" + result.Steps.Count + ")";
            FooterText.Text = "Schedule and execution times use the SQL Server's local time.";
            UpdateActions();
        }

        private void Steps_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!applyingSnapshot) SelectDraft(StepsList.SelectedItem as StepDraft);
        }

        private void SelectDraft(StepDraft draft)
        {
            if (!ReferenceEquals(selectedDraft, draft))
            {
                if (selectedDraft != null)
                {
                    selectedDraft.CaretOffset = CommandEditor.CaretOffset;
                    selectedDraft.VerticalOffset = CommandEditor.VerticalOffset;
                }
                selectedDraft = draft;
                CommandEditor.Document = draft?.Document ?? new TextDocument();
                if (draft != null)
                {
                    CommandEditor.CaretOffset = Math.Min(draft.CaretOffset, draft.Document.TextLength);
                    CommandEditor.ScrollToVerticalOffset(draft.VerticalOffset);
                }
            }
            EditorEmptyText.Visibility = draft == null ? Visibility.Visible : Visibility.Collapsed;
            EditorEmptyText.Text = snapshot == null ? "Loading job steps..." : "This job has no steps.";
            StepNameText.Text = draft == null ? "Select a job step" : draft.DisplayName;
            StepContextText.Text = draft?.Details ?? string.Empty;
            StepFlowText.Text = draft == null ? string.Empty : "On success: " + EmptyValue(draft.Latest.SuccessAction) +
                "    |    On failure: " + EmptyValue(draft.Latest.FailureAction);
            ApplyEditorTheme();
            UpdateActions();
        }

        private void Draft_Changed(object sender, PropertyChangedEventArgs e)
        {
            if (!closed && !applyingSnapshot) UpdateActions();
        }

        private void UpdateActions()
        {
            if (closed) return;
            bool canAct = snapshot != null && !busy && !refreshing;
            StartButton.IsEnabled = canAct && snapshot.IsRunning == false;
            StopButton.IsEnabled = canAct && snapshot.IsRunning == true;
            EnabledButton.IsEnabled = canAct;
            EnabledButton.Content = snapshot?.IsEnabled == false ? "Enable job" : "Disable job";
            RefreshButton.IsEnabled = !busy && !refreshing;
            StepsList.IsEnabled = !busy;
            CommandEditor.IsReadOnly = selectedDraft == null || busy;
            FormatButton.IsEnabled = !busy && !refreshing && selectedDraft?.Original.IsSql == true && !selectedDraft.IsRemoved;
            FindButton.IsEnabled = selectedDraft != null;
            SaveButton.IsEnabled = canAct && selectedDraft?.IsDirty == true && !selectedDraft.IsRemoved;
            DiscardButton.IsEnabled = !busy && selectedDraft?.IsDirty == true;
            if (selectedDraft == null) EditorStateText.Text = string.Empty;
            else if (selectedDraft.IsRemoved) EditorStateText.Text = "This step was removed on the server. Copy your draft before discarding it.";
            else if (selectedDraft.HasRemoteChanges) EditorStateText.Text = "The command or execution context changed on the server. Your draft is preserved; discard it to load that version.";
            else EditorStateText.Text = selectedDraft.IsDirty ? "Unsaved command changes" : "Saved command. Edit here, then save this step.";
            int dirtyCount = drafts.Count(d => d.IsDirty);
            Title = (dirtyCount == 0 ? string.Empty : "* ") + (snapshot?.Name ?? requestedJobName) + " - Job Quick View";
            Caret_PositionChanged(null, EventArgs.Empty);
        }

        private async void Refresh_Click(object sender, RoutedEventArgs e) { await RefreshAsync(false); }

        private async void Enabled_Click(object sender, RoutedEventArgs e)
        {
            if (snapshot == null) return;
            bool enabled = !snapshot.IsEnabled;
            await RunActionAsync(() => service.SetEnabledAsync(snapshot.JobId, enabled, token), enabled ? "Job enabled." : "Job disabled.");
        }

        private async void Start_Click(object sender, RoutedEventArgs e)
        {
            if (snapshot == null) return;
            if (drafts.Any(d => d.IsDirty) && MessageBox.Show(this,
                "There are unsaved step changes. Start the job using its currently saved commands?",
                "Start job", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes) return;
            await RunActionAsync(() => service.StartAsync(snapshot.JobId, token), "Start requested. SQL Server Agent will update the running status.");
        }

        private async void Stop_Click(object sender, RoutedEventArgs e)
        {
            if (snapshot == null) return;
            await RunActionAsync(() => service.StopAsync(snapshot.JobId, token), "Stop requested. SQL Server Agent may need time to finish stopping the job.");
        }

        private async Task RunActionAsync(Func<Task> action, string successMessage)
        {
            if (closed || busy || refreshing || snapshot == null) return;
            busy = true;
            mutating = true;
            UpdateActions();
            bool succeeded = false;
            try
            {
                await action();
                succeeded = !closed;
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested) { }
            catch (Exception ex) { if (!closed) ShowMessage(ex.Message, true); }
            finally
            {
                busy = false;
                mutating = false;
                if (!closed) UpdateActions();
            }
            if (succeeded)
            {
                bool refreshed = await RefreshAsync(false);
                if (closed) return;
                string warning = !refreshed ? MessageText.Text : snapshot.ActivityWarning;
                ShowMessage(successMessage + (string.IsNullOrWhiteSpace(warning) ? string.Empty : "\r\n" + warning), !refreshed);
                refreshMessage = !string.IsNullOrWhiteSpace(warning);
            }
        }

        private async void Save_Click(object sender, RoutedEventArgs e) { await SaveSelectedAsync(); }

        private async Task SaveSelectedAsync()
        {
            if (closed || busy || refreshing || snapshot == null || selectedDraft == null || !selectedDraft.IsDirty || selectedDraft.IsRemoved) return;
            var draft = selectedDraft;
            string command = draft.Document.Text;
            saving = true;
            try
            {
                await RunActionAsync(async () =>
                {
                    await service.SaveStepCommandAsync(snapshot.JobId, draft.Original, command, token);
                    if (!closed) draft.AcceptSaved(command);
                }, "Step " + draft.Original.StepId + " command saved.");
            }
            finally { saving = false; }
        }

        private void Discard_Click(object sender, RoutedEventArgs e)
        {
            if (busy || selectedDraft == null || !selectedDraft.IsDirty) return;
            var draft = selectedDraft;
            if (MessageBox.Show(this, "Discard unsaved changes to step '" + draft.Latest.Name + "'?",
                "Discard step changes", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes) return;
            if (draft.IsRemoved)
            {
                draft.PropertyChanged -= Draft_Changed;
                drafts.Remove(draft);
                StepsList.SelectedItem = drafts.FirstOrDefault();
            }
            else draft.Discard();
            SelectDraft(StepsList.SelectedItem as StepDraft);
        }

        private async void Format_Click(object sender, RoutedEventArgs e) { await FormatSelectedAsync(); }

        private async Task FormatSelectedAsync()
        {
            if (closed || busy || refreshing || selectedDraft?.Original.IsSql != true || selectedDraft.IsRemoved) return;
            var draft = selectedDraft;
            int start = CommandEditor.SelectionLength == 0 ? 0 : CommandEditor.SelectionStart;
            int length = CommandEditor.SelectionLength == 0 ? draft.Document.TextLength : CommandEditor.SelectionLength;
            string source = draft.Document.GetText(start, length);
            if (string.IsNullOrWhiteSpace(source)) return;
            busy = true;
            UpdateActions();
            try
            {
                var settings = SettingsManager.GetTSqlCodeFormatSettings();
                // Operational comments must survive a command-only edit. This settings copy is not persisted.
                settings.preserveComments = true;
                // Global settings work even when no ordinary SSMS query tab is open.
                var formatter = await SsmsFormatterHost.CreateAsync(null, token, settings.disregardSsmsFormatterSettings);
                string formatted = await Task.Run(() => TSqlFormatter.FormatCode(source, settings, formatter.Parser, formatter.Generator), token);
                token.ThrowIfCancellationRequested();
                if (!string.Equals(source, formatted, StringComparison.Ordinal))
                {
                    using (draft.Document.RunUpdate()) draft.Document.Replace(start, length, formatted);
                    CommandEditor.Select(start, formatted.Length);
                }
                ShowMessage("Formatting applied to this draft. Save step to update the job command; Ctrl+Z undoes formatting.", false);
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested) { }
            catch (Exception ex)
            {
                if (!closed) ShowMessage("SQL could not be formatted; the command is unchanged. SQL Server Agent tokens or unsupported syntax may prevent formatting. " + ex.Message, true);
            }
            finally
            {
                busy = false;
                if (!closed) UpdateActions();
            }
        }

        private void Find_Click(object sender, RoutedEventArgs e)
        {
            if (selectedDraft == null) return;
            searchPanel.Open();
            searchPanel.Reactivate();
        }

        private async void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F5 && Keyboard.Modifiers == ModifierKeys.None)
            {
                e.Handled = true;
                await RefreshAsync(false);
            }
            else if (e.Key == Key.F && Keyboard.Modifiers == ModifierKeys.Control)
            {
                e.Handled = true;
                Find_Click(sender, e);
            }
            else if (e.Key == Key.S && Keyboard.Modifiers == ModifierKeys.Control)
            {
                e.Handled = true;
                await SaveSelectedAsync();
            }
            else if (e.Key == Key.F && Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
            {
                e.Handled = true;
                await FormatSelectedAsync();
            }
        }

        private void Editor_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.Control) return;
            CommandEditor.FontSize = Math.Max(9, Math.Min(28, CommandEditor.FontSize + (e.Delta > 0 ? 1 : -1)));
            e.Handled = true;
        }

        private void Caret_PositionChanged(object sender, EventArgs e)
        {
            if (CaretText == null) return;
            CaretText.Text = selectedDraft == null ? "Ctrl+F Find  |  Ctrl+S Save  |  Ctrl+wheel Zoom" :
                "Ln " + CommandEditor.TextArea.Caret.Line + ", Col " + CommandEditor.TextArea.Caret.Column +
                "  |  Ctrl+F Find  |  Ctrl+S Save";
        }

        private void ApplyTheme()
        {
            ToolWindowThemeResources.ApplySharedTheme(this);
            ApplyEditorTheme();
        }

        private void ApplyEditorTheme()
        {
            JobCommandEditorSupport.ApplyTheme(CommandEditor, this, selectedDraft?.Latest.Subsystem);
        }

        private void ShowMessage(string message, bool error)
        {
            refreshMessage = false;
            MessageText.Text = message;
            MessageBanner.SetResourceReference(Border.BorderBrushProperty, error ? "AxialThemeStatusErrorBrush" : "AxialThemeAccentBrush");
            MessageText.SetResourceReference(TextBlock.ForegroundProperty, error ? "AxialThemeStatusErrorBrush" : "AxialThemeForegroundBrush");
            MessageBanner.Visibility = Visibility.Visible;
        }

        private void ContentGrid_LayoutUpdated(object sender, EventArgs e)
        {
            if (closed || ContentScroll.ViewportHeight <= 0 || double.IsInfinity(ContentScroll.ViewportHeight)) return;
            // Keep the editor's measure finite so large commands use its own scrollbar.
            // At small window sizes, the outer scrollbar makes expanded details and Save reachable.
            double minimum = ContentGrid.RowDefinitions[5].MinHeight;
            for (int i = 0; i < ContentGrid.RowDefinitions.Count; i++)
                if (i != 5) minimum += ContentGrid.RowDefinitions[i].ActualHeight;
            double height = Math.Max(ContentScroll.ViewportHeight, minimum);
            if (Math.Abs(ContentGrid.Height - height) > 0.5) ContentGrid.Height = height;
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (saving || mutating)
            {
                e.Cancel = true;
                ShowMessage("A job action is in progress. Wait for it to finish before closing this window.", false);
                return;
            }
            int count = drafts.Count(d => d.IsDirty);
            if (count > 0 && MessageBox.Show(this,
                "There are unsaved commands in " + count + " step(s). Close and discard those changes?",
                "Unsaved step changes", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) != MessageBoxResult.Yes) e.Cancel = true;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            closed = true;
            refreshTimer.Stop();
            refreshTimer.Tick -= RefreshTimer_Tick;
            lifetime.Cancel();
            lifetime.Dispose();
            theme.Dispose();
            searchPanel.Uninstall();
            CommandEditor.TextArea.Caret.PositionChanged -= Caret_PositionChanged;
            foreach (var draft in drafts) draft.PropertyChanged -= Draft_Changed;
        }

        private static string ServerDate(DateTime? value, string fallback = "Not available")
            => value.HasValue ? value.Value.ToString("g", CultureInfo.CurrentCulture) : fallback;

        private static string Duration(TimeSpan? value)
            => value.HasValue ? ((long)value.Value.TotalHours).ToString(CultureInfo.InvariantCulture) +
                value.Value.ToString(@"\:mm\:ss", CultureInfo.InvariantCulture) : "Duration unavailable";

        private static string EmptyValue(string value, string fallback = "Not available")
            => string.IsNullOrWhiteSpace(value) ? fallback : value;

        private sealed class StepDraft : INotifyPropertyChanged
        {
            private string baseline;
            public JobQuickViewStep Original { get; private set; }
            public JobQuickViewStep Latest { get; private set; }
            public TextDocument Document { get; }
            public int CaretOffset { get; set; }
            public double VerticalOffset { get; set; }
            public bool IsRemoved { get; private set; }
            public bool IsDirty => !string.Equals(baseline, Document.Text, StringComparison.Ordinal);
            public bool HasRemoteChanges => IsDirty &&
                (!string.Equals(baseline, Latest.Command ?? string.Empty, StringComparison.Ordinal) || Original.StepId != Latest.StepId ||
                 !string.Equals(Original.Subsystem, Latest.Subsystem, StringComparison.Ordinal) ||
                 !string.Equals(Original.DatabaseName, Latest.DatabaseName, StringComparison.Ordinal) ||
                 !string.Equals(Original.DatabaseUserName, Latest.DatabaseUserName, StringComparison.Ordinal) || Original.ProxyId != Latest.ProxyId);
            public string DisplayName => Latest.StepId + ". " + Latest.Name;
            public string Details => Latest.Subsystem + (string.IsNullOrWhiteSpace(Latest.DatabaseName) ? string.Empty : "  |  " + Latest.DatabaseName);
            public string EditState => IsRemoved ? "Removed on server - draft retained" : HasRemoteChanges ? "Changed on server - draft retained" : IsDirty ? "Unsaved changes" : string.Empty;
            public event PropertyChangedEventHandler PropertyChanged;

            public StepDraft(JobQuickViewStep step)
            {
                Original = Latest = step;
                baseline = step.Command ?? string.Empty;
                Document = new TextDocument(baseline);
                Document.TextChanged += (sender, args) => Notify();
            }

            public void Merge(JobQuickViewStep step)
            {
                bool dirty = IsDirty;
                Latest = step;
                IsRemoved = false;
                if (!dirty) ResetToLatest();
                Notify();
            }

            public void AcceptSaved(string command)
            {
                Original.Command = command;
                Latest = Original;
                baseline = command;
                Notify();
            }

            public void Discard()
            {
                ResetToLatest();
                Notify();
            }

            public void MarkRemoved() { IsRemoved = true; Notify(); }

            private void ResetToLatest()
            {
                Original = Latest;
                baseline = Latest.Command ?? string.Empty;
                if (!string.Equals(Document.Text, baseline, StringComparison.Ordinal))
                {
                    Document.Text = baseline;
                    Document.UndoStack.ClearAll();
                }
            }

            private void Notify() { PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(string.Empty)); }
        }
    }
}
