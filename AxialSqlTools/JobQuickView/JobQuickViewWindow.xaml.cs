using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Search;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace AxialSqlTools.JobQuickView
{
    public partial class JobQuickViewWindow : UserControl, IDisposable
    {
        private readonly JobQuickViewService service;
        private readonly string serverName;
        private readonly string requestedJobName;
        private readonly CancellationTokenSource lifetime = new CancellationTokenSource();
        private readonly CancellationToken token;
        private readonly ToolWindowThemeController theme;
        private readonly SearchPanel searchPanel;
        private readonly ObservableCollection<StepDraft> drafts = new ObservableCollection<StepDraft>();
        private JobQuickViewSnapshot snapshot;
        private StepDraft selectedDraft;
        private HistoryNode selectedHistory;
        private HistoryNode[] historyNodes = Array.Empty<HistoryNode>();
        private bool loaded;
        private bool closed;
        private bool refreshing;
        private bool busy;
        private bool saving;
        private bool mutating;
        private bool applyingSnapshot;
        private bool stateUpdatePending;

        internal string Caption { get; private set; }
        internal string JobName => snapshot?.Name ?? requestedJobName;
        internal event EventHandler CaptionChanged;

        private void SetCaption(string value)
        {
            if (Caption == value) return;
            Caption = value;
            CaptionChanged?.Invoke(this, EventArgs.Empty);
        }

        internal JobQuickViewWindow(JobQuickViewService service, string serverName, string jobName)
        {
            this.service = service ?? throw new ArgumentNullException(nameof(service));
            this.serverName = serverName ?? string.Empty;
            requestedJobName = jobName ?? throw new ArgumentNullException(nameof(jobName));
            token = lifetime.Token;
            InitializeComponent();
            JobNameText.Text = requestedJobName;
            ServerText.Text = this.serverName;
            SetCaption(requestedJobName + " - Quick Manage");
            StepsList.ItemsSource = drafts;
            CommandEditor.Options.ConvertTabsToSpaces = false;
            CommandEditor.Options.IndentationSize = 4;
            CommandEditor.TextArea.Caret.PositionChanged += Caret_PositionChanged;
            searchPanel = SearchPanel.Install(CommandEditor);
            searchPanel.SetResourceReference(Control.BackgroundProperty, "AxialThemeHeaderBackgroundBrush");
            searchPanel.SetResourceReference(Control.ForegroundProperty, "AxialThemeForegroundBrush");
            theme = new ToolWindowThemeController(this, ApplyTheme);
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (loaded || closed) return;
            loaded = true;
            JobCommandEditorSupport.ApplyHostEditorFont(CommandEditor);
            await RefreshAsync();
        }

        private async Task<bool> RefreshAsync()
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
                }
                else
                {
                    MessageBanner.Visibility = Visibility.Collapsed;
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
                    if (snapshot == null)
                    {
                        ScheduleStatusText.Text = "Schedules could not be loaded. Use Refresh to retry.";
                        HistoryStatusText.Text = "Execution history unavailable";
                        HistoryEmptyText.Text = "Execution history could not be loaded. Use Refresh to retry.";
                        LastRunSummaryText.Text = "Unavailable";
                        NextRunText.Text = "Unavailable";
                        EnabledText.Text = "Unknown";
                        RunningText.Text = "Status unknown";
                        ApplyStatusBadges();
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
            bool initialLoad = snapshot == null;
            if (!initialLoad)
            {
                // Regular refresh does not query or replace the independently refreshed history.
                result.History = snapshot.History;
                result.HistoryLoaded = snapshot.HistoryLoaded;
                result.HistoryNote = snapshot.HistoryNote;
                if (result.LastRunStarted.HasValue && result.LastRunStarted == snapshot.LastRunStarted
                    && string.Equals(result.LastRunOutcome, snapshot.LastRunOutcome, StringComparison.Ordinal))
                {
                    result.LastRunDuration = snapshot.LastRunDuration;
                    result.LastRunMessage = snapshot.LastRunMessage;
                }
            }
            snapshot = result;
            JobNameText.Text = result.Name;
            JobNameText.ToolTip = result.Name;
            ServerText.Text = serverName + "  |  Owner: " + EmptyValue(result.Owner);
            ApplyExecutionSummary(result);
            JobDetailsGrid.ItemsSource = new[]
            {
                new DetailRow("Description", EmptyValue(result.Description, "No description.")),
                new DetailRow("Category", EmptyValue(result.Category)),
                new DetailRow("Created", ServerDate(result.CreatedAt)),
                new DetailRow("Modified", ServerDate(result.ModifiedAt)),
                new DetailRow("Job ID", result.JobId.ToString())
            };
            ScheduleStatusText.Text = !result.SchedulesLoaded ? "Schedules could not be loaded. Use Refresh to retry." :
                result.Schedules.Count == 0 ? "No schedules attached. This job can be started manually." :
                result.Schedules.Count(s => s.IsEnabled) + " enabled / " + result.Schedules.Count + " total";
            SchedulesGrid.ItemsSource = result.Schedules.Select(schedule => new ScheduleRow
            {
                Name = schedule.Name,
                Status = schedule.IsEnabled ? "Enabled" : "Disabled",
                Description = EmptyValue(schedule.Description),
                NextRun = ServerDate(schedule.NextRun, schedule.IsEnabled ? "No scheduled time" : "Schedule disabled")
            }).ToList();
            ScheduleNoteText.Text = JobQuickViewSnapshot.NextRunNote;
            if (initialLoad) ApplyHistory(result);
            else UpdateHistoryStatus(result);

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
                    draft.SetStartingStep(result.StartStepId);
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
            UpdateActions();
        }

        private void ApplyExecutionSummary(JobQuickViewSnapshot result)
        {
            EnabledText.Text = result.IsEnabled ? "Enabled" : "Disabled";
            RunningText.Text = EmptyValue(result.ExecutionStatus, "Status unknown");
            RunningText.ToolTip = result.RunningSince.HasValue ? "Started " + ServerDate(result.RunningSince) + " (server local time)" : null;
            ApplyStatusBadges();
            LastRunSummaryText.Text = EmptyValue(result.LastRunOutcome, "No recorded execution");
            LastRunSummaryText.SetResourceReference(TextBlock.ForegroundProperty,
                string.Equals(result.LastRunOutcome, "Failed", StringComparison.OrdinalIgnoreCase)
                    ? "AxialThemeStatusErrorBrush" : "AxialThemeForegroundBrush");
            LastRunDetailText.Text = result.LastRunStarted.HasValue
                ? ServerDate(result.LastRunStarted) + (result.LastRunDuration.HasValue ? "  |  " + Duration(result.LastRunDuration) : string.Empty)
                : result.HistoryLoaded ? "No completed job history is available" : "Execution details could not be loaded";
            NextRunText.Text = !result.IsEnabled ? "Job disabled" : ServerDate(result.NextRun, "No scheduled time");
            NextRunText.ToolTip = NextRunText.Text + "\n" + (!result.IsEnabled
                ? "Enable the job to allow scheduled runs" : JobQuickViewSnapshot.NextRunNote);
        }

        private async void HistoryRefresh_Click(object sender, RoutedEventArgs e)
        {
            if (closed || busy || refreshing || snapshot == null) return;
            refreshing = true;
            HistoryRefreshActionText.Text = "Refreshing...";
            UpdateActions();
            try
            {
                var result = await service.LoadExecutionAsync(snapshot.JobId, token);
                if (closed) return;
                bool replaceLoadWarning = MessageBanner.Visibility == Visibility.Visible
                    && !string.IsNullOrWhiteSpace(snapshot.ActivityWarning)
                    && string.Equals(MessageText.Text, snapshot.ActivityWarning, StringComparison.Ordinal);
                // Update execution state only. Step documents, schedules and job details stay intact.
                snapshot.IsEnabled = result.IsEnabled;
                snapshot.IsRunning = result.IsRunning;
                snapshot.ExecutionStatus = result.ExecutionStatus;
                snapshot.RunningSince = result.RunningSince;
                snapshot.NextRun = result.NextRun;
                snapshot.HistoryLoaded = result.HistoryLoaded;
                if (result.HistoryLoaded)
                {
                    snapshot.History = result.History;
                    snapshot.HistoryNote = result.HistoryNote;
                    snapshot.LastRunStarted = result.LastRunStarted;
                    snapshot.LastRunDuration = result.LastRunDuration;
                    snapshot.LastRunOutcome = result.LastRunOutcome;
                    snapshot.LastRunMessage = result.LastRunMessage;
                }
                ApplyExecutionSummary(snapshot);
                ApplyHistory(snapshot);
                if (!string.IsNullOrWhiteSpace(result.ActivityWarning))
                    HistoryNoteText.Text += Environment.NewLine + result.ActivityWarning;
                snapshot.ActivityWarning = string.Join(Environment.NewLine, new[]
                {
                    snapshot.SchedulesLoaded ? null : ScheduleStatusText.Text,
                    result.ActivityWarning
                }.Where(warning => !string.IsNullOrWhiteSpace(warning)));
                // Retire an old load warning after recovery, preserving unrelated action messages.
                if (replaceLoadWarning)
                {
                    if (string.IsNullOrWhiteSpace(snapshot.ActivityWarning)) MessageBanner.Visibility = Visibility.Collapsed;
                    else ShowMessage(snapshot.ActivityWarning, false);
                }
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested) { }
            catch (Exception ex)
            {
                if (!closed)
                {
                    snapshot.HistoryLoaded = false;
                    ApplyHistory(snapshot);
                    HistoryNoteText.Text += Environment.NewLine + ex.Message;
                }
            }
            finally
            {
                refreshing = false;
                if (!closed)
                {
                    HistoryRefreshActionText.Text = "Refresh history";
                    UpdateActions();
                }
            }
        }

        private void UpdateHistoryStatus(JobQuickViewSnapshot result)
        {
            string currentStatus = "Current status: " + EmptyValue(result.ExecutionStatus, "Unknown") +
                (result.RunningSince.HasValue ? " (started " + HistoryDate(result.RunningSince) + ")" : string.Empty);
            HistoryStatusText.Text = result.HistoryLoaded
                ? result.History.Count + (result.History.Count == 1 ? " run" : " runs") + "  |  " + currentStatus
                : currentStatus + "  |  History refresh failed";
        }

        private void ApplyHistory(JobQuickViewSnapshot result)
        {
            UpdateHistoryStatus(result);
            if (!result.HistoryLoaded)
            {
                HistoryNoteText.Text = historyNodes.Length == 0
                    ? "Execution history could not be loaded. Use Refresh history to retry."
                    : "Showing previously loaded history. Use Refresh history to retry. Times use the SQL Server's local time.";
                HistoryEmptyText.Text = "Execution history could not be loaded. Use Refresh history to retry.";
                HistoryEmptyText.Visibility = historyNodes.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
                return;
            }

            var previousSelection = selectedHistory;
            var expandedRuns = historyNodes.Where(node => node.IsExpanded).Select(node => node.Item.InstanceId).ToArray();
            bool firstHistory = historyNodes.Length == 0;
            historyNodes = result.History.Select(run => new HistoryNode(run)).ToArray();
            foreach (var node in historyNodes)
                node.IsExpanded = expandedRuns.Contains(node.Item.InstanceId) || (firstHistory && node == historyNodes[0]);

            var selection = previousSelection == null ? null : historyNodes
                .SelectMany(node => new[] { node }.Concat(node.Children))
                .FirstOrDefault(node => node.Item.InstanceId == previousSelection.Item.InstanceId);
            if (selection == null && previousSelection != null)
                selection = historyNodes.FirstOrDefault(node => node.Item.InstanceId == previousSelection.Run.InstanceId);
            selection = selection ?? historyNodes.FirstOrDefault();
            if (selection != null)
            {
                selection.IsSelected = true;
                if (selection.Item.StepId != 0)
                    historyNodes.First(node => node.Run.InstanceId == selection.Run.InstanceId).IsExpanded = true;
            }
            HistoryTree.ItemsSource = historyNodes;
            ShowHistorySelection(selection);
            HistoryEmptyText.Text = "No retained execution history. This job may not have run, or its history was purged.";
            HistoryEmptyText.Visibility = historyNodes.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
            HistoryNoteText.Text = EmptyValue(result.HistoryNote,
                "Expand a run to see its recorded step attempts. Times use the SQL Server's local time.");
        }

        private void History_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (!closed && e.NewValue is HistoryNode node) ShowHistorySelection(node);
        }

        private void ShowHistorySelection(HistoryNode node)
        {
            selectedHistory = node;
            HistoryTitleText.Text = node?.Title ?? "Select a run or step";
            HistoryDetailText.Text = node == null ? string.Empty : EmptyValue(node.Item.Outcome) +
                "  |  Started " + HistoryDate(node.Item.StartedAt) + "  |  Duration " + Duration(node.Item.Duration);
            HistoryContextText.Text = node == null ? string.Empty :
                (node.Item.StepId == 0 ? string.Empty : "Run " + HistoryDate(node.Run.StartedAt) + "  |  ") +
                "Server: " + EmptyValue(node.Item.Server, serverName) +
                (node.Item.SqlMessageId != 0 || node.Item.SqlSeverity != 0
                    ? "  |  SQL message " + node.Item.SqlMessageId + ", severity " + node.Item.SqlSeverity : string.Empty);
            HistoryContextText.Visibility = string.IsNullOrEmpty(HistoryContextText.Text) ? Visibility.Collapsed : Visibility.Visible;
            HistorySelectionNoteText.Text = node?.Run.Note ?? string.Empty;
            HistorySelectionNoteText.Visibility = string.IsNullOrWhiteSpace(HistorySelectionNoteText.Text) ? Visibility.Collapsed : Visibility.Visible;
            HistoryMessageText.Text = node == null ? string.Empty : EmptyValue(node.Item.Message, "No message was recorded for this entry.");
            HistoryMessageText.ScrollToHome();
        }

        private static string HistoryDate(DateTime? value)
            => value.HasValue ? value.Value.ToString("G", CultureInfo.CurrentCulture) : "Unknown start time";

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
            // The editor footer and step list observe the same draft, including format/undo/redo.
            EditorFooter.DataContext = draft;
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
            if (closed || applyingSnapshot || stateUpdatePending) return;
            stateUpdatePending = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                stateUpdatePending = false;
                if (!closed) UpdateActions();
            }), DispatcherPriority.DataBind);
        }

        private void UpdateActions()
        {
            if (closed) return;
            bool canAct = snapshot != null && !busy && !refreshing;
            StartButton.IsEnabled = canAct && snapshot.IsRunning == false;
            StopButton.IsEnabled = canAct && snapshot.IsRunning == true;
            EnabledButton.IsEnabled = canAct;
            EnabledActionText.Text = snapshot?.IsEnabled == false ? "Enable job" : "Disable job";
            EnabledIcon.SetResourceReference(System.Windows.Shapes.Path.DataProperty,
                snapshot?.IsEnabled == false ? "EnableIconGeometry" : "DisableIconGeometry");
            RefreshButton.IsEnabled = !busy && !refreshing;
            HistoryRefreshButton.IsEnabled = canAct;
            StepsList.IsEnabled = !busy;
            CommandEditor.IsReadOnly = selectedDraft == null || busy;
            FormatButton.IsEnabled = !busy && !refreshing && selectedDraft?.Original.IsSql == true && !selectedDraft.IsRemoved;
            FindButton.IsEnabled = selectedDraft != null;
            // Child buttons bind to StepDraft.CanSave/IsDirty. Do not overwrite those bindings.
            EditorActions.IsEnabled = canAct;
            int dirtyCount = drafts.Count(d => d.IsDirty);
            SetCaption((dirtyCount == 0 ? string.Empty : "* ") + JobName + " - Quick Manage");
            Caret_PositionChanged(null, EventArgs.Empty);
        }

        private async void Refresh_Click(object sender, RoutedEventArgs e) { await RefreshAsync(); }

        private async void Enabled_Click(object sender, RoutedEventArgs e)
        {
            if (snapshot == null) return;
            bool enabled = !snapshot.IsEnabled;
            await RunActionAsync(() => service.SetEnabledAsync(snapshot.JobId, enabled, token), enabled ? "Job enabled." : "Job disabled.");
        }

        private async void Start_Click(object sender, RoutedEventArgs e)
        {
            if (snapshot == null) return;
            if (drafts.Any(d => d.IsDirty) && !Confirm(
                "There are unsaved step changes. Start the job using its currently saved commands?",
                "Start job")) return;
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
                bool refreshed = await RefreshAsync();
                if (closed) return;
                string warning = !refreshed ? MessageText.Text : snapshot.ActivityWarning;
                ShowMessage(successMessage + (string.IsNullOrWhiteSpace(warning) ? string.Empty : "\r\n" + warning), !refreshed);
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
            if (!Confirm("Discard unsaved changes to step '" + draft.Latest.Name + "'?", "Discard step changes")) return;
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
                draft.NotifyStateChanged();
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
            MainTabs.SelectedItem = StepsTab;
            searchPanel.Open();
            searchPanel.Reactivate();
        }

        private async void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F5 && Keyboard.Modifiers == ModifierKeys.None)
            {
                e.Handled = true;
                await RefreshAsync();
            }
            else if (e.Key == Key.F && Keyboard.Modifiers == ModifierKeys.Control)
            {
                e.Handled = true;
                Find_Click(sender, e);
            }
            else if (e.Key == Key.S && Keyboard.Modifiers == ModifierKeys.Control)
            {
                e.Handled = true;
                MainTabs.SelectedItem = StepsTab;
                await SaveSelectedAsync();
            }
            else if (e.Key == Key.F && Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
            {
                e.Handled = true;
                MainTabs.SelectedItem = StepsTab;
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
            ApplyHistoryTreeStyle();
            ApplySemanticColors();
            ApplyStatusBadges();
            ApplyEditorTheme();
        }

        private void ApplyHistoryTreeStyle()
        {
            // Keep the shell's expander, focus and selection treatment in every theme.
            var hostStyle = HistoryTree.TryFindResource(typeof(TreeViewItem)) as Style;
            if (HistoryTree.ItemContainerStyle != null && HistoryTree.ItemContainerStyle.BasedOn == hostStyle) return;
            var style = new Style(typeof(TreeViewItem), hostStyle);
            style.Setters.Add(new Setter(TreeViewItem.IsExpandedProperty, new Binding("IsExpanded") { Mode = BindingMode.TwoWay }));
            style.Setters.Add(new Setter(TreeViewItem.IsSelectedProperty, new Binding("IsSelected") { Mode = BindingMode.TwoWay }));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(5)));
            style.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch));
            HistoryTree.ItemContainerStyle = style;
        }

        private void ApplySemanticColors()
        {
            var background = VsThemeBrushResolver.GetBrushColor(Background, SystemColors.WindowColor);
            var foreground = VsThemeBrushResolver.GetBrushColor(Foreground, SystemColors.WindowTextColor);
            bool light = VsThemeBrushResolver.GetRelativeLuminance(background) > 0.6;
            SetSemanticPalette("Green", Color.FromRgb(0x12, 0x7A, 0x3A), background, foreground, light);
            SetSemanticPalette("Blue", Color.FromRgb(0x00, 0x67, 0xB8), background, foreground, light);
            SetSemanticPalette("Amber", Color.FromRgb(0xAB, 0x62, 0x00), background, foreground, light);
            var red = light ? Color.FromRgb(0xB4, 0x23, 0x18) : Color.FromRgb(0xFF, 0x8A, 0x80);
            var dirtyBackground = VsThemeBrushResolver.BlendColors(background, red, light ? 0.09 : 0.14);
            Resources["QuickManageDirtyBackgroundBrush"] = SystemParameters.HighContrast ? SystemColors.WindowBrush : new SolidColorBrush(dirtyBackground);
            Resources["QuickManageDirtyForegroundBrush"] = SystemParameters.HighContrast ? SystemColors.WindowTextBrush :
                new SolidColorBrush(VsThemeBrushResolver.EnsureTextContrast(red, dirtyBackground, foreground));
            Resources["QuickManageDirtyTextBrush"] = SystemParameters.HighContrast ? SystemColors.WindowTextBrush :
                new SolidColorBrush(VsThemeBrushResolver.EnsureTextContrast(red, background, foreground));
        }

        private void SetSemanticPalette(string name, Color color, Color background, Color foreground, bool light)
        {
            var fill = VsThemeBrushResolver.BlendColors(background, color, light ? 0.12 : 0.26);
            Resources["QuickManage" + name + "BackgroundBrush"] = SystemParameters.HighContrast ? SystemColors.WindowBrush : new SolidColorBrush(fill);
            Resources["QuickManage" + name + "ForegroundBrush"] = SystemParameters.HighContrast ? SystemColors.WindowTextBrush :
                new SolidColorBrush(VsThemeBrushResolver.EnsureTextContrast(color, fill, foreground));
        }

        private void ApplyStatusBadges()
        {
            SetStatusBadge(EnabledBadge, EnabledText, snapshot?.IsEnabled == true ? "Green" : "Amber");
            SetStatusBadge(RunningBadge, RunningText, snapshot?.IsRunning == true ? "Green" : snapshot?.IsRunning == false ? "Blue" : "Amber");
        }

        private static void SetStatusBadge(Border badge, TextBlock text, string color)
        {
            badge.SetResourceReference(Border.BackgroundProperty, "QuickManage" + color + "BackgroundBrush");
            badge.SetResourceReference(Border.BorderBrushProperty, "QuickManage" + color + "ForegroundBrush");
            text.SetResourceReference(TextBlock.ForegroundProperty, "QuickManage" + color + "ForegroundBrush");
        }

        private void ApplyEditorTheme()
        {
            JobCommandEditorSupport.ApplyTheme(CommandEditor, this, selectedDraft?.Latest.Subsystem);
        }

        private void ShowMessage(string message, bool error)
        {
            MessageText.Text = message;
            MessageBanner.SetResourceReference(Border.BorderBrushProperty, error ? "AxialThemeStatusErrorBrush" : "AxialThemeAccentBrush");
            MessageText.SetResourceReference(TextBlock.ForegroundProperty, error ? "AxialThemeStatusErrorBrush" : "AxialThemeForegroundBrush");
            MessageBanner.Visibility = Visibility.Visible;
        }

        private static bool Confirm(string message, string title, bool warning = false)
        {
            // The pane is hosted by SSMS, not by a standalone WPF Window.
            return VsShellUtilities.ShowMessageBox(ServiceProvider.GlobalProvider, message, title,
                warning ? OLEMSGICON.OLEMSGICON_WARNING : OLEMSGICON.OLEMSGICON_QUERY,
                OLEMSGBUTTON.OLEMSGBUTTON_YESNO, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_SECOND) == (int)System.Windows.Forms.DialogResult.Yes;
        }

        private void DocumentLayout_LayoutUpdated(object sender, EventArgs e)
        {
            if (closed || DocumentScroll.ViewportHeight <= 0 || DocumentScroll.ViewportWidth <= 0) return;
            // Keep AvalonEdit's measure finite; narrow/split document groups scroll the shell content.
            double width = Math.Max(860, DocumentScroll.ViewportWidth);
            double height = Math.Max(600, DocumentScroll.ViewportHeight);
            if (!double.IsInfinity(width) && Math.Abs(DocumentLayout.Width - width) > 0.5) DocumentLayout.Width = width;
            if (!double.IsInfinity(height) && Math.Abs(DocumentLayout.Height - height) > 0.5) DocumentLayout.Height = height;
        }

        internal bool CanClose()
        {
            if (closed) return true;
            if (saving || mutating)
            {
                ShowMessage("A job action is in progress. Wait for it to finish before closing this tab.", false);
                return false;
            }
            int count = drafts.Count(d => d.IsDirty);
            return count == 0 || Confirm("There are unsaved commands in " + count + " step(s). Close and discard those changes?",
                "Unsaved step changes", warning: true);
        }

        public void Dispose()
        {
            if (closed) return;
            closed = true;
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

        private sealed class DetailRow
        {
            public string Field { get; }
            public string Value { get; }
            public DetailRow(string field, string value) { Field = field; Value = value; }
        }

        private sealed class ScheduleRow
        {
            public string Name { get; set; }
            public string Status { get; set; }
            public string Description { get; set; }
            public string NextRun { get; set; }
        }

        private sealed class HistoryNode
        {
            public JobQuickViewHistoryItem Item { get; }
            public JobQuickViewHistoryRun Run { get; }
            public HistoryNode[] Children { get; }
            public bool IsExpanded { get; set; }
            public bool IsSelected { get; set; }
            public string Title => Item.StepId == 0 ? HistoryDate(Item.StartedAt) :
                "Step " + Item.StepId + ": " + EmptyValue(Item.StepName, "Unnamed step");
            public string Summary => EmptyValue(Item.Outcome) + "  |  " + Duration(Item.Duration) +
                (Item.StepId == 0 ? "  |  " + Children.Length + (Children.Length == 1 ? " step record" : " step records") : string.Empty);

            public HistoryNode(JobQuickViewHistoryRun run)
            {
                Item = Run = run;
                Children = run.Steps.Select(step => new HistoryNode(step, run)).ToArray();
            }

            private HistoryNode(JobQuickViewHistoryItem item, JobQuickViewHistoryRun run)
            {
                Item = item;
                Run = run;
                Children = Array.Empty<HistoryNode>();
            }
        }

        private sealed class StepDraft : INotifyPropertyChanged
        {
            private string baseline;
            public JobQuickViewStep Original { get; private set; }
            public JobQuickViewStep Latest { get; private set; }
            public TextDocument Document { get; }
            public int CaretOffset { get; set; }
            public double VerticalOffset { get; set; }
            public bool IsRemoved { get; private set; }
            public bool IsStartingStep { get; private set; }
            public bool IsDirty => !string.Equals(baseline, Document.Text, StringComparison.Ordinal);
            public bool CanSave => IsDirty && !IsRemoved;
            public string EditorStatus => IsRemoved ? "This step was removed on the server. Copy your draft before discarding it." :
                HasRemoteChanges ? "The command or execution context changed on the server. Your draft is preserved; discard it to load that version." :
                IsDirty ? "Unsaved command changes" : "Saved command. Edit here, then save this step.";
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

            public void SetStartingStep(int startStepId)
            {
                bool isStartingStep = !IsRemoved && Latest.StepId == startStepId;
                if (IsStartingStep == isStartingStep) return;
                IsStartingStep = isStartingStep;
                Notify();
            }

            public void MarkRemoved() { IsRemoved = true; IsStartingStep = false; Notify(); }
            public void NotifyStateChanged() { Notify(); }

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
