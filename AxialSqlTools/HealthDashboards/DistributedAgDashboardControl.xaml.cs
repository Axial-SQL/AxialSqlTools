using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace AxialSqlTools
{
    public partial class DistributedAgDashboardControl : UserControl, IDisposable
    {
        private readonly ToolWindowThemeController themeController;
        private readonly DispatcherTimer refreshTimer;
        private CancellationTokenSource refreshCancellation;
        private string connectionString;
        private string selectedGroup;
        private bool groupsLoaded;
        private bool changingSelection;
        private bool refreshOnResume;
        private bool disposed;
        private DistributedAgHealthSnapshot snapshot;

        public DistributedAgDashboardControl()
        {
            InitializeComponent();
            themeController = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));
            refreshTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
            refreshTimer.Tick += TimerTick;
            Loaded += ControlLoaded;
            Unloaded += ControlUnloaded;
            ShowUnavailable("Not refreshed", "Open an availability group to load its health.");
        }

        public void Initialize(ScriptFactoryAccess.ConnectionInfo connection, string availabilityGroupName)
        {
            if (connection == null || string.IsNullOrWhiteSpace(connection.FullConnectionString))
                throw new ArgumentException("Select a connected SQL Server instance in Object Explorer.", nameof(connection));
            connectionString = connection.FullConnectionString;
            selectedGroup = availabilityGroupName;
            ConnectionLabel.Text = "Connected instance: " + connection.ServerName;
            RefreshButton.IsEnabled = true;
            if (IsLoaded) _ = RefreshAsync();
        }

        private async void ControlLoaded(object sender, RoutedEventArgs e)
        {
            if (disposed) return;
            refreshTimer.Start();
            if (refreshCancellation != null)
            {
                // A pane can be shown again before its cancelled query has finished.
                refreshOnResume = refreshCancellation.IsCancellationRequested;
                return;
            }
            await RefreshAsync();
        }

        private void ControlUnloaded(object sender, RoutedEventArgs e)
        {
            refreshTimer.Stop();
            refreshOnResume = false;
            refreshCancellation?.Cancel();
        }

        private async void TimerTick(object sender, EventArgs e)
        {
            if (IsVisible && AutoRefreshBox.IsChecked == true) await RefreshAsync();
        }

        private async void RefreshClick(object sender, RoutedEventArgs e) => await RefreshAsync();

        private async void GroupChanged(object sender, SelectionChangedEventArgs e)
        {
            if (changingSelection || disposed) return;
            selectedGroup = GroupPicker.SelectedItem as string;
            snapshot = null;
            MemberCards.ItemsSource = null;
            LocalDatabaseGrid.ItemsSource = null;
            ReplicaGrid.ItemsSource = null;
            LocalCountLabel.Text = "";
            ReplicaCountLabel.Text = "";
            LocalRoleLabel.Text = "";
            GroupScopeLabel.Text = "AG VISIBILITY";
            UpdatedLabel.Text = "Not refreshed yet";
            await RefreshAsync();
        }

        private async Task RefreshAsync()
        {
            if (disposed || refreshCancellation != null || string.IsNullOrEmpty(connectionString)) return;
            var cancellation = new CancellationTokenSource();
            refreshCancellation = cancellation;
            RefreshButton.IsEnabled = false;
            GroupPicker.IsEnabled = false;
            BusyIndicator.Visibility = Visibility.Visible;
            ErrorLabel.Visibility = Visibility.Collapsed;
            ShowUnavailable("Refreshing...", "Reading availability group state from the connected instance.");
            SnapshotNoteLabel.Text = snapshot == null ? "Loading local database health..."
                : "Refreshing. The tables show the previous snapshot from " + snapshot.CollectedAt.ToString("HH:mm:ss") + " (local time).";
            try
            {
                if (!groupsLoaded)
                {
                    var names = new List<string>(await DistributedAgHealthService.ListGroupsAsync(connectionString, cancellation.Token));
                    cancellation.Token.ThrowIfCancellationRequested();
                    if (disposed) return;
                    if (!string.IsNullOrEmpty(selectedGroup) && !names.Contains(selectedGroup)) names.Insert(0, selectedGroup);
                    changingSelection = true;
                    try
                    {
                        GroupPicker.ItemsSource = names;
                        if (string.IsNullOrEmpty(selectedGroup)) selectedGroup = names.FirstOrDefault();
                        GroupPicker.SelectedItem = selectedGroup;
                    }
                    finally { changingSelection = false; }
                    groupsLoaded = names.Count > 0;
                    if (!groupsLoaded)
                    {
                        ShowUnavailable("No availability groups", "This instance has no visible AGs. Check the server selection and metadata permissions.");
                        SnapshotNoteLabel.Text = "No availability groups are visible on this instance.";
                        LocalEmptyLabel.Text = "No availability groups found.";
                        LocalEmptyLabel.Visibility = Visibility.Visible;
                        ReplicaEmptyLabel.Visibility = Visibility.Visible;
                        return;
                    }
                }

                var result = await DistributedAgHealthService.LoadAsync(connectionString, selectedGroup, cancellation.Token);
                cancellation.Token.ThrowIfCancellationRequested();
                if (disposed) return;
                snapshot = result;
                ConnectionLabel.Text = "Connected instance: " + result.ServerName + " | SQL Server " + result.ServerVersion;
                SetSummary(LocalSummary, result.LocalStatusLevel, result.LocalStatusText, result.LocalStatusDetail);
                SetSummary(GroupSummary, result.StatusLevel, result.StatusText, result.StatusDetail);
                GroupScopeLabel.Text = result.IsDistributed ? "DISTRIBUTED AG LINK" : "AVAILABILITY GROUP";
                LocalRoleLabel.Text = "Local role: " + result.LocalRole
                    + (result.LocalGroupNames.Count == 0 ? "" : " | AG: " + string.Join(", ", result.LocalGroupNames));
                LocalDatabaseGrid.ItemsSource = result.LocalDatabases.OrderBy(row => row.DatabaseName).ThenBy(row => row.GroupName);
                ReplicaGrid.ItemsSource = result.Databases.OrderBy(row => row.DatabaseName)
                    .ThenBy(row => row.GroupName).ThenBy(row => row.IsLocal == true ? 0 : 1).ThenBy(row => row.MemberName);
                MemberCards.ItemsSource = result.Members.OrderBy(member => member.IsDistributedGroup ? 0 : 1)
                    .ThenBy(member => member.GroupName).ThenBy(member => member.IsLocal == true ? 0 : 1).ThenBy(member => member.Name);
                LocalCountLabel.Text = result.LocalDatabaseCount + " databases on this instance: "
                    + result.LocalHealthyCount + " healthy, " + result.LocalWarningCount + " partially healthy, "
                    + result.LocalCriticalCount + " need attention, " + result.LocalUnknownCount + " not visible.";
                ReplicaCountLabel.Text = result.Databases.Count + " reported database/replica rows. AG and scope identify the source of each row; remote values are reported by this instance.";
                UpdatedLabel.Text = "Refreshed " + result.CollectedAt.ToString("HH:mm:ss") + " (local time)";
                SnapshotNoteLabel.Text = result.IsDistributed
                    ? "Local database health comes from the participating AG hosted here. Distributed-link health is shown separately."
                    : "The first tab shows databases on this instance. Replica details includes every visible replica in the selected AG.";
                LocalEmptyLabel.Text = "No local database state is visible for this group. Check local AG membership and metadata permissions.";
                LocalEmptyLabel.Visibility = result.LocalDatabases.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
                ReplicaEmptyLabel.Visibility = result.Databases.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            }
            catch (OperationCanceledException)
            {
                if (!disposed) ShowRefreshFailure("Refresh cancelled", null);
            }
            catch (Exception error)
            {
                if (!disposed) ShowRefreshFailure("Unable to refresh", error.Message);
            }
            finally
            {
                refreshCancellation = null;
                cancellation.Dispose();
                if (!disposed)
                {
                    BusyIndicator.Visibility = Visibility.Collapsed;
                    RefreshButton.IsEnabled = true;
                    GroupPicker.IsEnabled = true;
                    if (refreshOnResume && IsLoaded)
                    {
                        refreshOnResume = false;
                        await RefreshAsync();
                    }
                }
            }
        }

        private void ShowRefreshFailure(string title, string error)
        {
            string detail = snapshot == null ? "No current health snapshot is available."
                : "Stale snapshot from " + snapshot.CollectedAt.ToString("yyyy-MM-dd HH:mm:ss") + " (local time).";
            ShowUnavailable(title, detail);
            SnapshotNoteLabel.Text = detail + " Refresh again for current health.";
            ErrorLabel.Text = error ?? "";
            ErrorLabel.Visibility = string.IsNullOrEmpty(error) ? Visibility.Collapsed : Visibility.Visible;
            if (snapshot == null)
            {
                LocalEmptyLabel.Text = "No health data available. Resolve the error above and refresh.";
                LocalEmptyLabel.Visibility = Visibility.Visible;
                ReplicaEmptyLabel.Visibility = Visibility.Visible;
            }
        }

        private void ShowUnavailable(string title, string detail)
        {
            SetSummary(LocalSummary, DistributedAgHealthLevel.Unknown, title, detail);
            SetSummary(GroupSummary, DistributedAgHealthLevel.Unknown, title, detail);
        }

        private static void SetSummary(Border target, DistributedAgHealthLevel level, string title, string detail)
        {
            target.DataContext = new { StatusLevel = level, StatusText = title, StatusDetail = detail };
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            refreshOnResume = false;
            refreshTimer.Stop();
            refreshTimer.Tick -= TimerTick;
            Loaded -= ControlLoaded;
            Unloaded -= ControlUnloaded;
            refreshCancellation?.Cancel();
            connectionString = null;
            themeController.Dispose();
        }
    }
}
