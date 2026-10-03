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
        }

        public void Initialize(ScriptFactoryAccess.ConnectionInfo connection, string availabilityGroupName)
        {
            if (connection == null || string.IsNullOrWhiteSpace(connection.FullConnectionString))
                throw new ArgumentException("Select a connected SQL Server instance in Object Explorer.", nameof(connection));
            connectionString = connection.FullConnectionString;
            selectedGroup = availabilityGroupName;
            ConnectionLabel.Text = "Observed from " + connection.ServerName;
            RefreshButton.IsEnabled = true;
            if (IsLoaded) _ = RefreshAsync();
        }

        private async void ControlLoaded(object sender, RoutedEventArgs e)
        {
            if (disposed) return;
            refreshTimer.Start();
            await RefreshAsync();
        }

        private void ControlUnloaded(object sender, RoutedEventArgs e)
        {
            refreshTimer.Stop();
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
            DatabaseGrid.ItemsSource = null;
            DatabaseCountLabel.Text = "";
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
            StatusLabel.DataContext = null;
            StatusLabel.Text = "Refreshing...";
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
                        StatusLabel.Text = "No distributed availability groups visible";
                        StatusDetailLabel.Text = "This instance has no visible distributed AGs. Check the server selection and metadata permissions.";
                        EmptyLabel.Text = "No distributed availability groups are visible on this instance.";
                        EmptyLabel.Visibility = Visibility.Visible;
                        return;
                    }
                }

                var result = await DistributedAgHealthService.LoadAsync(connectionString, selectedGroup, cancellation.Token);
                cancellation.Token.ThrowIfCancellationRequested();
                if (disposed) return;
                snapshot = result;
                ConnectionLabel.Text = "Observed from " + result.ServerName + " | SQL Server " + result.ServerVersion;
                StatusLabel.DataContext = result;
                StatusLabel.Text = result.StatusText;
                StatusDetailLabel.Text = result.StatusDetail;
                MemberCards.ItemsSource = result.Members.OrderBy(member => member.RoleCode == 1 ? 0 : 1).ThenBy(member => member.Name);
                DatabaseGrid.ItemsSource = result.Databases;
                DatabaseCountLabel.Text = result.Databases.Select(row => row.DatabaseName).Distinct().Count()
                    + " visible databases / " + result.Databases.Count + " reported rows";
                UpdatedLabel.Text = "Refreshed " + result.CollectedAt.ToString("HH:mm:ss") + " (local time)";
                EmptyLabel.Text = "No database replica state is visible from this instance. Connect to the global primary or forwarder and check permissions.";
                EmptyLabel.Visibility = result.Databases.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            }
            catch (OperationCanceledException)
            {
                if (!disposed)
                {
                    StatusLabel.DataContext = null;
                    StatusLabel.Text = "Refresh cancelled";
                    StatusDetailLabel.Text = snapshot == null ? "Refresh again to obtain current health."
                        : "Showing stale data from " + snapshot.CollectedAt.ToString("yyyy-MM-dd HH:mm:ss") + " (local time). Refresh again for current health.";
                }
            }
            catch (Exception error)
            {
                if (!disposed)
                {
                    StatusLabel.DataContext = null;
                    StatusLabel.Text = "Unable to refresh";
                    StatusDetailLabel.Text = snapshot == null ? "No current health snapshot is available."
                        : "Showing stale data from " + snapshot.CollectedAt.ToString("yyyy-MM-dd HH:mm:ss") + " (local time).";
                    ErrorLabel.Text = error.Message;
                    ErrorLabel.Visibility = Visibility.Visible;
                    if (snapshot == null)
                    {
                        EmptyLabel.Text = "No health data available. Resolve the error above and refresh.";
                        EmptyLabel.Visibility = Visibility.Visible;
                    }
                }
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
                }
            }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
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
