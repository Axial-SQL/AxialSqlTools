using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace AxialSqlTools.DataCompare
{
    public partial class ComparisonEndpointControl : UserControl, IDisposable
    {
        private readonly ToolWindowThemeController themeController;
        private List<TableSchema> tables = new List<TableSchema>();
        private CancellationTokenSource loading;
        private bool disposed, updatingSelection, tablesLoaded;
        public string ConnectionString { get; private set; }
        public TableSchema SelectedTable => TableBox.SelectedItem as TableSchema;
        public bool IsBusy => loading != null;
        public event EventHandler EndpointChanged;
        public ComparisonEndpointControl()
        {
            InitializeComponent();
            themeController = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));
        }

        private void TableSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (updatingSelection || disposed) return;
            UpdateTableStatus();
            EndpointChanged?.Invoke(this, EventArgs.Empty);
        }

        private async void ObjectExplorerClick(object sender, RoutedEventArgs e)
        {
            if (disposed || IsBusy) return;
            try
            {
                var info = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();
                if (info == null) throw new InvalidOperationException("Select a database or table in Object Explorer first.");
                await SetConnectionAsync(info.FullConnectionString);
            }
            catch (Exception ex) { if (!disposed) SetError(ex.Message); }
        }

        private async void RefreshClick(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(ConnectionString)) await SetConnectionAsync(ConnectionString);
        }

        private void TableSearchChanged(object sender, TextChangedEventArgs e)
        {
            if (TableBox == null || updatingSelection || disposed || IsBusy) return;
            ApplyTableFilter(SelectedTable, true);
        }

        private void ClearSearchClick(object sender, RoutedEventArgs e)
        {
            TableSearchBox.Clear();
            TableSearchBox.Focus();
        }

        private void ApplyTableFilter(TableSchema selectedTable, bool notifyChange)
        {
            string search = TableSearchBox.Text.Trim();
            var visibleTables = tables.Where(table => search.Length == 0 ||
                table.QualifiedName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0 ||
                (table.Schema + "." + table.Name).IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            var previousSelection = SelectedTable;
            updatingSelection = true;
            try
            {
                TableBox.ItemsSource = visibleTables;
                // Do not let the view's current item choose a table when its filter changes.
                TableBox.SelectedItem = selectedTable != null && visibleTables.Contains(selectedTable) ? selectedTable : null;
            }
            finally { updatingSelection = false; }
            UpdateControls();
            UpdateTableStatus();
            if (notifyChange && !ReferenceEquals(previousSelection, SelectedTable))
                EndpointChanged?.Invoke(this, EventArgs.Empty);
        }

        private void UpdateControls()
        {
            ObjectExplorerButton.IsEnabled = !IsBusy;
            RefreshButton.IsEnabled = !IsBusy && !string.IsNullOrEmpty(ConnectionString);
            TableSearchBox.IsEnabled = !IsBusy && tablesLoaded && tables.Count > 0;
            ClearSearchButton.IsEnabled = TableSearchBox.IsEnabled && TableSearchBox.Text.Length > 0;
            TableBox.IsEnabled = !IsBusy && tablesLoaded && TableBox.Items.Count > 0;
        }

        private void UpdateTableStatus()
        {
            if (IsBusy)
                StatusLabel.Text = "Loading tables...";
            else if (!tablesLoaded)
                StatusLabel.Text = string.IsNullOrEmpty(ConnectionString)
                    ? "Choose a database in Object Explorer, then select Use Object Explorer."
                    : "Tables could not be loaded. Refresh to retry or choose another database.";
            else if (tables.Count == 0)
                StatusLabel.Text = "No accessible user tables were found in this database.";
            else if (TableBox.Items.Count == 0)
                StatusLabel.Text = "No matching tables. Clear or change the search.";
            else
                StatusLabel.Text = TableBox.Items.Count == tables.Count
                    ? tables.Count.ToString("N0") + " tables. " + (SelectedTable == null ? "Select a table to continue." : "Table selected.")
                    : TableBox.Items.Count.ToString("N0") + " of " + tables.Count.ToString("N0") + " tables. " +
                      (SelectedTable == null ? "Select a table to continue." : "Table selected.");
        }

        private void SetError(string message)
        {
            ErrorLabel.Text = message ?? string.Empty;
            ErrorLabel.Visibility = string.IsNullOrEmpty(message) ? Visibility.Collapsed : Visibility.Visible;
        }

        public async Task SetConnectionAsync(string connectionString)
        {
            if (disposed || IsBusy) return;
            SqlConnectionStringBuilder builder;
            try
            {
                if (string.IsNullOrWhiteSpace(connectionString))
                    throw new InvalidOperationException("Select a connected database in Object Explorer first.");
                builder = new SqlConnectionStringBuilder(connectionString);
                if (string.IsNullOrWhiteSpace(builder.InitialCatalog))
                    throw new InvalidOperationException("Select a database or table in Object Explorer, rather than the server node.");
            }
            catch (Exception ex) { SetError(ex.Message); return; }

            bool sameConnection = string.Equals(ConnectionString, builder.ConnectionString, StringComparison.Ordinal);
            var previousTable = sameConnection ? SelectedTable : null;
            var cancellation = new CancellationTokenSource();
            loading = cancellation;
            try
            {
                ConnectionString = builder.ConnectionString;
                ConnectionLabel.Text = builder.DataSource + " / " + builder.InitialCatalog;
                tablesLoaded = false;
                tables.Clear();
                updatingSelection = true;
                try
                {
                    TableBox.ItemsSource = null;
                    if (!sameConnection) TableSearchBox.Clear();
                }
                finally { updatingSelection = false; }
                SetError(null);
                UpdateControls();
                UpdateTableStatus();
                EndpointChanged?.Invoke(this, EventArgs.Empty);

                var loadedTables = await Task.Run(() => SqlTableReader.ListTables(builder.ConnectionString, cancellation.Token), cancellation.Token);
                if (disposed) return;
                tables = loadedTables;
                tablesLoaded = true;
                var restoredTable = previousTable == null ? null : tables.FirstOrDefault(table =>
                    string.Equals(table.Schema, previousTable.Schema, StringComparison.Ordinal) &&
                    string.Equals(table.Name, previousTable.Name, StringComparison.Ordinal));
                ApplyTableFilter(restoredTable, false);
            }
            catch (Exception ex)
            {
                if (!disposed) SetError(cancellation.IsCancellationRequested ? "Loading tables was cancelled." : ex.Message);
            }
            finally
            {
                loading = null;
                cancellation.Dispose();
                if (!disposed)
                {
                    UpdateControls();
                    UpdateTableStatus();
                    EndpointChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            themeController.Dispose();
            loading?.Cancel();
            ConnectionString = null;
            tables.Clear();
            TableBox.ItemsSource = null;
        }
    }
}
