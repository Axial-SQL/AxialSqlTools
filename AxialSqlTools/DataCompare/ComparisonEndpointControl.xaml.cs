using Microsoft.Data.SqlClient;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace AxialSqlTools.DataCompare
{
    public partial class ComparisonEndpointControl : UserControl, IDisposable
    {
        private CancellationTokenSource loading;
        private bool disposed;
        public string ConnectionString { get; private set; }
        public TableSchema SelectedTable => TableBox.SelectedItem as TableSchema;
        public bool IsBusy => loading != null;
        public event EventHandler EndpointChanged;
        public ComparisonEndpointControl() { InitializeComponent(); }

        private void TableSelectionChanged(object sender, SelectionChangedEventArgs e) => EndpointChanged?.Invoke(this, EventArgs.Empty);
        private async void ObjectExplorerClick(object sender, RoutedEventArgs e) => await UseSsmsConnection(false);
        private async void ActiveQueryClick(object sender, RoutedEventArgs e) => await UseSsmsConnection(true);
        private async Task UseSsmsConnection(bool query)
        {
            try
            {
                var info = query ? ScriptFactoryAccess.GetCurrentConnectionInfo() : ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();
                if (info == null) throw new InvalidOperationException(query ? "Activate a connected SQL query first." : "Select a database in Object Explorer first.");
                await SetConnectionAsync(info.FullConnectionString);
            }
            catch (Exception ex) { ErrorLabel.Text = ex.Message; }
        }
        private async void NewConnectionClick(object sender, RoutedEventArgs e)
        {
            var dialog = new SqlConnectionDialog(ConnectionString) { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() == true) await SetConnectionAsync(dialog.ConnectionString);
        }
        private async void RefreshClick(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(ConnectionString)) await SetConnectionAsync(ConnectionString);
        }

        public async Task SetConnectionAsync(string connectionString)
        {
            if (disposed || IsBusy) return;
            var cancellation = new CancellationTokenSource(); loading = cancellation;
            ConnectionString = connectionString;
            ConnectionButtons.IsEnabled = false; TableBox.IsEnabled = false; TableBox.ItemsSource = null;
            var builder = new SqlConnectionStringBuilder(connectionString);
            ConnectionLabel.Text = builder.DataSource + " / " + builder.InitialCatalog;
            ErrorLabel.Text = "Loading tables...";
            EndpointChanged?.Invoke(this, EventArgs.Empty);
            try
            {
                var tables = await Task.Run(() => SqlTableReader.ListTables(connectionString, cancellation.Token), cancellation.Token);
                if (disposed) return;
                TableBox.ItemsSource = tables;
                ErrorLabel.Text = tables.Count == 0 ? "No accessible user tables were found." : "";
            }
            catch (Exception ex) { if (!disposed) ErrorLabel.Text = cancellation.IsCancellationRequested ? "Cancelled." : ex.Message; }
            finally
            {
                loading = null; cancellation.Dispose();
                if (!disposed)
                {
                    ConnectionButtons.IsEnabled = true; TableBox.IsEnabled = true;
                    EndpointChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        public void Dispose() { disposed = true; loading?.Cancel(); ConnectionString = null; }
    }
}
