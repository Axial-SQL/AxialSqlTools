using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace AxialSqlTools
{
    public partial class SavedConnectionManagerWindow : Window
    {
        private readonly ObservableCollection<SettingsManager.DataTransferSavedConnection> _connections;

        public SavedConnectionManagerWindow()
        {
            InitializeComponent();

            ProviderComboBox.ItemsSource = Enum.GetValues(typeof(SettingsManager.DataTransferProvider));

            _connections = new ObservableCollection<SettingsManager.DataTransferSavedConnection>(
                SettingsManager.GetDataTransferSavedConnections());

            ConnectionsListBox.ItemsSource = _connections;
        }

        private void AddNewButton_Click(object sender, RoutedEventArgs e)
        {
            var newConnection = new SettingsManager.DataTransferSavedConnection
            {
                Name = "New Connection",
                Provider = SettingsManager.DataTransferProvider.PostgreSql,
                Server = "127.0.0.1",
                Port = 5432,
                Database = "postgres",
                Username = "postgres",
                Password = ""
            };

            _connections.Add(newConnection);
            ConnectionsListBox.SelectedItem = newConnection;
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (ConnectionsListBox.SelectedItem is SettingsManager.DataTransferSavedConnection selected)
            {
                _connections.Remove(selected);
                SaveConnections();
                ClearEditor();
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (ConnectionsListBox.SelectedItem is SettingsManager.DataTransferSavedConnection selected)
            {
                ApplyEditorValues(selected);
                RefreshListItem(selected);
                SaveConnections();
            }
            else
            {
                MessageBox.Show("Select a connection to save.", "Saved Connections");
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ConnectionsListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ConnectionsListBox.SelectedItem is SettingsManager.DataTransferSavedConnection selected)
            {
                DetailsPanel.IsEnabled = true;
                NameTextBox.Text = selected.Name;
                ProviderComboBox.SelectedItem = selected.Provider;
                ServerTextBox.Text = selected.Server;
                PortTextBox.Text = selected.Port.ToString();
                DatabaseTextBox.Text = selected.Database;
                UsernameTextBox.Text = selected.Username;
                PasswordBox.Password = selected.Password ?? string.Empty;
            }
            else
            {
                DetailsPanel.IsEnabled = false;
                ClearEditor();
            }
        }

        private void ApplyEditorValues(SettingsManager.DataTransferSavedConnection selected)
        {
            selected.Name = NameTextBox.Text?.Trim();
            if (ProviderComboBox.SelectedItem is SettingsManager.DataTransferProvider provider)
            {
                selected.Provider = provider;
            }

            selected.Server = ServerTextBox.Text?.Trim();
            if (int.TryParse(PortTextBox.Text, out int port))
            {
                selected.Port = port;
            }

            selected.Database = DatabaseTextBox.Text?.Trim();
            selected.Username = UsernameTextBox.Text?.Trim();
            selected.Password = PasswordBox.Password;
        }

        private void SaveConnections()
        {
            SettingsManager.SaveDataTransferSavedConnections(_connections.ToList());
        }

        private void ClearEditor()
        {
            NameTextBox.Text = string.Empty;
            ProviderComboBox.SelectedItem = null;
            ServerTextBox.Text = string.Empty;
            PortTextBox.Text = string.Empty;
            DatabaseTextBox.Text = string.Empty;
            UsernameTextBox.Text = string.Empty;
            PasswordBox.Password = string.Empty;
        }

        private void RefreshListItem(SettingsManager.DataTransferSavedConnection selected)
        {
            var index = _connections.IndexOf(selected);
            if (index >= 0)
            {
                _connections[index] = selected;
                ConnectionsListBox.SelectedItem = selected;
            }
        }
    }
}
