using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace AxialSqlTools
{
    public partial class SavedConnectionPickerWindow : Window
    {
        public SettingsManager.DataTransferSavedConnection SelectedConnection { get; private set; }

        public SavedConnectionPickerWindow(IEnumerable<SettingsManager.DataTransferSavedConnection> connections, string title)
        {
            InitializeComponent();

            HeaderTextBlock.Text = title;
            ConnectionsListBox.ItemsSource = connections.ToList();
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            if (ConnectionsListBox.SelectedItem is SettingsManager.DataTransferSavedConnection selected)
            {
                SelectedConnection = selected;
                DialogResult = true;
            }
            else
            {
                MessageBox.Show("Select a saved connection.", "Saved Connections");
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
