using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace AxialSqlTools
{
    internal partial class ScriptObjectPickerDialog : Window
    {
        internal ScriptObjectSelectionItem SelectedObject { get; private set; }

        internal ScriptObjectPickerDialog(IEnumerable<ScriptObjectSelectionItem> matches)
        {
            InitializeComponent();

            HeaderTextBlock.Text = "Select the object to script.";
            ObjectsListBox.ItemsSource = matches.ToList();
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            if (ObjectsListBox.SelectedItem is ScriptObjectSelectionItem selected)
            {
                SelectedObject = selected;
                DialogResult = true;
            }
            else
            {
                MessageBox.Show("Select an object to script.", "Script Object");
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
