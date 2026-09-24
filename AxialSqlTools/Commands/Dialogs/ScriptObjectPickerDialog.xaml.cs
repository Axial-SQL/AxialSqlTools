using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace AxialSqlTools
{
    public partial class ScriptObjectPickerDialog : Microsoft.VisualStudio.PlatformUI.DialogWindow
    {
        private readonly ToolWindowThemeController themeController;
        public ScriptObjectSelectionItem SelectedObject { get; set; }

        public ScriptObjectPickerDialog(IEnumerable<ScriptObjectSelectionItem> matches, string action = "script")
        {
            InitializeComponent();
            themeController = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));

            HeaderTextBlock.Text = "Select the object to " + action + ".";
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
                MessageBox.Show("Select an object from the list.", "Select Object");
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
