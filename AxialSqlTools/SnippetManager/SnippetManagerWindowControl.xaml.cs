using System;
using System.Windows;
using System.Windows.Controls;

namespace AxialSqlTools
{
    public partial class SnippetManagerWindowControl : UserControl
    {
        public SnippetManagerWindowControl()
        {
            InitializeComponent();
            var vm = new SnippetManagerViewModel();
            DataContext = vm;

            // Set the ComboBox to the current ReplaceKey value
            cmbReplaceKey.SelectedValue = vm.ReplaceKey.ToString();
        }

        private void cmbReplaceKey_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DataContext is SnippetManagerViewModel vm && cmbReplaceKey.SelectedValue is string selectedValue)
            {
                if (Enum.TryParse(selectedValue, out SettingsManager.SnippetReplaceKey key))
                {
                    vm.ReplaceKey = key;
                }
            }
        }
    }
}
