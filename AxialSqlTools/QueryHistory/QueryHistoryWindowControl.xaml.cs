using System;
using System.Windows.Controls;
using System.Windows.Input;

namespace AxialSqlTools
{
    public partial class QueryHistoryWindowControl : UserControl
    {
        public QueryHistoryWindowControl()
        {
            InitializeComponent();
            DataContext = new QueryHistoryViewModel();
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (DataContext is QueryHistoryViewModel vm && vm.RefreshCommand.CanExecute(null))
                {
                    vm.RefreshCommand.Execute(null);
                }
                e.Handled = true;
            }
        }

        private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            // force a refresh right away:
            if (DataContext is QueryHistoryViewModel vm)
            {
                vm.RefreshCommand.Execute(null);
            }
        }

    }
}
