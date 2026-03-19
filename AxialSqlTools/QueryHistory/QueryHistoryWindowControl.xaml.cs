using System;
using System.Diagnostics;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Navigation;

namespace AxialSqlTools
{
    public partial class QueryHistoryWindowControl : UserControl
    {
        private readonly ToolWindowThemeController _themeController;

        public QueryHistoryWindowControl()
        {
            InitializeComponent();
            _themeController = new ToolWindowThemeController(this, ApplyThemeBrushResources);
            DataContext = new QueryHistoryViewModel();
        }

        private void ApplyThemeBrushResources()
        {
            ToolWindowThemeResources.ApplySharedTheme(this);
        }

        private void WikiLink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(e.Uri?.AbsoluteUri))
            {
                return;
            }

            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri)
            {
                UseShellExecute = true,
            });

            e.Handled = true;
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
