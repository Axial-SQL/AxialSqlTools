using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using Microsoft.VisualStudio.Shell;

namespace AxialSqlTools
{
    public partial class StatisticsSummaryWindowControl : UserControl
    {
        private readonly ToolWindowThemeController _themeController;
        private readonly StatisticsSummaryViewModel _viewModel;
        private bool _isSummarySubscribed;

        public StatisticsSummaryWindowControl()
        {
            InitializeComponent();
            _themeController = new ToolWindowThemeController(this, ApplyThemeBrushResources);
            _viewModel = new StatisticsSummaryViewModel();
            DataContext = _viewModel;
            EnsureSummarySubscription();

            Loaded += OnLoaded;
            IsVisibleChanged += OnIsVisibleChanged;
        }

        private void ApplyThemeBrushResources()
        {
            ToolWindowThemeResources.ApplySharedTheme(this);
        }

        private void OnLoaded(object sender, System.Windows.RoutedEventArgs e)
        {
            EnsureSummarySubscription();
            StatisticsSummaryStore.SetWindowOpen(true);
            _viewModel.RefreshFromStore();
        }

        private void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (!IsVisible)
            {
                return;
            }

            StatisticsSummaryStore.SetWindowOpen(true);
            EnsureSummarySubscription();
            _viewModel.RefreshFromStore();
        }

        private void EnsureSummarySubscription()
        {
            if (_isSummarySubscribed)
            {
                return;
            }

            StatisticsSummaryStore.SummaryChanged += StatisticsSummaryStore_SummaryChanged;
            _isSummarySubscribed = true;
        }

        private void StatisticsSummaryStore_SummaryChanged(object sender, EventArgs e)
        {
            if (Dispatcher.CheckAccess())
            {
                _viewModel.RefreshFromStore();
                return;
            }

            ThreadHelper.JoinableTaskFactory.Run(async delegate
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                _viewModel.RefreshFromStore();
            });
        }

        private void WikiLink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            ToolWindowNavigation.HandleRequestNavigate(e);
        }

        private void CancelLoading_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            AxialSqlToolsPackage.CancelStatisticsCapture();
        }
    }
}