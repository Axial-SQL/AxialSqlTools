using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace AxialSqlTools
{
    public class StatisticsSummaryResultViewModel
    {
        public string Title { get; set; }
        public bool ShowDivider { get; set; }
        public ObservableCollection<StatisticsTableSummary> Tables { get; } = new ObservableCollection<StatisticsTableSummary>();
        public string CapturedAtText { get; set; }
        public string SourceText { get; set; }
        public string TotalReadsText { get; set; }
        public string TotalElapsedText { get; set; }
        public string TotalCpuText { get; set; }
        public string QueryTextPreview { get; set; }
        public bool HasTableData => Tables.Count > 0;

        public static StatisticsSummaryResultViewModel Create(string title, StatisticsSummary summary, bool showDivider)
        {
            var result = new StatisticsSummaryResultViewModel
            {
                Title = title,
                ShowDivider = showDivider,
                CapturedAtText = "-",
                SourceText = "-",
                TotalReadsText = "0",
                TotalElapsedText = "-",
                TotalCpuText = "-",
                QueryTextPreview = string.Empty,
            };

            if (summary == null || !summary.HasData)
            {
                return result;
            }

            foreach (var table in summary.Tables.OrderByDescending(table => table.TotalReads))
            {
                result.Tables.Add(table);
            }

            result.CapturedAtText = summary.CapturedAt.ToString("yyyy-MM-dd HH:mm:ss");
            result.SourceText = BuildSourceText(summary);
            result.TotalReadsText = summary.TotalReads.ToString("N0");
            result.TotalElapsedText = FormatMilliseconds(summary.TotalElapsedMilliseconds);
            result.TotalCpuText = FormatMilliseconds(summary.TotalCpuMilliseconds);
            result.QueryTextPreview = summary.QueryText ?? string.Empty;
            return result;
        }

        private static string BuildSourceText(StatisticsSummary summary)
        {
            if (string.IsNullOrWhiteSpace(summary.DataSource) && string.IsNullOrWhiteSpace(summary.DatabaseName))
            {
                return "-";
            }

            if (string.IsNullOrWhiteSpace(summary.DatabaseName))
            {
                return summary.DataSource;
            }

            return summary.DataSource + " / " + summary.DatabaseName;
        }

        private static string FormatMilliseconds(long? milliseconds)
        {
            if (!milliseconds.HasValue)
            {
                return "-";
            }

            if (milliseconds.Value >= 1000)
            {
                return (milliseconds.Value / 1000d).ToString("0.00") + " s";
            }

            return milliseconds.Value.ToString("N0") + " ms";
        }
    }

    public class StatisticsSummaryViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<StatisticsSummaryResultViewModel> Results { get; } = new ObservableCollection<StatisticsSummaryResultViewModel>();

        private bool _isLoading;
        public bool IsLoading
        {
            get => _isLoading;
            set
            {
                if (_isLoading != value)
                {
                    _isLoading = value;
                    OnPropertyChanged(nameof(IsLoading));
                }
            }
        }

        private string _statusMessage;
        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                if (_statusMessage != value)
                {
                    _statusMessage = value;
                    OnPropertyChanged(nameof(StatusMessage));
                }
            }
        }

        private bool _showStatus;
        public bool ShowStatus
        {
            get => _showStatus;
            set
            {
                if (_showStatus != value)
                {
                    _showStatus = value;
                    OnPropertyChanged(nameof(ShowStatus));
                }
            }
        }

        public StatisticsSummaryViewModel()
        {
            RefreshFromStore();
        }

        public void RefreshFromStore()
        {
            var state = StatisticsSummaryStore.GetState();
            IsLoading = state.IsLoading;
            ShowStatus = state.IsLoading;
            StatusMessage = state.IsLoading ? "Loading latest statistics output..." : string.Empty;

            var summaries = state.Summaries ?? Array.Empty<StatisticsSummary>();
            Results.Clear();
            Results.Add(StatisticsSummaryResultViewModel.Create("Latest result", summaries.ElementAtOrDefault(0), showDivider: false));
            Results.Add(StatisticsSummaryResultViewModel.Create("Previous result", summaries.ElementAtOrDefault(1), showDivider: true));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
