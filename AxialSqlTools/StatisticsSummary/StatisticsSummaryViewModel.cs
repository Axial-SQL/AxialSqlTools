using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace AxialSqlTools
{
    public class StatisticsSummaryViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<StatisticsTableSummary> Tables { get; } = new ObservableCollection<StatisticsTableSummary>();

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

        private string _capturedAtText;
        public string CapturedAtText
        {
            get => _capturedAtText;
            set
            {
                if (_capturedAtText != value)
                {
                    _capturedAtText = value;
                    OnPropertyChanged(nameof(CapturedAtText));
                }
            }
        }

        private string _sourceText;
        public string SourceText
        {
            get => _sourceText;
            set
            {
                if (_sourceText != value)
                {
                    _sourceText = value;
                    OnPropertyChanged(nameof(SourceText));
                }
            }
        }

        private string _totalReadsText;
        public string TotalReadsText
        {
            get => _totalReadsText;
            set
            {
                if (_totalReadsText != value)
                {
                    _totalReadsText = value;
                    OnPropertyChanged(nameof(TotalReadsText));
                }
            }
        }

        private string _totalElapsedText;
        public string TotalElapsedText
        {
            get => _totalElapsedText;
            set
            {
                if (_totalElapsedText != value)
                {
                    _totalElapsedText = value;
                    OnPropertyChanged(nameof(TotalElapsedText));
                }
            }
        }

        private string _totalCpuText;
        public string TotalCpuText
        {
            get => _totalCpuText;
            set
            {
                if (_totalCpuText != value)
                {
                    _totalCpuText = value;
                    OnPropertyChanged(nameof(TotalCpuText));
                }
            }
        }

        private string _queryTextPreview;
        public string QueryTextPreview
        {
            get => _queryTextPreview;
            set
            {
                if (_queryTextPreview != value)
                {
                    _queryTextPreview = value;
                    OnPropertyChanged(nameof(QueryTextPreview));
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
            var summary = state.Summary;
            IsLoading = state.IsLoading;

            Tables.Clear();
            if (summary == null || !summary.HasData)
            {
                if (state.IsLoading)
                {
                    StatusMessage = "Loading latest statistics output...";
                }
                else if (state.LastCaptureFailed)
                {
                    StatusMessage = "The last execution did not expose SET STATISTICS IO/TIME output before the retry window expired.";
                }
                else
                {
                    StatusMessage = "No SET STATISTICS IO/TIME output was captured for the last execution.";
                }

                CapturedAtText = "-";
                SourceText = "-";
                TotalReadsText = "0";
                TotalElapsedText = "-";
                TotalCpuText = "-";
                QueryTextPreview = string.Empty;
                return;
            }

            foreach (var table in summary.Tables.OrderByDescending(table => table.TotalReads))
            {
                Tables.Add(table);
            }

            if (state.IsLoading)
            {
                StatusMessage = "Loading latest statistics output...";
            }
            else if (state.LastCaptureFailed)
            {
                StatusMessage = $"The latest execution did not expose statistics output before the retry window expired. Showing the most recent captured summary from {summary.CapturedAt:yyyy-MM-dd HH:mm:ss}.";
            }
            else
            {
                StatusMessage = string.Empty;
            }

            CapturedAtText = summary.CapturedAt.ToString("yyyy-MM-dd HH:mm:ss");
            SourceText = BuildSourceText(summary);
            TotalReadsText = summary.TotalReads.ToString("N0");
            TotalElapsedText = FormatMilliseconds(summary.TotalElapsedMilliseconds);
            TotalCpuText = FormatMilliseconds(summary.TotalCpuMilliseconds);
            QueryTextPreview = summary.QueryText ?? string.Empty;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
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
}