using System;
using System.Collections.Generic;
using System.Linq;

namespace AxialSqlTools
{
    internal enum StatisticsSummaryCaptureStatus
    {
        None,
        Loading,
        Success,
        Failed,
        StatisticsDisabled,
    }

    internal sealed class StatisticsSummaryStoreResult
    {
        public StatisticsSummary Summary { get; set; }
        public StatisticsSummaryCaptureStatus Status { get; set; }
    }

    internal sealed class StatisticsSummaryStoreState
    {
        internal StatisticsSummary Summary { get; set; }
        internal IReadOnlyList<StatisticsSummary> Summaries { get; set; }
        internal IReadOnlyList<StatisticsSummaryStoreResult> Results { get; set; }
        internal bool IsLoading { get; set; }
        internal bool LastCaptureFailed { get; set; }
    }

    internal static class StatisticsSummaryStore
    {
        private static readonly object SyncRoot = new object();
        private static readonly List<StatisticsSummaryStoreResult> Results = new List<StatisticsSummaryStoreResult>();
        private static bool _isLoading;
        private static bool _lastCaptureFailed;
        private static bool _isWindowOpen;
        private static int _activeCaptureVersion;

        internal static event EventHandler SummaryChanged;

        internal static StatisticsSummary GetCurrent()
        {
            lock (SyncRoot)
            {
                return Results.FirstOrDefault(result => result.Status == StatisticsSummaryCaptureStatus.Success)?.Summary;
            }
        }

        internal static StatisticsSummaryStoreState GetState()
        {
            lock (SyncRoot)
            {
                var results = Results.Select(CloneResult).ToList();
                var summaries = results
                    .Where(result => result.Status == StatisticsSummaryCaptureStatus.Success && result.Summary != null)
                    .Select(result => result.Summary)
                    .ToList();

                return new StatisticsSummaryStoreState
                {
                    Summary = summaries.FirstOrDefault(),
                    Summaries = summaries,
                    Results = results,
                    IsLoading = _isLoading,
                    LastCaptureFailed = _lastCaptureFailed,
                };
            }
        }

        internal static bool IsWindowOpen()
        {
            lock (SyncRoot)
            {
                return _isWindowOpen;
            }
        }

        internal static void SetWindowOpen(bool isWindowOpen)
        {
            var shouldRaise = false;

            lock (SyncRoot)
            {
                if (_isWindowOpen == isWindowOpen)
                {
                    return;
                }

                _isWindowOpen = isWindowOpen;
                shouldRaise = true;

                if (!isWindowOpen)
                {
                    _isLoading = false;
                    _lastCaptureFailed = false;
                    _activeCaptureVersion++;
                    RemoveLoadingResult();
                }
            }

            if (shouldRaise)
            {
                SummaryChanged?.Invoke(null, EventArgs.Empty);
            }
        }

        internal static bool BeginCapture(int captureVersion)
        {
            lock (SyncRoot)
            {
                if (!_isWindowOpen || captureVersion < _activeCaptureVersion)
                {
                    return false;
                }

                _activeCaptureVersion = captureVersion;
                _isLoading = true;
                _lastCaptureFailed = false;

                if (Results.Count == 0 || Results[0].Status != StatisticsSummaryCaptureStatus.Loading)
                {
                    Results.Insert(0, new StatisticsSummaryStoreResult
                    {
                        Status = StatisticsSummaryCaptureStatus.Loading,
                    });
                    TrimResults();
                }
                else
                {
                    Results[0].Summary = null;
                    Results[0].Status = StatisticsSummaryCaptureStatus.Loading;
                }
            }

            SummaryChanged?.Invoke(null, EventArgs.Empty);
            return true;
        }

        internal static void Set(StatisticsSummary summary, int captureVersion)
        {
            lock (SyncRoot)
            {
                if (captureVersion != _activeCaptureVersion)
                {
                    return;
                }

                SetActiveResult(summary, StatisticsSummaryCaptureStatus.Success);

                _isLoading = false;
                _lastCaptureFailed = false;
            }

            SummaryChanged?.Invoke(null, EventArgs.Empty);
        }

        internal static void MarkUnavailable(int captureVersion)
        {
            MarkUnavailable(captureVersion, StatisticsSummaryCaptureStatus.Failed);
        }

        internal static void MarkUnavailable(int captureVersion, StatisticsSummaryCaptureStatus status)
        {
            lock (SyncRoot)
            {
                if (captureVersion != _activeCaptureVersion)
                {
                    return;
                }

                SetActiveResult(null, status);
                _isLoading = false;
                _lastCaptureFailed = true;
            }

            SummaryChanged?.Invoke(null, EventArgs.Empty);
        }

        internal static void CancelCapture()
        {
            lock (SyncRoot)
            {
                _isLoading = false;
                _lastCaptureFailed = false;
                _activeCaptureVersion++;
                RemoveLoadingResult();
            }

            SummaryChanged?.Invoke(null, EventArgs.Empty);
        }

        internal static void Clear()
        {
            lock (SyncRoot)
            {
                Results.Clear();
                _isLoading = false;
                _lastCaptureFailed = false;
                _isWindowOpen = false;
                _activeCaptureVersion = 0;
            }

            SummaryChanged?.Invoke(null, EventArgs.Empty);
        }

        private static void SetActiveResult(StatisticsSummary summary, StatisticsSummaryCaptureStatus status)
        {
            if (Results.Count == 0)
            {
                Results.Insert(0, new StatisticsSummaryStoreResult());
            }

            Results[0].Summary = summary;
            Results[0].Status = status;
            TrimResults();
        }

        private static void RemoveLoadingResult()
        {
            if (Results.Count > 0 && Results[0].Status == StatisticsSummaryCaptureStatus.Loading)
            {
                Results.RemoveAt(0);
            }
        }

        private static void TrimResults()
        {
            while (Results.Count > 2)
            {
                Results.RemoveAt(Results.Count - 1);
            }
        }

        private static StatisticsSummaryStoreResult CloneResult(StatisticsSummaryStoreResult result)
        {
            return new StatisticsSummaryStoreResult
            {
                Summary = result.Summary,
                Status = result.Status,
            };
        }
    }
}
