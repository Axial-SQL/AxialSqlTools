using System;

namespace AxialSqlTools
{
    internal sealed class StatisticsSummaryStoreState
    {
        public StatisticsSummary Summary { get; set; }
        public bool IsLoading { get; set; }
        public bool LastCaptureFailed { get; set; }
    }

    internal static class StatisticsSummaryStore
    {
        private static readonly object SyncRoot = new object();
        private static StatisticsSummary _current;
        private static bool _isLoading;
        private static bool _lastCaptureFailed;
        private static bool _isWindowOpen;
        private static int _activeCaptureVersion;

        public static event EventHandler SummaryChanged;

        public static StatisticsSummary GetCurrent()
        {
            lock (SyncRoot)
            {
                return _current;
            }
        }

        public static StatisticsSummaryStoreState GetState()
        {
            lock (SyncRoot)
            {
                return new StatisticsSummaryStoreState
                {
                    Summary = _current,
                    IsLoading = _isLoading,
                    LastCaptureFailed = _lastCaptureFailed,
                };
            }
        }

        public static bool IsWindowOpen()
        {
            lock (SyncRoot)
            {
                return _isWindowOpen;
            }
        }

        public static void SetWindowOpen(bool isWindowOpen)
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
                }
            }

            if (shouldRaise)
            {
                SummaryChanged?.Invoke(null, EventArgs.Empty);
            }
        }

        public static bool BeginCapture(int captureVersion)
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
            }

            SummaryChanged?.Invoke(null, EventArgs.Empty);
            return true;
        }

        public static void Set(StatisticsSummary summary, int captureVersion)
        {
            lock (SyncRoot)
            {
                if (captureVersion != _activeCaptureVersion)
                {
                    return;
                }

                _current = summary;
                _isLoading = false;
                _lastCaptureFailed = false;
            }

            SummaryChanged?.Invoke(null, EventArgs.Empty);
        }

        public static void MarkUnavailable(int captureVersion)
        {
            lock (SyncRoot)
            {
                if (captureVersion != _activeCaptureVersion)
                {
                    return;
                }

                _isLoading = false;
                _lastCaptureFailed = true;
            }

            SummaryChanged?.Invoke(null, EventArgs.Empty);
        }

        public static void CancelCapture()
        {
            lock (SyncRoot)
            {
                _isLoading = false;
                _lastCaptureFailed = false;
                _activeCaptureVersion++;
            }

            SummaryChanged?.Invoke(null, EventArgs.Empty);
        }

        public static void Clear()
        {
            lock (SyncRoot)
            {
                _current = null;
                _isLoading = false;
                _lastCaptureFailed = false;
                _isWindowOpen = false;
                _activeCaptureVersion = 0;
            }

            SummaryChanged?.Invoke(null, EventArgs.Empty);
        }
    }
}