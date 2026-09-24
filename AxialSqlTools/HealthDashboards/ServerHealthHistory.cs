using System;
using System.Collections.Generic;
using System.Linq;

namespace AxialSqlTools
{
    // All wait types must be sampled, even if only the busiest are displayed.
    // Taking TOP N before calculating deltas makes newly ranked waits look like spikes.
    internal sealed class ServerHealthWaitHistory
    {
        internal const int BucketSeconds = 15;
        internal const int BucketCount = 15 * 60 / BucketSeconds;

        internal static DateTime BucketStart(DateTime timestamp) => new DateTime(timestamp.Year, timestamp.Month,
            timestamp.Day, timestamp.Hour, timestamp.Minute, timestamp.Second / BucketSeconds * BucketSeconds, timestamp.Kind);

        internal static DateTime WindowStart(DateTime timestamp) => BucketStart(timestamp).AddSeconds(-(BucketCount - 1) * BucketSeconds);

        private Dictionary<string, decimal> previous;
        private DateTime previousTimestamp;
        private DateTime serverStart;
        private readonly SortedDictionary<DateTime, Dictionary<string, decimal>> buckets =
            new SortedDictionary<DateTime, Dictionary<string, decimal>>();

        internal void Update(Dictionary<string, decimal> totals, DateTime timestamp, DateTime startTime)
        {
            DateTime bucketStart = BucketStart(timestamp);
            DateTime windowStart = WindowStart(timestamp);
            foreach (DateTime key in buckets.Keys.Where(x => x < windowStart).ToArray())
                buckets.Remove(key);

            // Baseline after startup, restart, clock reversal or a long sampling gap.
            // A decreased counter indicates DBCC SQLPERF(..., CLEAR) or another reset.
            bool reset = previous == null || serverStart != startTime || timestamp <= previousTimestamp
                || timestamp - previousTimestamp > TimeSpan.FromSeconds(BucketSeconds)
                || previous.Any(x => (totals.TryGetValue(x.Key, out decimal value) ? value : 0) < x.Value);
            if (reset)
            {
                if (previous != null && (serverStart != startTime || timestamp <= previousTimestamp)) buckets.Clear();
            }
            else
            {
                if (!buckets.TryGetValue(bucketStart, out Dictionary<string, decimal> bucket))
                    buckets[bucketStart] = bucket = new Dictionary<string, decimal>(StringComparer.Ordinal);
                foreach (var current in totals)
                {
                    decimal delta = current.Value - (previous.TryGetValue(current.Key, out decimal old) ? old : 0);
                    bucket[current.Key] = (bucket.TryGetValue(current.Key, out decimal accumulated) ? accumulated : 0) + Math.Max(0, delta);
                }
            }

            previous = new Dictionary<string, decimal>(totals, StringComparer.Ordinal);
            previousTimestamp = timestamp;
            serverStart = startTime;
        }

        internal KeyValuePair<DateTime, Dictionary<string, decimal>>[] Snapshot() => buckets.ToArray();
    }

    internal static class ServerHealthCounterMath
    {
        internal static double Rate(long current, long previous, double seconds, bool validBaseline)
        {
            return validBaseline && seconds > 0 && current >= previous
                ? (current - previous) / seconds : double.NaN;
        }
    }
}
