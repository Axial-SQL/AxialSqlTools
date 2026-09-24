using System;
using System.Collections.Generic;
using System.Linq;

namespace AxialSqlTools
{
    // All wait types must be sampled, even if only the busiest are displayed.
    // Taking TOP N before calculating deltas makes newly ranked waits look like spikes.
    internal sealed class ServerHealthWaitHistory
    {
        private Dictionary<string, decimal> previous;
        private DateTime previousTimestamp;
        private DateTime serverStart;
        private readonly SortedDictionary<DateTime, Dictionary<string, decimal>> minutes =
            new SortedDictionary<DateTime, Dictionary<string, decimal>>();

        internal void Update(Dictionary<string, decimal> totals, DateTime timestamp, DateTime startTime)
        {
            DateTime minute = new DateTime(timestamp.Year, timestamp.Month, timestamp.Day, timestamp.Hour, timestamp.Minute, 0);
            foreach (DateTime key in minutes.Keys.Where(x => x < minute.AddMinutes(-14)).ToArray())
                minutes.Remove(key);

            // Baseline after startup, restart, clock reversal or a long sampling gap.
            // A decreased counter indicates DBCC SQLPERF(..., CLEAR) or another reset.
            bool reset = previous == null || serverStart != startTime || timestamp <= previousTimestamp
                || timestamp - previousTimestamp > TimeSpan.FromMinutes(1)
                || previous.Any(x => (totals.TryGetValue(x.Key, out decimal value) ? value : 0) < x.Value);
            if (reset)
            {
                if (previous != null && (serverStart != startTime || timestamp <= previousTimestamp)) minutes.Clear();
            }
            else
            {
                if (!minutes.TryGetValue(minute, out Dictionary<string, decimal> bucket))
                    minutes[minute] = bucket = new Dictionary<string, decimal>(StringComparer.Ordinal);
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

        internal KeyValuePair<DateTime, Dictionary<string, decimal>>[] Snapshot() => minutes.ToArray();
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
