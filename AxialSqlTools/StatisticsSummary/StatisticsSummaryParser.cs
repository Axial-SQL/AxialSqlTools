using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace AxialSqlTools
{
    internal static class StatisticsSummaryParser
    {
        private static readonly Regex TableLineRegex = new Regex(
            @"Table\s+'(?<name>[^']+)'\.\s*(?<metrics>.+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ScanCountRegex = new Regex(
            @"Scan count\s+(?<value>\d[\d,]*)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex LogicalReadMetricRegex = new Regex(
            @"(?<label>logical reads|lob logical reads)\s+(?<value>\d[\d,]*)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex TimeBlockRegex = new Regex(
            @"(?:SQL Server Execution Times|SQL Server parse and compile time):\s*(?:\r?\n\s*)?CPU time\s*=\s*(?<cpu>\d[\d,]*)\s*ms,\s*elapsed time\s*=\s*(?<elapsed>\d[\d,]*)\s*ms\.",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static StatisticsSummary Parse(string text)
        {
            if (!GridAccess.LooksLikeStatisticsOutput(text))
            {
                return null;
            }

            var summary = new StatisticsSummary
            {
                RawText = text,
            };

            var tablesByName = new Dictionary<string, StatisticsTableSummary>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in TableLineRegex.Matches(text))
            {
                var tableName = match.Groups["name"].Value;
                if (!tablesByName.TryGetValue(tableName, out var tableSummary))
                {
                    tableSummary = new StatisticsTableSummary
                    {
                        TableName = tableName,
                    };
                    tablesByName.Add(tableName, tableSummary);
                }

                var metricsText = match.Groups["metrics"].Value;
                var scanCountMatch = ScanCountRegex.Match(metricsText);
                if (scanCountMatch.Success)
                {
                    tableSummary.ScanCount += ParseLong(scanCountMatch.Groups["value"].Value);
                }

                foreach (Match readMatch in LogicalReadMetricRegex.Matches(metricsText))
                {
                    tableSummary.TotalReads += ParseLong(readMatch.Groups["value"].Value);
                }
            }

            long totalCpu = 0;
            long totalElapsed = 0;
            bool sawExecutionTime = false;
            foreach (Match match in TimeBlockRegex.Matches(text))
            {
                sawExecutionTime = true;
                totalCpu += ParseLong(match.Groups["cpu"].Value);
                totalElapsed += ParseLong(match.Groups["elapsed"].Value);
            }

            foreach (var table in tablesByName.Values)
            {
                summary.Tables.Add(table);
                summary.TotalReads += table.TotalReads;
            }

            summary.Tables.Sort((left, right) => right.TotalReads.CompareTo(left.TotalReads));

            if (sawExecutionTime)
            {
                summary.TotalCpuMilliseconds = totalCpu;
                summary.TotalElapsedMilliseconds = totalElapsed;
            }

            return summary.HasData ? summary : null;
        }

        private static long ParseLong(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return 0;
            }

            value = value.Replace(",", string.Empty).Trim();
            return long.TryParse(value, out var parsedValue) ? parsedValue : 0;
        }
    }
}