using System;
using System.Collections.Generic;

namespace AxialSqlTools
{
    public class StatisticsSummary
    {
        public DateTime CapturedAt { get; set; } = DateTime.Now;
        public string DataSource { get; set; }
        public string DatabaseName { get; set; }
        public string QueryText { get; set; }
        public string RawText { get; set; }
        public int? TotalCpuMilliseconds { get; set; }
        public int? TotalElapsedMilliseconds { get; set; }
        public long TotalReads { get; set; }
        public List<StatisticsTableSummary> Tables { get; } = new List<StatisticsTableSummary>();

        public bool HasData => Tables.Count > 0 || TotalCpuMilliseconds.HasValue || TotalElapsedMilliseconds.HasValue;
    }

    public class StatisticsTableSummary
    {
        public string TableName { get; set; }
        public long ScanCount { get; set; }
        public long TotalReads { get; set; }
    }
}