using System;

namespace AxialSqlTools
{
    public class QueryHistoryRecord
    {
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public DateTime FinishTime { get; set; }
        public string ElapsedTime { get; set; }
        public long TotalRowsReturned { get; set; }
        public string ExecResult { get; set; }
        public string QueryText { get; set; }
        public string QueryTextShort { get; set; }
        public string DataSource { get; set; }
        public string DatabaseName { get; set; }
        public string LoginName { get; set; }
        public string WorkstationId { get; set; }
    }
}
