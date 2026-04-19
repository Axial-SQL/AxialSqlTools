using Microsoft.Data.SqlClient; // Make sure you have Microsoft.Data.SqlClient
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Newtonsoft.Json;
using static AxialSqlTools.QueryHistoryWindowControl;

namespace AxialSqlTools
{
    public class QueryHistoryViewModel : INotifyPropertyChanged
    {
        private const string QueryHistoryStorageModeTextFiles = "TextFiles";
        private const string QueryHistoryStorageModeDisabled = "Disabled";
        private class QueryHistoryFileEntry
        {
            public DateTime StartTime { get; set; }
            public DateTime FinishTime { get; set; }
            public string ElapsedTime { get; set; }
            public long TotalRowsReturned { get; set; }
            public string ExecResult { get; set; }
            public string QueryText { get; set; }
            public string DataSource { get; set; }
            public string DatabaseName { get; set; }
            public string LoginName { get; set; }
            public string WorkstationId { get; set; }
        }

        // Underlying full list (before filtering)
        private List<QueryHistoryRecord> _allRecords;

        // ObservableCollection bound to DataGrid
        public ObservableCollection<QueryHistoryRecord> QueryHistoryRecords { get; }

        private QueryHistoryRecord _selectedRecord;
        public QueryHistoryRecord SelectedRecord
        {
            get => _selectedRecord;
            set
            {
                if (_selectedRecord != value)
                {
                    _selectedRecord = value;
                    OnPropertyChanged(nameof(SelectedRecord));
                }
            }
        }

        // ========== FILTER PROPERTIES ==========
        private DateTime? _filterFromDate;
        public DateTime? FilterFromDate
        {
            get => _filterFromDate;
            set
            {
                if (_filterFromDate != value)
                {
                    _filterFromDate = value;
                    OnPropertyChanged(nameof(FilterFromDate));
                }
            }
        }

        private DateTime? _filterToDate;
        public DateTime? FilterToDate
        {
            get => _filterToDate;
            set
            {
                if (_filterToDate != value)
                {
                    _filterToDate = value;
                    OnPropertyChanged(nameof(FilterToDate));
                }
            }
        }

        private string _filterServer;
        public string FilterServer
        {
            get => _filterServer;
            set
            {
                if (_filterServer != value)
                {
                    _filterServer = value;
                    OnPropertyChanged(nameof(FilterServer));
                }
            }
        }

        private string _filterDatabase;
        public string FilterDatabase
        {
            get => _filterDatabase;
            set
            {
                if (_filterDatabase != value)
                {
                    _filterDatabase = value;
                    OnPropertyChanged(nameof(FilterDatabase));
                }
            }
        }

        private string _filterQueryText;
        public string FilterQueryText
        {
            get => _filterQueryText;
            set
            {
                if (_filterQueryText != value)
                {
                    _filterQueryText = value;
                    OnPropertyChanged(nameof(FilterQueryText));
                }
            }
        }

        // Commands
        public ICommand RefreshCommand { get; }
        public ICommand ClearFilterCommand { get; }

        public QueryHistoryViewModel()
        {
            QueryHistoryRecords = new ObservableCollection<QueryHistoryRecord>();
            RefreshCommand = new RelayCommand(RefreshData);
            ClearFilterCommand = new RelayCommand(ClearAllFilters);

            // Optionally, initialize filters to today’s date range or leave null
            // FilterFromDate = DateTime.Today.AddDays(-7);
            // FilterToDate = DateTime.Today;

            RefreshData();
        }

        private void ClearAllFilters()
        {
            FilterFromDate = null;
            FilterToDate = null;
            FilterServer = string.Empty;
            FilterDatabase = string.Empty;
            FilterQueryText = string.Empty;
            RefreshData();
        }

        private void RefreshData()
        {
            _allRecords = new List<QueryHistoryRecord>();

            try
            {
                string storageMode = SettingsManager.GetQueryHistoryStorageMode();
                if (string.Equals(storageMode, QueryHistoryStorageModeDisabled, StringComparison.OrdinalIgnoreCase))
                {
                    // Query history is disabled; keep the list empty.
                }
                else if (string.Equals(storageMode, QueryHistoryStorageModeTextFiles, StringComparison.OrdinalIgnoreCase))
                {
                    LoadFromTextFiles();
                }
                else
                {
                    LoadFromDatabase();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading data: " + ex.Message,
                                "Error",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }

            // Push into the ObservableCollection
            QueryHistoryRecords.Clear();
            foreach (var rec in _allRecords)
            {
                QueryHistoryRecords.Add(rec);
            }
        }

        private void LoadFromDatabase()
        {
            string connectionString = SettingsManager.GetQueryHistoryConnectionString();
            string qhTableName = SettingsManager.GetQueryHistoryTableNameOrDefault();
            string sql = $@"
                SELECT TOP 1000 
                    [QueryID],
                    [StartTime],
                    [FinishTime],
                    [ElapsedTime],
                    [TotalRowsReturned],
                    [ExecResult],
                    [QueryText],
                    [DataSource],
                    [DatabaseName],
                    [LoginName],
                    [WorkstationId]
                FROM {qhTableName} ";

            // Build WHERE clauses dynamically
            var whereClauses = new List<string>();
            var parameters = new List<SqlParameter>();

            if (FilterFromDate.HasValue)
            {
                whereClauses.Add("[StartTime] >= @FromDate");
                parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime) { Value = FilterFromDate.Value.Date });
            }
            if (FilterToDate.HasValue)
            {
                // Include the entire day for the “To” date (set time to 23:59:59)
                DateTime endOfDay = FilterToDate.Value.Date.AddDays(1).AddSeconds(-1);
                whereClauses.Add("[StartTime] <= @ToDate");
                parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime) { Value = endOfDay });
            }
            if (!string.IsNullOrWhiteSpace(FilterServer))
            {
                whereClauses.Add("[DataSource] LIKE @Server");
                parameters.Add(new SqlParameter("@Server", SqlDbType.NVarChar, 256) { Value = "%" + FilterServer + "%" });
            }
            if (!string.IsNullOrWhiteSpace(FilterDatabase))
            {
                whereClauses.Add("[DatabaseName] LIKE @Database");
                parameters.Add(new SqlParameter("@Database", SqlDbType.NVarChar, 256) { Value = "%" + FilterDatabase + "%" });
            }
            if (!string.IsNullOrWhiteSpace(FilterQueryText))
            {
                whereClauses.Add("[QueryText] LIKE @QueryText");
                parameters.Add(new SqlParameter("@QueryText", SqlDbType.NVarChar) { Value = "%" + FilterQueryText + "%" });
            }

            if (whereClauses.Count > 0)
            {
                sql += "WHERE " + string.Join(" AND ", whereClauses) + " ";
            }

            sql += "ORDER BY [QueryID] DESC;";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    foreach (var p in parameters)
                        cmd.Parameters.Add(p);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var record = new QueryHistoryRecord
                            {
                                Id = reader.GetInt32(reader.GetOrdinal("QueryID")),
                                Date = reader.GetDateTime(reader.GetOrdinal("StartTime")),
                                FinishTime = reader.GetDateTime(reader.GetOrdinal("FinishTime")),
                                ElapsedTime = reader.GetString(reader.GetOrdinal("ElapsedTime")),
                                TotalRowsReturned = reader.GetInt64(reader.GetOrdinal("TotalRowsReturned")),
                                ExecResult = reader.GetString(reader.GetOrdinal("ExecResult")),
                                QueryText = reader.GetString(reader.GetOrdinal("QueryText")),
                                DataSource = reader.GetString(reader.GetOrdinal("DataSource")),
                                DatabaseName = reader.GetString(reader.GetOrdinal("DatabaseName")),
                                LoginName = reader.GetString(reader.GetOrdinal("LoginName")),
                                WorkstationId = reader.GetString(reader.GetOrdinal("WorkstationId"))
                            };

                            record.QueryTextShort = BuildShortQueryText(record.QueryText);
                            _allRecords.Add(record);
                        }
                    }
                }
            }
        }

        private void LoadFromTextFiles()
        {
            string folder = SettingsManager.GetQueryHistoryTextFileFolder();
            if (!Directory.Exists(folder))
            {
                return;
            }

            var records = new List<QueryHistoryRecord>();
            foreach (string filePath in Directory.GetFiles(folder, "*.jsonl", SearchOption.TopDirectoryOnly))
            {
                foreach (string line in File.ReadLines(filePath))
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    try
                    {
                        var fileEntry = JsonConvert.DeserializeObject<QueryHistoryFileEntry>(line);
                        if (fileEntry == null)
                        {
                            continue;
                        }

                        var record = new QueryHistoryRecord
                        {
                            Date = fileEntry.StartTime,
                            FinishTime = fileEntry.FinishTime,
                            ElapsedTime = fileEntry.ElapsedTime ?? string.Empty,
                            TotalRowsReturned = fileEntry.TotalRowsReturned,
                            ExecResult = fileEntry.ExecResult ?? string.Empty,
                            QueryText = fileEntry.QueryText ?? string.Empty,
                            DataSource = fileEntry.DataSource ?? string.Empty,
                            DatabaseName = fileEntry.DatabaseName ?? string.Empty,
                            LoginName = fileEntry.LoginName ?? string.Empty,
                            WorkstationId = fileEntry.WorkstationId ?? string.Empty
                        };

                        record.QueryTextShort = BuildShortQueryText(record.QueryText);
                        records.Add(record);
                    }
                    catch
                    {
                        // ignore malformed lines and continue
                    }
                }
            }

            IEnumerable<QueryHistoryRecord> filtered = records;
            if (FilterFromDate.HasValue)
            {
                filtered = filtered.Where(r => r.Date >= FilterFromDate.Value.Date);
            }
            if (FilterToDate.HasValue)
            {
                DateTime endOfDay = FilterToDate.Value.Date.AddDays(1).AddSeconds(-1);
                filtered = filtered.Where(r => r.Date <= endOfDay);
            }
            if (!string.IsNullOrWhiteSpace(FilterServer))
            {
                filtered = filtered.Where(r => (r.DataSource ?? string.Empty).IndexOf(FilterServer, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            if (!string.IsNullOrWhiteSpace(FilterDatabase))
            {
                filtered = filtered.Where(r => (r.DatabaseName ?? string.Empty).IndexOf(FilterDatabase, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            if (!string.IsNullOrWhiteSpace(FilterQueryText))
            {
                filtered = filtered.Where(r => (r.QueryText ?? string.Empty).IndexOf(FilterQueryText, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            _allRecords = filtered
                .OrderByDescending(r => r.Date)
                .Take(1000)
                .Select((r, idx) =>
                {
                    r.Id = idx + 1;
                    return r;
                })
                .ToList();
        }

        private static string BuildShortQueryText(string queryText)
        {
            queryText = queryText ?? string.Empty;
            return queryText.Length > 100
                ? queryText.Substring(0, 100)
                : queryText;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }
    }
}
