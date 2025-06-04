using Microsoft.Data.SqlClient; // Make sure you have Microsoft.Data.SqlClient
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Windows;
using System.Windows.Input;
using static AxialSqlTools.QueryHistoryWindowControl;

namespace AxialSqlTools
{
    public class QueryHistoryViewModel : INotifyPropertyChanged
    {
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

            // Retrieve connection and table settings (update as needed).
            string connectionString = SettingsManager.GetQueryHistoryConnectionString();
            string qhTableName = SettingsManager.GetQueryHistoryTableNameOrDefault();

            // Base SELECT
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

            try
            {
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

                                // Create a short (ellipsized) version of QueryText
                                record.QueryTextShort = record.QueryText.Length > 100
                                    ? record.QueryText.Substring(0, 100)
                                    : record.QueryText;

                                _allRecords.Add(record);
                            }
                        }
                    }
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

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }
    }
}
