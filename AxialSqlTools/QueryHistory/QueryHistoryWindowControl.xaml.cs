namespace AxialSqlTools
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using Microsoft.Data.SqlClient; // For ADO.NET SQL Server access
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;

    /// <summary>
    /// Interaction logic for QueryHistoryWindowControl.
    /// </summary>
    public partial class QueryHistoryWindowControl : UserControl
    {
        public QueryHistoryWindowControl()
        {
            InitializeComponent();
            DataContext = new QueryHistoryViewModel();
        }

        // Model representing a record from the QueryHistory table.
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

        // A simple RelayCommand implementation for command binding.
        public class RelayCommand : ICommand
        {
            private readonly Action _execute;
            private readonly Func<bool> _canExecute;

            public RelayCommand(Action execute, Func<bool> canExecute = null)
            {
                _execute = execute ?? throw new ArgumentNullException(nameof(execute));
                _canExecute = canExecute;
            }

            public bool CanExecute(object parameter) => _canExecute == null || _canExecute();

            public void Execute(object parameter) => _execute();

            public event EventHandler CanExecuteChanged
            {
                add { CommandManager.RequerySuggested += value; }
                remove { CommandManager.RequerySuggested -= value; }
            }
        }

        // The ViewModel that backs your UI.
        public class QueryHistoryViewModel : INotifyPropertyChanged
        {
            // Holds the records loaded from SQL Server.
            private List<QueryHistoryRecord> _allRecords;

            // The list bound to the DataGrid.
            public ObservableCollection<QueryHistoryRecord> QueryHistoryRecords { get; set; }

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

            private string _filterText;
            public string FilterText
            {
                get => _filterText;
                set
                {
                    if (_filterText != value)
                    {
                        _filterText = value;
                        OnPropertyChanged(nameof(FilterText));
                    }
                }
            }

            // RefreshCommand triggers reloading of data with the current filter.
            public ICommand RefreshCommand { get; }
            public ICommand ClearFilterCommand { get; }

            public QueryHistoryViewModel()
            {
                QueryHistoryRecords = new ObservableCollection<QueryHistoryRecord>();
                RefreshCommand = new RelayCommand(RefreshData);
                ClearFilterCommand = new RelayCommand(() => { FilterText = string.Empty; RefreshData(); });
                // Initial load.
                RefreshData();
            }

            // Loads up to 1000 records from SQL Server using a filter (if provided).
            private void RefreshData()
            {
                _allRecords = new List<QueryHistoryRecord>();

                // Retrieve connection and table settings (update as needed).
                string connectionString = SettingsManager.GetQueryHistoryConnectionString();
                string qhTableName = SettingsManager.GetQueryHistoryTableNameOrDefault();

                // Build the query.
                string query = $@"
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

                if (!string.IsNullOrWhiteSpace(FilterText))
                {
                    query += "WHERE [QueryText] LIKE @FilterText ";
                }

                query += "ORDER BY QueryID DESC";

                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            if (!string.IsNullOrWhiteSpace(FilterText))
                            {
                                command.Parameters.AddWithValue("@FilterText", "%" + FilterText + "%");
                            }
                            using (SqlDataReader reader = command.ExecuteReader())
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

                                    // Limit the short version of QueryText to 100 characters.
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
                    MessageBox.Show("Error loading data: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                // Update the bound collection.
                QueryHistoryRecords.Clear();
                foreach (var record in _allRecords)
                {
                    QueryHistoryRecords.Add(record);
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
