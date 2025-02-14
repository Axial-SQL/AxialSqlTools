namespace AxialSqlTools
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Linq;

    /// <summary>
    /// Interaction logic for QueryHistoryWindowControl.
    /// </summary>
    public partial class QueryHistoryWindowControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="QueryHistoryWindowControl"/> class.
        /// </summary>
        public QueryHistoryWindowControl()
        {
            this.InitializeComponent();

            DataContext = new QueryHistoryViewModel();

        }

        // Model representing a record from the QueryHistory table
        public class QueryHistoryRecord
        {
            public int Id { get; set; }
            public DateTime Date { get; set; }
            public string QueryText { get; set; }
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
            // Complete set of records loaded from a data source.
            private List<QueryHistoryRecord> _allRecords;

            // The paginated and filtered list bound to the DataGrid.
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
                        CurrentPage = 1; // Reset to first page on filter change.
                        OnPropertyChanged(nameof(FilterText));
                        UpdateRecords();
                    }
                }
            }

            private int _currentPage = 1;
            public int CurrentPage
            {
                get => _currentPage;
                set
                {
                    if (_currentPage != value)
                    {
                        _currentPage = value;
                        OnPropertyChanged(nameof(CurrentPage));
                        UpdateRecords();
                    }
                }
            }

            private int _pageSize = 100; // Number of records per page.
            public int PageSize
            {
                get => _pageSize;
                set
                {
                    if (_pageSize != value)
                    {
                        _pageSize = value;
                        OnPropertyChanged(nameof(PageSize));
                        UpdateRecords();
                    }
                }
            }

            public ICommand NextPageCommand { get; }
            public ICommand PreviousPageCommand { get; }

            public QueryHistoryViewModel()
            {
                QueryHistoryRecords = new ObservableCollection<QueryHistoryRecord>();
                NextPageCommand = new RelayCommand(NextPage, CanGoNextPage);
                PreviousPageCommand = new RelayCommand(PreviousPage, CanGoPreviousPage);

                // Simulate loading records (replace this with your actual data retrieval)
                LoadDummyData();

                // Display the initial page
                UpdateRecords();
            }

            // Dummy data generator for demonstration purposes.
            private void LoadDummyData()
            {
                _allRecords = new List<QueryHistoryRecord>();
                for (int i = 1; i <= 100; i++)
                {
                    _allRecords.Add(new QueryHistoryRecord
                    {
                        Id = i,
                        Date = DateTime.Now.AddMinutes(-i * 10),
                        QueryText = $"SELECT * FROM Table WHERE Id = {i}; -- Sample query {i}"
                    });
                }
            }

            // Apply filtering and pagination to update the displayed records.
            private void UpdateRecords()
            {
                // Filter the records based on the FilterText (if provided)
                var filtered = string.IsNullOrWhiteSpace(FilterText)
                    ? _allRecords
                    : _allRecords.Where(r => r.QueryText.IndexOf(FilterText, StringComparison.OrdinalIgnoreCase) >= 0).ToList();

                // Calculate total pages.
                int totalPages = (int)Math.Ceiling((double)filtered.Count / PageSize);
                // Ensure the current page is within the valid range.
                if (CurrentPage > totalPages)
                    CurrentPage = totalPages > 0 ? totalPages : 1;

                // Get records for the current page.
                var pageRecords = filtered
                    .Skip((CurrentPage - 1) * PageSize)
                    .Take(PageSize)
                    .ToList();

                // Refresh the ObservableCollection.
                QueryHistoryRecords.Clear();
                foreach (var record in pageRecords)
                {
                    QueryHistoryRecords.Add(record);
                }

                // Update command states.
                CommandManager.InvalidateRequerySuggested();
            }

            // Command logic to determine if the "Next" button can be enabled.
            private bool CanGoNextPage()
            {
                var filteredCount = string.IsNullOrWhiteSpace(FilterText)
                    ? _allRecords.Count
                    : _allRecords.Count(r => r.QueryText.IndexOf(FilterText, StringComparison.OrdinalIgnoreCase) >= 0);
                int totalPages = (int)Math.Ceiling((double)filteredCount / PageSize);
                return CurrentPage < totalPages;
            }

            // Go to the next page.
            private void NextPage()
            {
                if (CanGoNextPage())
                {
                    CurrentPage++;
                }
            }

            // Command logic to determine if the "Previous" button can be enabled.
            private bool CanGoPreviousPage() => CurrentPage > 1;

            // Go to the previous page.
            private void PreviousPage()
            {
                if (CanGoPreviousPage())
                {
                    CurrentPage--;
                }
            }

            #region INotifyPropertyChanged Implementation

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }

            #endregion
        }
    }
}