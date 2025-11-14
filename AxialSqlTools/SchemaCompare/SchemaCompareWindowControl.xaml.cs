using Microsoft.SqlServer.Dac.Compare;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace AxialSqlTools
{
    /// <summary>
    /// Interaction logic for SchemaCompareWindowControl.
    /// </summary>
    public partial class SchemaCompareWindowControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SchemaCompareWindowControl"/> class.
        /// </summary>
        public SchemaCompareWindowControl()
        {
            InitializeComponent();
            DataContext = new SchemaCompareViewModel();
        }
    }

    internal class SchemaCompareViewModel : INotifyPropertyChanged
    {
        private ScriptFactoryAccess.ConnectionInfo _sourceConnection;
        private ScriptFactoryAccess.ConnectionInfo _targetConnection;
        private bool _isBusy;
        private string _status = "Select a source and target using Object Explorer.";
        private string _deploymentScript = string.Empty;

        public SchemaCompareViewModel()
        {
            Differences = new ObservableCollection<SchemaDifferenceViewModel>();
            PickSourceCommand = new RelayCommand(() => SetConnection(isSource: true));
            PickTargetCommand = new RelayCommand(() => SetConnection(isSource: false));
            ClearSourceCommand = new RelayCommand(() => SourceConnection = null, () => SourceConnection != null);
            ClearTargetCommand = new RelayCommand(() => TargetConnection = null, () => TargetConnection != null);
            SwapConnectionsCommand = new RelayCommand(SwapConnections, () => SourceConnection != null || TargetConnection != null);
            CompareCommand = new AsyncRelayCommand(CompareAsync, () => HasConnections && !IsBusy);
        }

        public ObservableCollection<SchemaDifferenceViewModel> Differences { get; }

        public ICommand PickSourceCommand { get; }

        public ICommand PickTargetCommand { get; }

        public ICommand ClearSourceCommand { get; }

        public ICommand ClearTargetCommand { get; }

        public ICommand SwapConnectionsCommand { get; }

        public ICommand CompareCommand { get; }

        public bool HasConnections => SourceConnection != null && TargetConnection != null;

        public ScriptFactoryAccess.ConnectionInfo SourceConnection
        {
            get => _sourceConnection;
            private set
            {
                _sourceConnection = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(SourceDisplay));
                OnPropertyChanged(nameof(HasConnections));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public ScriptFactoryAccess.ConnectionInfo TargetConnection
        {
            get => _targetConnection;
            private set
            {
                _targetConnection = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(TargetDisplay));
                OnPropertyChanged(nameof(HasConnections));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public string SourceDisplay => SourceConnection?.DisplayName ?? "No source selected";

        public string TargetDisplay => TargetConnection?.DisplayName ?? "No target selected";

        public bool IsBusy
        {
            get => _isBusy;
            private set
            {
                if (_isBusy != value)
                {
                    _isBusy = value;
                    OnPropertyChanged();
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        public string Status
        {
            get => _status;
            private set
            {
                _status = value;
                OnPropertyChanged();
            }
        }

        public string DeploymentScript
        {
            get => _deploymentScript;
            private set
            {
                _deploymentScript = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void SetConnection(bool isSource)
        {
            var connection = ScriptFactoryAccess.GetCurrentConnectionInfoFromObjectExplorer();
            if (connection == null)
            {
                Status = "Select a node in Object Explorer that has a database connection.";
                return;
            }

            if (isSource)
            {
                SourceConnection = connection;
            }
            else
            {
                TargetConnection = connection;
            }
        }

        private void SwapConnections()
        {
            (SourceConnection, TargetConnection) = (TargetConnection, SourceConnection);
            OnPropertyChanged(nameof(SourceConnection));
            OnPropertyChanged(nameof(TargetConnection));
            OnPropertyChanged(nameof(SourceDisplay));
            OnPropertyChanged(nameof(TargetDisplay));
            OnPropertyChanged(nameof(HasConnections));
            CommandManager.InvalidateRequerySuggested();
        }

        private async Task CompareAsync()
        {
            if (!HasConnections)
            {
                Status = "Select both source and target connections before comparing.";
                return;
            }

            try
            {
                IsBusy = true;
                Status = "Comparing schemas...";
                DeploymentScript = string.Empty;
                Differences.Clear();

                var sourceConnectionString = SourceConnection.FullConnectionString;
                var targetConnectionString = TargetConnection.FullConnectionString;
                var targetDatabase = string.IsNullOrWhiteSpace(TargetConnection.Database) ? "Target" : TargetConnection.Database;

                var payload = await Task.Run(() =>
                {
                    var sourceEndpoint = new SchemaCompareDatabaseEndpoint(sourceConnectionString);
                    var targetEndpoint = new SchemaCompareDatabaseEndpoint(targetConnectionString);
                    var comparison = new SchemaComparison(sourceEndpoint, targetEndpoint);
                    var result = comparison.Compare();
                    var scriptResult = result.GenerateScript(targetDatabase);
                    return (Result: result, Script: scriptResult?.Script ?? string.Empty);
                });

                foreach (var difference in payload.Result.Differences
                    .OrderBy(d => d.Name?.ToString() ?? d.DifferenceType.ToString()))
                {
                    Differences.Add(new SchemaDifferenceViewModel(difference));
                }

                DeploymentScript = payload.Script;
                Status = Differences.Count == 0 ? "No differences were found." : $"Found {Differences.Count} difference(s).";
            }
            catch (Exception ex)
            {
                Status = ex.Message;
                MessageBox.Show(ex.Message, "Schema Compare", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    internal class SchemaDifferenceViewModel
    {
        public SchemaDifferenceViewModel(SchemaDifference difference)
        {
            Name = difference.Name?.ToString() ?? string.Empty;
            DifferenceType = difference.DifferenceType.ToString();
            Action = difference.UpdateAction.ToString();
            SourceObject = difference.SourceObject?.ToString() ?? string.Empty;
            TargetObject = difference.TargetObject?.ToString() ?? string.Empty;
        }

        public string Name { get; }

        public string DifferenceType { get; }

        public string Action { get; }

        public string SourceObject { get; }

        public string TargetObject { get; }
    }

    internal class AsyncRelayCommand : ICommand
    {
        private readonly Func<Task> _executeAsync;
        private readonly Func<bool> _canExecute;
        private bool _isExecuting;

        public AsyncRelayCommand(Func<Task> executeAsync, Func<bool> canExecute = null)
        {
            _executeAsync = executeAsync ?? throw new ArgumentNullException(nameof(executeAsync));
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return !_isExecuting && (_canExecute?.Invoke() ?? true);
        }

        public async void Execute(object parameter)
        {
            if (!CanExecute(parameter))
            {
                return;
            }

            _isExecuting = true;
            RaiseCanExecuteChanged();

            try
            {
                await _executeAsync();
            }
            finally
            {
                _isExecuting = false;
                RaiseCanExecuteChanged();
            }
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        private void RaiseCanExecuteChanged()
        {
            CommandManager.InvalidateRequerySuggested();
        }
    }
}
