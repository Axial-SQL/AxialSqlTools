using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using Microsoft.SqlServer.Dac.Compare;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
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
        private readonly SideBySideDiffBuilder _diffBuilder = new SideBySideDiffBuilder(new Differ());
        private bool _isBusy;
        private string _status = "Select a source and target using Object Explorer.";
        private string _deploymentScript = string.Empty;
        private SchemaDifferenceViewModel _selectedDifference;

        public SchemaCompareViewModel()
        {
            Differences = new ObservableCollection<SchemaDifferenceViewModel>();
            SourceDiffLines = new ObservableCollection<DiffLineViewModel>();
            TargetDiffLines = new ObservableCollection<DiffLineViewModel>();
            PickSourceCommand = new RelayCommand(() => SetConnection(isSource: true));
            PickTargetCommand = new RelayCommand(() => SetConnection(isSource: false));
            ClearSourceCommand = new RelayCommand(() => SourceConnection = null, () => SourceConnection != null);
            ClearTargetCommand = new RelayCommand(() => TargetConnection = null, () => TargetConnection != null);
            SwapConnectionsCommand = new RelayCommand(SwapConnections, () => SourceConnection != null || TargetConnection != null);
            CompareCommand = new AsyncRelayCommand(CompareAsync, () => HasConnections && !IsBusy);
        }

        public ObservableCollection<SchemaDifferenceViewModel> Differences { get; }

        public ObservableCollection<DiffLineViewModel> SourceDiffLines { get; }

        public ObservableCollection<DiffLineViewModel> TargetDiffLines { get; }

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

        public SchemaDifferenceViewModel SelectedDifference
        {
            get => _selectedDifference;
            set
            {
                if (_selectedDifference != value)
                {
                    _selectedDifference = value;
                    OnPropertyChanged();
                    UpdateDiffPreview(value);
                }
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
                SelectedDifference = null;
                SourceDiffLines.Clear();
                TargetDiffLines.Clear();

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

        private void UpdateDiffPreview(SchemaDifferenceViewModel difference)
        {
            SourceDiffLines.Clear();
            TargetDiffLines.Clear();

            if (difference == null)
            {
                return;
            }

            var diffModel = _diffBuilder.BuildDiffModel(difference.SourceDefinition ?? string.Empty, difference.TargetDefinition ?? string.Empty);

            if (diffModel?.OldText?.Lines != null)
            {
                foreach (var line in diffModel.OldText.Lines)
                {
                    SourceDiffLines.Add(new DiffLineViewModel(line));
                }
            }

            if (diffModel?.NewText?.Lines != null)
            {
                foreach (var line in diffModel.NewText.Lines)
                {
                    TargetDiffLines.Add(new DiffLineViewModel(line));
                }
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
            SourceDefinition = SchemaCompareObjectDefinitionExtractor.GetDefinition(difference.SourceObject);
            TargetDefinition = SchemaCompareObjectDefinitionExtractor.GetDefinition(difference.TargetObject);
        }

        public string Name { get; }

        public string DifferenceType { get; }

        public string Action { get; }

        public string SourceObject { get; }

        public string TargetObject { get; }

        public string SourceDefinition { get; }

        public string TargetDefinition { get; }
    }

    internal static class SchemaCompareObjectDefinitionExtractor
    {
        private static readonly string[] PreferredProperties = new[] { "Definition", "Script", "ObjectDefinition" };
        private static readonly string[] PreferredMethods = new[] { "GetScript", "GetDefinition", "GetSourceDefinition" };

        public static string GetDefinition(SchemaCompareObject schemaObject)
        {
            if (schemaObject == null)
            {
                return string.Empty;
            }

            var type = schemaObject.GetType();

            foreach (var propertyName in PreferredProperties)
            {
                var property = type.GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
                if (property?.PropertyType == typeof(string))
                {
                    if (property.GetValue(schemaObject) is string value && !string.IsNullOrWhiteSpace(value))
                    {
                        return value;
                    }
                }
            }

            foreach (var methodName in PreferredMethods)
            {
                var method = type.GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
                if (method?.ReturnType == typeof(string))
                {
                    if (method.Invoke(schemaObject, null) is string value && !string.IsNullOrWhiteSpace(value))
                    {
                        return value;
                    }
                }
            }

            return schemaObject.ToString() ?? string.Empty;
        }
    }

    internal class DiffLineViewModel
    {
        public DiffLineViewModel(DiffPiece piece)
        {
            Text = piece?.Text ?? string.Empty;
            LineNumber = piece?.Position?.ToString() ?? string.Empty;
            DiffType = piece?.Type.ToString() ?? ChangeType.Unchanged.ToString();
            Indicator = GetIndicator(piece?.Type ?? ChangeType.Unchanged);
            Opacity = (piece?.Type ?? ChangeType.Unchanged) == ChangeType.Imaginary ? 0.5 : 1.0;
        }

        public string Text { get; }

        public string LineNumber { get; }

        public string DiffType { get; }

        public string Indicator { get; }

        public double Opacity { get; }

        private static string GetIndicator(ChangeType changeType)
        {
            switch (changeType)
            {
                case ChangeType.Deleted:
                    return "-";
                case ChangeType.Inserted:
                    return "+";
                case ChangeType.Modified:
                    return "*";
                default:
                    return string.Empty;
            }
        }
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
