using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using Microsoft.SqlServer.Dac.Compare;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Text;
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
        private CancellationTokenSource _cancellationTokenSource;
        private SchemaComparisonResult _comparisonResult;
        private string _targetDatabaseName = string.Empty;

        public SchemaCompareViewModel()
        {
            Differences = new ObservableCollection<SchemaDifferenceViewModel>();
            SourceDiffLines = new ObservableCollection<DiffLineViewModel>();
            TargetDiffLines = new ObservableCollection<DiffLineViewModel>();
            ObjectTypeFilters = new ObservableCollection<ObjectTypeFilterViewModel>(CreateObjectTypeFilters());
            PickSourceCommand = new RelayCommand(() => SetConnection(isSource: true));
            PickTargetCommand = new RelayCommand(() => SetConnection(isSource: false));
            ClearSourceCommand = new RelayCommand(() => SourceConnection = null, () => SourceConnection != null);
            ClearTargetCommand = new RelayCommand(() => TargetConnection = null, () => TargetConnection != null);
            SwapConnectionsCommand = new RelayCommand(
                SwapConnections,
                () => !IsBusy && (SourceConnection != null || TargetConnection != null));
            CompareCommand = new AsyncRelayCommand(CompareAsync, () => HasConnections && !IsBusy);
            CancelCompareCommand = new RelayCommand(CancelCompare, () => IsBusy);
            GenerateDeploymentScriptCommand = new AsyncRelayCommand(GenerateDeploymentScriptAsync, () => _comparisonResult != null && Differences.Any(d => d.IsSelected));
            CopyDeploymentScriptCommand = new RelayCommand(CopyDeploymentScript, () => !string.IsNullOrEmpty(DeploymentScript));
            CopySourceDefinitionCommand = new RelayCommand(CopySourceDefinition, () => !string.IsNullOrEmpty(SelectedDifference?.SourceDefinition));
            CopyTargetDefinitionCommand = new RelayCommand(CopyTargetDefinition, () => !string.IsNullOrEmpty(SelectedDifference?.TargetDefinition));
            IncludeAllObjectTypesCommand = new RelayCommand(() => SetAllObjectTypeFilters(true));
            ExcludeAllObjectTypesCommand = new RelayCommand(() => SetAllObjectTypeFilters(false));
        }

        public ObservableCollection<SchemaDifferenceViewModel> Differences { get; }

        public ObservableCollection<DiffLineViewModel> SourceDiffLines { get; }

        public ObservableCollection<DiffLineViewModel> TargetDiffLines { get; }

        public ObservableCollection<ObjectTypeFilterViewModel> ObjectTypeFilters { get; }

        public ICommand PickSourceCommand { get; }

        public ICommand PickTargetCommand { get; }

        public ICommand ClearSourceCommand { get; }

        public ICommand ClearTargetCommand { get; }

        public ICommand SwapConnectionsCommand { get; }

        public ICommand CompareCommand { get; }

        public ICommand CancelCompareCommand { get; }

        public ICommand GenerateDeploymentScriptCommand { get; }

        public ICommand CopyDeploymentScriptCommand { get; }

        public ICommand CopySourceDefinitionCommand { get; }

        public ICommand CopyTargetDefinitionCommand { get; }

        public ICommand IncludeAllObjectTypesCommand { get; }

        public ICommand ExcludeAllObjectTypesCommand { get; }

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
                CommandManager.InvalidateRequerySuggested();
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
                    CommandManager.InvalidateRequerySuggested();
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
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = new CancellationTokenSource();
                var cancellationToken = _cancellationTokenSource.Token;
                Status = "Comparing schemas...";
                DeploymentScript = string.Empty;
                Differences.Clear();
                SelectedDifference = null;
                SourceDiffLines.Clear();
                TargetDiffLines.Clear();
                _comparisonResult = null;
                _targetDatabaseName = string.Empty;

                var sourceConnectionString = SourceConnection.FullConnectionString;
                var targetConnectionString = TargetConnection.FullConnectionString;
                var targetDatabase = string.IsNullOrWhiteSpace(TargetConnection.Database) ? "Target" : TargetConnection.Database;

                var payload = await Task.Run(() =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var sourceEndpoint = new SchemaCompareDatabaseEndpoint(sourceConnectionString);
                    var targetEndpoint = new SchemaCompareDatabaseEndpoint(targetConnectionString);
                    var comparison = new SchemaComparison(sourceEndpoint, targetEndpoint);
                    ApplyObjectTypeFilters(comparison);
                    var result = comparison.Compare();
                    cancellationToken.ThrowIfCancellationRequested();
                    var scriptResult = result.GenerateScript(targetDatabase);
                    cancellationToken.ThrowIfCancellationRequested();
                    return (Result: result, Script: scriptResult?.Script ?? string.Empty);
                }, cancellationToken);

                foreach (var difference in payload.Result.Differences
                    .OrderBy(d => d.Name?.ToString() ?? d.DifferenceType.ToString()))
                {
                    Differences.Add(new SchemaDifferenceViewModel(difference, payload.Result));
                }

                DeploymentScript = payload.Script;
                _comparisonResult = payload.Result;
                _targetDatabaseName = targetDatabase;
                Status = Differences.Count == 0 ? "No differences were found." : $"Found {Differences.Count} difference(s).";
                CommandManager.InvalidateRequerySuggested();
            }
            catch (OperationCanceledException)
            {
                Status = "Schema compare canceled.";
            }
            catch (Exception ex)
            {
                Status = ex.Message;
                MessageBox.Show(ex.Message, "Schema Compare", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                IsBusy = false;
            }
        }

        private void CancelCompare()
        {
            if (!IsBusy)
            {
                return;
            }

            Status = "Cancelling schema compare...";
            _cancellationTokenSource?.Cancel();
        }

        private async Task GenerateDeploymentScriptAsync()
        {
            if (_comparisonResult == null)
            {
                Status = "Run a comparison before generating a deployment script.";
                return;
            }

            try
            {
                IsBusy = true;
                var selectedCount = Differences.Count(d => d.IsSelected);
                if (selectedCount == 0)
                {
                    DeploymentScript = string.Empty;
                    Status = "No objects selected for script generation.";
                    return;
                }

                string script = string.Empty;

                await Task.Run(() =>
                {
                    foreach (var difference in Differences)
                    {
                        difference.ApplySelection();
                    }

                    script = _comparisonResult.GenerateScript(_targetDatabaseName)?.Script ?? string.Empty;
                });

                DeploymentScript = script;
                Status = $"Deployment script generated for {selectedCount} selected object(s).";
            }
            catch (Exception ex)
            {
                Status = ex.Message;
                MessageBox.Show(ex.Message, "Schema Compare", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                CommandManager.InvalidateRequerySuggested();
            }
        }

        private void CopyDeploymentScript()
        {
            CopyTextToClipboard(DeploymentScript);
        }

        private void CopySourceDefinition()
        {
            CopyTextToClipboard(SelectedDifference?.SourceDefinition);
        }

        private void CopyTargetDefinition()
        {
            CopyTextToClipboard(SelectedDifference?.TargetDefinition);
        }

        private static void CopyTextToClipboard(string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                Clipboard.SetDataObject(text);
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

            var diffModel = _diffBuilder.BuildDiffModel(difference.TargetDefinition ?? string.Empty, difference.SourceDefinition ?? string.Empty);

            if (diffModel?.NewText?.Lines != null)
            {
                foreach (var line in diffModel.NewText.Lines)
                {
                    SourceDiffLines.Add(new DiffLineViewModel(line));
                }
            }

            if (diffModel?.OldText?.Lines != null)
            {
                foreach (var line in diffModel.OldText.Lines)
                {
                    TargetDiffLines.Add(new DiffLineViewModel(line));
                }
            }
        }

        private void ApplyObjectTypeFilters(SchemaComparison comparison)
        {
            if (comparison?.Options?.ExcludedObjectTypes == null)
            {
                return;
            }

            comparison.Options.ExcludedObjectTypes.Clear();
            foreach (var filter in ObjectTypeFilters.Where(f => !f.IsIncluded))
            {
                comparison.Options.ExcludedObjectTypes.Add(filter.ObjectType);
            }
        }

        private IEnumerable<ObjectTypeFilterViewModel> CreateObjectTypeFilters()
        {
            foreach (SchemaCompareObjectType objectType in Enum.GetValues(typeof(SchemaCompareObjectType)))
            {
                yield return new ObjectTypeFilterViewModel(objectType);
            }
        }

        private void SetAllObjectTypeFilters(bool include)
        {
            foreach (var filter in ObjectTypeFilters)
            {
                filter.IsIncluded = include;
            }

            CommandManager.InvalidateRequerySuggested();
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    internal class SchemaDifferenceViewModel : INotifyPropertyChanged
    {
        private readonly SchemaDifference _difference;
        private readonly SchemaComparisonResult _result;
        private bool _isSelected;

        public SchemaDifferenceViewModel(SchemaDifference difference, SchemaComparisonResult result)
        {
            _difference = difference ?? throw new ArgumentNullException(nameof(difference));
            _result = result ?? throw new ArgumentNullException(nameof(result));
            _isSelected = difference.Included;

            Name = difference.Name?.ToString() ?? string.Empty;
            DifferenceType = difference.DifferenceType.ToString();
            Action = difference.UpdateAction.ToString();

            SourceObject = difference.SourceObject?.Name.ToString() ?? string.Empty;
            TargetObject = difference.TargetObject?.Name.ToString() ?? string.Empty;

            var sbSource = new StringBuilder();
            GetAllSourceChidren(difference, sbSource);
            SourceDefinition = sbSource.ToString();

            var sbTarget = new StringBuilder();
            GetAllTargetChidren(difference, sbTarget);
            TargetDefinition = sbTarget.ToString();
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged();
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        public void ApplySelection()
        {
            SetIncludedState(_difference, _isSelected);
        }

        private void GetAllSourceChidren(SchemaDifference difference, StringBuilder sb)
        {
            string stringResult = _result.GetDiffEntrySourceScript(difference);
            if (!string.IsNullOrWhiteSpace(stringResult))
            {
                sb.AppendLine(stringResult);
                sb.AppendLine("GO");
            }

            foreach (var diff in difference.Children)
            {
                if (diff.Included)
                {
                    GetAllSourceChidren(diff, sb);
                }
            }
        }

        private void GetAllTargetChidren(SchemaDifference difference, StringBuilder sb)
        {
            string stringResult = _result.GetDiffEntryTargetScript(difference);
            if (!string.IsNullOrWhiteSpace(stringResult))
            {
                sb.AppendLine(stringResult);
                sb.AppendLine("GO");
            }

            foreach (var diff in difference.Children)
            {
                if (diff.Included)
                {
                    GetAllTargetChidren(diff, sb);
                }
            }
        }

        private void SetIncludedState(SchemaDifference difference, bool isIncluded)
        {
            difference.Included = isIncluded;

            foreach (var child in difference.Children)
            {
                SetIncludedState(child, isIncluded);
            }
        }

        public string Name { get; }

        public string DifferenceType { get; }

        public string Action { get; }

        public string SourceObject { get; }

        public string TargetObject { get; }

        public string SourceDefinition { get; }

        public string TargetDefinition { get; }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
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

    internal class ObjectTypeFilterViewModel : INotifyPropertyChanged
    {
        private bool _isIncluded = true;

        public ObjectTypeFilterViewModel(SchemaCompareObjectType objectType)
        {
            ObjectType = objectType;
            DisplayName = objectType.ToString();
        }

        public SchemaCompareObjectType ObjectType { get; }

        public string DisplayName { get; }

        public bool IsIncluded
        {
            get => _isIncluded;
            set
            {
                if (_isIncluded != value)
                {
                    _isIncluded = value;
                    OnPropertyChanged();
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
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
