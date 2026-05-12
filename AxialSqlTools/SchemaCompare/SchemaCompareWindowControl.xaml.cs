using Azure.Storage.Blobs.Models;
using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Math;
using Microsoft.SqlServer.Dac.Compare;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;

namespace AxialSqlTools
{
    /// <summary>
    /// Interaction logic for SchemaCompareWindowControl.
    /// </summary>
    public partial class SchemaCompareWindowControl : UserControl
    {
        private readonly ToolWindowThemeController _themeController;
        private ScrollViewer _sourceScrollViewer;
        private ScrollViewer _targetScrollViewer;
        private bool _isSyncingScroll;

        /// <summary>
        /// Initializes a new instance of the <see cref="SchemaCompareWindowControl"/> class.
        /// </summary>
        public SchemaCompareWindowControl()
        {
            InitializeComponent();
            _themeController = new ToolWindowThemeController(this, ApplyThemeBrushResources);
            DataContext = new SchemaCompareViewModel();
            Loaded += OnControlLoaded;
            Unloaded += OnControlUnloaded;
        }

        private void ApplyThemeBrushResources()
        {
            ToolWindowThemeResources.ApplySharedTheme(this);
        }

        private void OnControlLoaded(object sender, RoutedEventArgs e)
        {
            HookDiffScrollSynchronization();
        }

        private void OnControlUnloaded(object sender, RoutedEventArgs e)
        {
            UnhookDiffScrollSynchronization();
        }

        private void WikiLink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            ToolWindowNavigation.HandleRequestNavigate(e);
        }

        private void HookDiffScrollSynchronization()
        {
            UnhookDiffScrollSynchronization();

            _sourceScrollViewer = FindVisualChild<ScrollViewer>(SourceDiffListView);
            _targetScrollViewer = FindVisualChild<ScrollViewer>(TargetDiffListView);

            if (_sourceScrollViewer == null || _targetScrollViewer == null)
            {
                return;
            }

            _sourceScrollViewer.ScrollChanged += SourceScrollViewer_ScrollChanged;
            _targetScrollViewer.ScrollChanged += TargetScrollViewer_ScrollChanged;
        }

        private void UnhookDiffScrollSynchronization()
        {
            if (_sourceScrollViewer != null)
            {
                _sourceScrollViewer.ScrollChanged -= SourceScrollViewer_ScrollChanged;
            }

            if (_targetScrollViewer != null)
            {
                _targetScrollViewer.ScrollChanged -= TargetScrollViewer_ScrollChanged;
            }

            _sourceScrollViewer = null;
            _targetScrollViewer = null;
            _isSyncingScroll = false;
        }

        private void SourceScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            SyncScrollOffsets(_sourceScrollViewer, _targetScrollViewer, e);
        }

        private void TargetScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            SyncScrollOffsets(_targetScrollViewer, _sourceScrollViewer, e);
        }

        private void SyncScrollOffsets(ScrollViewer from, ScrollViewer to, ScrollChangedEventArgs e)
        {
            if (_isSyncingScroll || from == null || to == null)
            {
                return;
            }

            if (e.VerticalChange == 0 && e.HorizontalChange == 0)
            {
                return;
            }

            _isSyncingScroll = true;

            try
            {
                to.ScrollToVerticalOffset(from.VerticalOffset);
                to.ScrollToHorizontalOffset(from.HorizontalOffset);
            }
            finally
            {
                _isSyncingScroll = false;
            }
        }

        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null)
            {
                return null;
            }

            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                {
                    return typedChild;
                }

                T descendant = FindVisualChild<T>(child);
                if (descendant != null)
                {
                    return descendant;
                }
            }

            return null;
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

        public SchemaCompareViewModel()
        {
            Differences = new ObservableCollection<SchemaDifferenceViewModel>();
            SourceDiffLines = new ObservableCollection<DiffLineViewModel>();
            TargetDiffLines = new ObservableCollection<DiffLineViewModel>();
            PickSourceCommand = new RelayCommand(() => SetConnection(isSource: true));
            PickTargetCommand = new RelayCommand(() => SetConnection(isSource: false));
            ClearSourceCommand = new RelayCommand(() => SourceConnection = null, () => SourceConnection != null);
            ClearTargetCommand = new RelayCommand(() => TargetConnection = null, () => TargetConnection != null);
            SwapConnectionsCommand = new RelayCommand(
                SwapConnections,
                () => !IsBusy && (SourceConnection != null || TargetConnection != null));
            CompareCommand = new AsyncRelayCommand(CompareAsync, () => HasConnections && !IsBusy);
            CancelCompareCommand = new RelayCommand(CancelCompare, () => IsBusy);
            CopyDeploymentScriptCommand = new RelayCommand(CopyDeploymentScript, () => !string.IsNullOrEmpty(DeploymentScript));
            CopySourceDefinitionCommand = new RelayCommand(CopySourceDefinition, () => SelectedDifference != null);
            CopyTargetDefinitionCommand = new RelayCommand(CopyTargetDefinition, () => SelectedDifference != null);
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

        public ICommand CancelCompareCommand { get; }

        public ICommand CopyDeploymentScriptCommand { get; }

        public ICommand CopySourceDefinitionCommand { get; }

        public ICommand CopyTargetDefinitionCommand { get; }

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

                var sourceConnectionString = SourceConnection.FullConnectionString;
                var targetConnectionString = TargetConnection.FullConnectionString;
                var targetDatabase = string.IsNullOrWhiteSpace(TargetConnection.Database) ? "Target" : TargetConnection.Database;

                var payload = await Task.Run(() =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var sourceEndpoint = new SchemaCompareDatabaseEndpoint(sourceConnectionString);
                    var targetEndpoint = new SchemaCompareDatabaseEndpoint(targetConnectionString);
                    var comparison = new SchemaComparison(sourceEndpoint, targetEndpoint);
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
                Status = Differences.Count == 0 ? "No differences were found." : $"Found {Differences.Count} difference(s).";
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

        private void CopyDeploymentScript()
        {
            CopyTextToClipboard(DeploymentScript);
        }

        private void CopySourceDefinition()
        {
            if (SelectedDifference == null)
            {
                return;
            }

            SelectedDifference.EnsureDefinitionsLoaded();
            CopyTextToClipboard(SelectedDifference.SourceDefinition);
        }

        private void CopyTargetDefinition()
        {
            if (SelectedDifference == null)
            {
                return;
            }

            SelectedDifference.EnsureDefinitionsLoaded();
            CopyTextToClipboard(SelectedDifference.TargetDefinition);
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

            difference.EnsureDefinitionsLoaded();

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

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    internal class SchemaDifferenceViewModel
    {
        private readonly SchemaDifference _difference;
        private readonly SchemaComparisonResult _result;
        private string _sourceDefinition;
        private string _targetDefinition;
        private bool _definitionsLoaded;

        public SchemaDifferenceViewModel(SchemaDifference difference, SchemaComparisonResult result)
        {
            _difference = difference;
            _result = result;

            Name = difference.Name?.ToString() ?? string.Empty;
            DifferenceType = difference.DifferenceType.ToString();
            Action = difference.UpdateAction.ToString();

            SourceObject = difference.SourceObject?.Name.ToString() ?? string.Empty;
            TargetObject = difference.TargetObject?.Name.ToString() ?? string.Empty;
        }

        public void EnsureDefinitionsLoaded()
        {
            if (_definitionsLoaded)
            {
                return;
            }

            var sbSource = new StringBuilder();
            GetAllSourceChidren(_difference, _result, sbSource);
            _sourceDefinition = sbSource.ToString();

            var sbTarget = new StringBuilder();
            GetAllTargetChidren(_difference, _result, sbTarget);
            _targetDefinition = sbTarget.ToString();

            _definitionsLoaded = true;
        }

        public void GetAllSourceChidren(SchemaDifference difference, SchemaComparisonResult result, StringBuilder sb) 
        {
            string stringResult = result.GetDiffEntrySourceScript(difference);
            if (!string.IsNullOrWhiteSpace(stringResult))
            {
                sb.AppendLine(stringResult);
                sb.AppendLine("GO");
            }

            foreach (var diff in difference.Children)
            {
                if (diff.Included)
                {
                    GetAllSourceChidren(diff, result, sb);
                }
            }
        }

        public void GetAllTargetChidren(SchemaDifference difference, SchemaComparisonResult result, StringBuilder sb)
        {
            string stringResult = result.GetDiffEntryTargetScript(difference);
            if (!string.IsNullOrWhiteSpace(stringResult))
            {
                sb.AppendLine(stringResult);
                sb.AppendLine("GO");
            }

            foreach (var diff in difference.Children)
            {
                if (diff.Included)
                {
                    GetAllTargetChidren(diff, result, sb);
                }
            }
        }

        public string Name { get; }

        public string DifferenceType { get; }

        public string Action { get; }

        public string SourceObject { get; }

        public string TargetObject { get; }

        public string SourceDefinition => _sourceDefinition;

        public string TargetDefinition => _targetDefinition;
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
