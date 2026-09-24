using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace AxialSqlTools.DataCompare
{
    public partial class DataCompareWindowControl : UserControl, IDisposable
    {
        private const int PageSize = 200;
        private const int PreviewCharacters = 500000;
        private readonly ToolWindowThemeController theme;
        private readonly ObservableCollection<ColumnMapping> mappings = new ObservableCollection<ColumnMapping>();
        public ObservableCollection<TableColumn> TargetColumns { get; } = new ObservableCollection<TableColumn>();
        private TableSchema sourceSchema, targetSchema;
        private ComparisonResult result;
        private CancellationTokenSource operation;
        private long page;
        private bool initialized, settingKey, disposed;
        public DataCompareWindowControl()
        {
            InitializeComponent();
            DataContext = this;
            theme = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));
            MappingGrid.ItemsSource = mappings;
            SourceEndpoint.EndpointChanged += EndpointChanged;
            TargetEndpoint.EndpointChanged += EndpointChanged;
            initialized = true;
        }

        private async void EndpointChanged(object sender, EventArgs e)
        {
            if (!initialized || disposed) return;
            InvalidateResult();
            foreach (var mapping in mappings) mapping.PropertyChanged -= MappingChanged;
            mappings.Clear(); TargetColumns.Clear(); KeyBox.ItemsSource = null;
            sourceSchema = targetSchema = null; CompareButton.IsEnabled = false;
            MappingStatus.Text = "Select a source and target table to load their columns.";
            if (SourceEndpoint.IsBusy || TargetEndpoint.IsBusy || SourceEndpoint.SelectedTable == null || TargetEndpoint.SelectedTable == null) return;
            var source = SourceEndpoint.SelectedTable; var target = TargetEndpoint.SelectedTable;
            string sourceConnection = SourceEndpoint.ConnectionString, targetConnection = TargetEndpoint.ConnectionString;
            await RunAsync(async token =>
            {
                var schemas = await Task.WhenAll(
                    Task.Run(() => SqlTableReader.LoadSchema(sourceConnection, source.Schema, source.Name, token), token),
                    Task.Run(() => SqlTableReader.LoadSchema(targetConnection, target.Schema, target.Name, token), token));
                if (disposed) return;
                sourceSchema = schemas[0]; targetSchema = schemas[1];
                foreach (var column in targetSchema.Columns) TargetColumns.Add(column);
                foreach (var mapping in ComparisonPlan.AutoMap(sourceSchema, targetSchema))
                { mappings.Add(mapping); mapping.PropertyChanged += MappingChanged; }
                var keys = new List<KeyChoice> { new KeyChoice { Label = "Custom key - select columns below" } };
                keys.AddRange(sourceSchema.Keys.Select(k => new KeyChoice { Key = k, Label = k.Name + " (" + string.Join(", ", k.Columns) + ")" }));
                settingKey = true;
                KeyBox.ItemsSource = keys;
                KeyBox.SelectedItem = keys.FirstOrDefault(k => k.Key != null && k.Key.Columns.Count == mappings.Count(m => m.IsKey) &&
                    k.Key.Columns.All(n => mappings.Any(m => m.IsKey && m.Source.Name == n))) ?? keys[0];
                settingKey = false;
                int unsupported = mappings.Count(m => !m.Source.Supported || (m.Target != null && !m.Target.Supported));
                int missing = mappings.Count(m => m.Target == null);
                MappingStatus.Text = mappings.Count + " source columns; " + mappings.Count(m => m.Include) + " included. " +
                    missing + " unmapped; " + unsupported + " unsupported. Computed and rowversion columns are excluded by default.";
                StatusText.Text = "Review the mappings and comparison key, then select Compare.";
            }, "Loading table metadata...");
        }

        private void MappingChanged(object sender, PropertyChangedEventArgs e)
        {
            InvalidateResult();
            if (e.PropertyName == nameof(ColumnMapping.IsKey) && !settingKey && KeyBox.Items.Count > 0)
            { settingKey = true; KeyBox.SelectedIndex = 0; settingKey = false; }
        }

        private void KeyChanged(object sender, SelectionChangedEventArgs e)
        {
            if (settingKey || !(KeyBox.SelectedItem is KeyChoice choice) || choice.Key == null) return;
            settingKey = true;
            foreach (var mapping in mappings) mapping.IsKey = choice.Key.Columns.Contains(mapping.Source.Name);
            settingKey = false;
        }

        private void OptionsChanged(object sender, RoutedEventArgs e) { if (initialized) InvalidateResult(); }

        private ComparisonPlan BuildPlan()
        {
            MappingGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            MappingGrid.CommitEdit(DataGridEditingUnit.Row, true);
            if (!int.TryParse(TimeoutBox.Text, out int timeout) || timeout < 1 || timeout > 86400)
                throw new InvalidOperationException("Enter a query timeout between 1 and 86400 seconds.");
            if (!int.TryParse(DiskLimitBox.Text, out int disk) || disk < 1 || disk > 1024)
                throw new InvalidOperationException("Enter a temporary storage limit between 1 and 1024 GiB.");
            var plan = new ComparisonPlan { Source = sourceSchema, Target = targetSchema,
                Columns = mappings.Where(m => m.Include).Select(m => m.Copy()).ToList(),
                Options = new ComparisonOptions { CommandTimeoutSeconds = timeout, MaxTemporaryBytes = disk * 1024L * 1024 * 1024,
                    UseSnapshotIsolation = SnapshotBox.IsChecked == true, IgnoreCase = IgnoreCaseBox.IsChecked == true, IgnoreTrailingSpaces = IgnoreSpacesBox.IsChecked == true } };
            plan.Validate();
            return plan;
        }

        private async void CompareClick(object sender, RoutedEventArgs e)
        {
            await RunAsync(async token =>
            {
                var plan = BuildPlan();
                InvalidateResult();
                string sourceConnection = SourceEndpoint.ConnectionString, targetConnection = TargetEndpoint.ConnectionString;
                var progress = Progress();
                var comparison = await Task.Run(() => ComparisonEngine.Compare(plan,
                    SqlTableReader.ReadRows(sourceConnection, plan, true, token), SqlTableReader.ReadRows(targetConnection, plan, false, token), token, progress), token);
                if (disposed) { comparison.Dispose(); return; }
                result = comparison; ResultsTab.IsEnabled = true;
                KindBox.ItemsSource = Enum.GetValues(typeof(DifferenceKind)).Cast<DifferenceKind>().Select(k => new KindChoice { Kind = k, Label = KindLabel(k) + " (" + result.Count(k).ToString("N0") + ")" }).ToList();
                KindBox.SelectedIndex = Array.FindIndex(Enum.GetValues(typeof(DifferenceKind)).Cast<DifferenceKind>().ToArray(), k => result.Count(k) > 0);
                if (KindBox.SelectedIndex < 0) KindBox.SelectedIndex = 0;
                MainTabs.SelectedItem = ResultsTab;
                ShowPage();
                StatusText.Text = "Comparison complete. Values represent the rows read during this run.";
            }, "Starting comparison...");
        }

        private IProgress<ComparisonProgress> Progress()
        {
            var currentOperation = operation;
            return new Progress<ComparisonProgress>(value => { if (!disposed && operation == currentOperation) StatusText.Text = value.ToString(); });
        }

        private async Task RunAsync(Func<CancellationToken, Task> action, string status)
        {
            if (disposed || operation != null) return;
            var cancellation = new CancellationTokenSource(); operation = cancellation;
            SetBusy(true); StatusText.Text = status;
            try { await action(cancellation.Token); }
            catch (Exception ex)
            {
                if (!disposed)
                {
                    StatusText.Text = cancellation.IsCancellationRequested ? "Operation cancelled. Compare again before synchronizing." : ex.Message;
                    if (!cancellation.IsCancellationRequested) MessageBox.Show(Window.GetWindow(this), ex.Message, "Table Data Compare", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            finally
            {
                operation = null; cancellation.Dispose();
                if (disposed) { result?.Dispose(); result = null; }
                else SetBusy(false);
            }
        }

        private void SetBusy(bool busy)
        {
            SetupPanel.IsEnabled = !busy; ResultActions.IsEnabled = !busy; RowsGrid.IsEnabled = !busy;
            SyncActions.IsEnabled = !busy; CancelButton.IsEnabled = busy;
            CompareButton.IsEnabled = !busy && sourceSchema != null && targetSchema != null;
            BusyProgress.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
            if (!busy) UpdateSummary();
        }

        private void CancelClick(object sender, RoutedEventArgs e) { operation?.Cancel(); CancelButton.IsEnabled = false; StatusText.Text = "Cancelling..."; }

        private void InvalidateScript() { ScriptBox.Clear(); ScriptTab.IsEnabled = false; }
        private void InvalidateResult()
        {
            if (result != null) { result.Dispose(); result = null; StatusText.Text = "Setup changed. Run the comparison again."; }
            if (RowsGrid != null) RowsGrid.ItemsSource = null;
            if (DetailsGrid != null) DetailsGrid.ItemsSource = null;
            if (ResultsTab != null) ResultsTab.IsEnabled = false;
            if (ScriptBox != null) InvalidateScript();
        }

        private DifferenceKind CurrentKind => (KindBox.SelectedItem as KindChoice)?.Kind ?? DifferenceKind.Different;
        private void KindChanged(object sender, SelectionChangedEventArgs e) { page = 0; ShowPage(); }
        private void PreviousClick(object sender, RoutedEventArgs e) { if (page > 0) page--; ShowPage(); }
        private void NextClick(object sender, RoutedEventArgs e) { page++; ShowPage(); }

        private void ShowPage()
        {
            if (result == null) return;
            long pages = Math.Max(1, (result.Count(CurrentKind) + PageSize - 1) / PageSize);
            page = Math.Max(0, Math.Min(page, pages - 1));
            RowsGrid.ItemsSource = result.Read(CurrentKind, page * PageSize, PageSize).Select(r => new RowView(result, r, () => { InvalidateScript(); UpdateSummary(); })).ToList();
            PageLabel.Text = (page + 1).ToString("N0") + " / " + pages.ToString("N0");
            PreviousButton.IsEnabled = page > 0; NextButton.IsEnabled = page + 1 < pages;
            if (RowsGrid.Items.Count > 0) RowsGrid.SelectedIndex = 0;
            else DetailsGrid.ItemsSource = null;
            UpdateSummary();
        }

        private void RowChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!(RowsGrid.SelectedItem is RowView row) || result == null) { DetailsGrid.ItemsSource = null; return; }
            var difference = result.Read(row.Kind, row.Index, 1).Single();
            DetailsGrid.ItemsSource = result.Plan.Columns.Select((m, i) => new ColumnView {
                Name = m.Source.Name + (m.Source.Name == m.Target.Name ? "" : " -> " + m.Target.Name) + (m.IsKey ? " (key)" : ""),
                Source = CellDisplay(difference.Source, i), Target = CellDisplay(difference.Target, i), Changed = difference.Changed[i],
                SourceValue = difference.Source?[i], TargetValue = difference.Target?[i], HasSource = difference.Source != null, HasTarget = difference.Target != null }).ToList();
        }

        private void DetailDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (!(DetailsGrid.SelectedItem is ColumnView column)) return;
            var panel = new Grid { Margin = new Thickness(12) };
            panel.ColumnDefinitions.Add(new ColumnDefinition()); panel.ColumnDefinitions.Add(new ColumnDefinition());
            for (int i = 0; i < 2; i++)
            {
                var group = new GroupBox { Header = i == 0 ? "Source" : "Target", Margin = new Thickness(4) };
                bool exists = i == 0 ? column.HasSource : column.HasTarget;
                object value = i == 0 ? column.SourceValue : column.TargetValue;
                group.Content = new TextBox { Text = exists ? ValueCodec.Display(value, int.MaxValue) : "<no row>", IsReadOnly = true,
                    FontFamily = new System.Windows.Media.FontFamily("Consolas"), AcceptsReturn = true,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
                Grid.SetColumn(group, i); panel.Children.Add(group);
            }
            var window = new Microsoft.VisualStudio.PlatformUI.DialogWindow { Title = column.Name, Owner = Window.GetWindow(this), Width = 1000, Height = 600, Content = panel,
                WindowStartupLocation = WindowStartupLocation.CenterOwner };
            window.SetResourceReference(Control.BackgroundProperty, "AxialThemeBackgroundBrush");
            window.SetResourceReference(Control.ForegroundProperty, "AxialThemeForegroundBrush");
            window.Resources.MergedDictionaries.Add(new ResourceDictionary { Source = new Uri("/AxialSqlTools;component/Themes/SharedToolWindowTheme.xaml", UriKind.Relative) });
            using (var controller = new ToolWindowThemeController(window, () => ToolWindowThemeResources.ApplySharedTheme(window))) window.ShowDialog();
        }

        private void SelectCategoryClick(object sender, RoutedEventArgs e) { if (result != null) { result.Selection.SetAll(CurrentKind, true); InvalidateScript(); ShowPage(); } }
        private void ClearCategoryClick(object sender, RoutedEventArgs e) { if (result != null) { result.Selection.SetAll(CurrentKind, false); InvalidateScript(); ShowPage(); } }

        private void UpdateSummary()
        {
            if (result == null) return;
            SummaryText.Text = string.Join("   |   ", Enum.GetValues(typeof(DifferenceKind)).Cast<DifferenceKind>().Select(k => KindLabel(k) + ": " + result.Count(k).ToString("N0"))) +
                "   |   Selected: " + result.SelectedCount.ToString("N0");
            GenerateButton.IsEnabled = ExportScriptButton.IsEnabled = ApplyButton.IsEnabled = result.SelectedCount > 0;
        }

        private async void GenerateClick(object sender, RoutedEventArgs e)
        {
            await RunAsync(async token =>
            {
                var writer = new PreviewWriter(PreviewCharacters); var progress = Progress();
                await Task.Run(() => SynchronizationScript.Write(result, writer, token, progress), token);
                if (disposed) return;
                ScriptBox.Text = writer.ToString();
                ScriptNote.Text = writer.Truncated ? "Preview is limited to 500,000 characters. Use Save SQL for the complete script. Do not execute this truncated preview." :
                    "Complete script for " + result.Plan.Target.Server + " / " + result.Plan.Target.Database + " / " + result.Plan.Target.QualifiedName + ". Select and copy, or use Save SQL.";
                if (writer.Truncated) ScriptBox.Text = "-- INCOMPLETE PREVIEW - USE SAVE SQL FOR THE EXECUTABLE SCRIPT\nTHROW 51010, 'This is an incomplete preview. Export the full script.', 1;\n" + ScriptBox.Text;
                ScriptTab.IsEnabled = true; MainTabs.SelectedItem = ScriptTab; StatusText.Text = "Synchronization script generated.";
            }, "Generating synchronization script...");
        }

        private async void ExportScriptClick(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog { Filter = "SQL script (*.sql)|*.sql", FileName = "AxialDataCompare.sql" };
            if (dialog.ShowDialog(Window.GetWindow(this)) != true) return;
            await RunAsync(async token =>
            {
                var progress = Progress();
                await Task.Run(() => WriteFile(dialog.FileName, writer => SynchronizationScript.Write(result, writer, token, progress), token), token);
                StatusText.Text = "Synchronization script saved.";
            }, "Saving synchronization script...");
        }

        private async void ExportCsvClick(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog { Filter = "CSV file (*.csv)|*.csv", FileName = "AxialDataCompare.csv" };
            if (dialog.ShowDialog(Window.GetWindow(this)) != true) return;
            await RunAsync(async token =>
            {
                var progress = Progress();
                await Task.Run(() => WriteFile(dialog.FileName, writer => ComparisonCsv.Write(result, writer, token, progress), token), token);
                StatusText.Text = "Comparison exported. CSV contains every category and full values.";
            }, "Exporting comparison...");
        }

        private async void ApplyClick(object sender, RoutedEventArgs e)
        {
            if (result == null) return;
            var plan = result.Plan;
            string message = "Apply source values to this target?\n\n" + plan.Target.Server + " / " + plan.Target.Database + " / " + plan.Target.QualifiedName +
                "\n\nInserts: " + result.Selection.Count(DifferenceKind.OnlySource, result.Count(DifferenceKind.OnlySource)).ToString("N0") +
                "\nUpdates: " + result.Selection.Count(DifferenceKind.Different, result.Count(DifferenceKind.Different)).ToString("N0") +
                "\nDeletes: " + result.Selection.Count(DifferenceKind.OnlyTarget, result.Count(DifferenceKind.OnlyTarget)).ToString("N0") +
                "\n\nChanges run in one transaction. Constraints and triggers stay enabled. A conflict stops the operation. Use Preview script first if you need to review the SQL.";
            if (MessageBox.Show(Window.GetWindow(this), message, "Apply selected rows", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) != MessageBoxResult.Yes) return;
            await RunAsync(async token =>
            {
                var progress = Progress(); string connection = TargetEndpoint.ConnectionString;
                try
                {
                    long count = await Task.Run(() => SqlSynchronizer.Apply(connection, result, token, progress), token);
                    if (!disposed) StatusText.Text = count.ToString("N0") + " changes committed. Run Compare again to verify the target.";
                }
                finally
                {
                    // A cancellation or connection loss may leave the outcome uncertain. Never reuse the old snapshot.
                    string status = StatusText.Text; InvalidateResult();
                    if (!disposed) { MainTabs.SelectedIndex = 0; StatusText.Text = status; }
                }
            }, "Applying selected rows...");
        }

        private static void WriteFile(string path, Action<TextWriter> write, CancellationToken token)
        {
            string temporary = Path.Combine(Path.GetDirectoryName(path), ".axial-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (var writer = new StreamWriter(temporary, false, new UTF8Encoding(true))) write(writer);
                token.ThrowIfCancellationRequested();
                if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static string CellDisplay(object[] row, int ordinal) => row == null ? "<no row>" : row[ordinal] is string text && text.Length == 0 ? "<empty string>" : ValueCodec.Display(row[ordinal]);
        private static string KindLabel(DifferenceKind kind) => kind == DifferenceKind.OnlySource ? "Only in source" : kind == DifferenceKind.OnlyTarget ? "Only in target" : kind.ToString();

        public void Dispose()
        {
            if (disposed) return;
            disposed = true; operation?.Cancel();
            SourceEndpoint.EndpointChanged -= EndpointChanged; TargetEndpoint.EndpointChanged -= EndpointChanged;
            SourceEndpoint.Dispose(); TargetEndpoint.Dispose();
            foreach (var mapping in mappings) mapping.PropertyChanged -= MappingChanged;
            theme.Dispose();
            if (operation == null) { result?.Dispose(); result = null; }
        }

        public sealed class KeyChoice { public string Label { get; set; } public TableKey Key { get; set; } }
        public sealed class KindChoice { public string Label { get; set; } public DifferenceKind Kind { get; set; } }
        public sealed class ColumnView
        {
            public string Name { get; set; } public string Source { get; set; } public string Target { get; set; } public bool Changed { get; set; }
            public object SourceValue { get; set; } public object TargetValue { get; set; } public bool HasSource { get; set; } public bool HasTarget { get; set; }
        }
        public sealed class RowView : INotifyPropertyChanged
        {
            private readonly ComparisonResult result;
            private readonly Action selectionChanged;
            public DifferenceKind Kind { get; }
            public long Index { get; }
            public RowView(ComparisonResult result, DifferenceRow row, Action selectionChanged)
            {
                this.result = result; this.selectionChanged = selectionChanged; Kind = row.Kind; Index = row.Index;
                Key = string.Join("; ", result.Plan.KeyOrdinals.Select(i => result.Plan.Columns[i].Source.Name + "=" + ValueCodec.Display((row.Source ?? row.Target)[i], 200)));
                Changes = row.Kind == DifferenceKind.Different ? string.Join(", ", result.Plan.Columns.Where((m, i) => row.Changed[i]).Select(m => m.Source.Name)) : KindLabel(row.Kind);
                SourceSummary = Summary(row.Source); TargetSummary = Summary(row.Target);
                // Retain only display strings for this page; full row values are loaded on selection.
            }
            public bool CanSelect => Kind != DifferenceKind.Identical;
            public bool IsSelected
            {
                get => result.Selection.IsSelected(Kind, Index);
                set { result.Selection.Set(Kind, Index, value); PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected))); selectionChanged(); }
            }
            public string Key { get; }
            public string Changes { get; }
            public string SourceSummary { get; }
            public string TargetSummary { get; }
            private string Summary(object[] values) => values == null ? "<no row>" : string.Join(" | ", values.Take(3).Select(v => ValueCodec.Display(v, 120)));
            public event PropertyChangedEventHandler PropertyChanged;
        }
        private sealed class PreviewWriter : TextWriter
        {
            private readonly StringBuilder text = new StringBuilder(); private readonly int limit;
            public bool Truncated { get; private set; }
            public override Encoding Encoding => Encoding.UTF8;
            public PreviewWriter(int limit) { this.limit = limit; }
            public override void Write(char value) { if (text.Length < limit) text.Append(value); else Truncated = true; }
            public override void Write(string value)
            {
                if (value == null) return;
                int count = Math.Min(value.Length, limit - text.Length);
                text.Append(value, 0, count); if (count < value.Length) Truncated = true;
            }
            public override string ToString() => text.ToString();
        }
    }
}
