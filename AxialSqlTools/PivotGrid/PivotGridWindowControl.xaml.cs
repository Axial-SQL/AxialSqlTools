using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.VisualStudio.Shell;
using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace AxialSqlTools.PivotGrid
{
    public partial class PivotGridWindowControl : UserControl, IDisposable
    {
        private readonly ToolWindowThemeController theme;
        private PivotSnapshot snapshot;
        private CancellationTokenSource operation;
        private bool disposed;

        private sealed class CellDisplayConverter : IValueConverter
        {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
                value == DBNull.Value ? "(NULL)" : value;
            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
        }

        private static readonly IValueConverter CellDisplay = new CellDisplayConverter();

        private sealed class AggregationChoice
        {
            public string Label { get; set; }
            public PivotAggregation Value { get; set; }
        }

        public PivotGridWindowControl()
        {
            InitializeComponent();
            theme = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));
            AggregationBox.ItemsSource = new[]
            {
                new AggregationChoice { Label = "Count rows", Value = PivotAggregation.CountRows },
                new AggregationChoice { Label = "Count non-null values", Value = PivotAggregation.CountValues },
                new AggregationChoice { Label = "Distinct count", Value = PivotAggregation.DistinctCount },
                new AggregationChoice { Label = "Sum", Value = PivotAggregation.Sum },
                new AggregationChoice { Label = "Average", Value = PivotAggregation.Average },
                new AggregationChoice { Label = "Minimum", Value = PivotAggregation.Minimum },
                new AggregationChoice { Label = "Maximum", Value = PivotAggregation.Maximum }
            };
            AggregationBox.SelectedIndex = 0;
        }

        internal Task LoadGridAsync(IGridControl grid)
        {
            return RunAsync(async token =>
            {
                snapshot = await PivotGridSnapshot.CaptureAsync(grid, token,
                    (done, total) => Status.Text = string.Format("Copying grid: {0:N0} / {1:N0} rows...", done, total));
                if (disposed) return;
                SourceInfo.Text = string.Format("{0:N0} rows, {1:N0} columns. Choose grouping fields, then Apply. The source query is not re-run.",
                    snapshot.Rows.Length, snapshot.Fields.Length);
                RowFields.ItemsSource = snapshot.Fields;
                ColumnFields.ItemsSource = snapshot.Fields;
                FilterField.ItemsSource = new[] { new PivotField(-1, "(No filter)", typeof(string)) }.Concat(snapshot.Fields).ToArray();
                FilterField.SelectedIndex = 0;
                RowFields.SelectedItem = snapshot.Fields.FirstOrDefault(f => !f.IsNumeric) ?? snapshot.Fields[0];
                UpdateValueFields();
                await BuildAsync(token);
            });
        }

        private void ApplyClicked(object sender, RoutedEventArgs e) =>
            AxialSqlToolsPackage.PackageInstance.JoinableTaskFactory.RunAsync(() => RunAsync(BuildAsync))
                .FileAndForget("AxialSqlTools/PivotGrid/Apply");

        private async Task BuildAsync(CancellationToken token)
        {
            if (snapshot == null) return;
            var request = new PivotRequest
            {
                Rows = RowFields.SelectedItems.Cast<PivotField>().OrderBy(f => f.Index).Select(f => f.Index).ToArray(),
                Columns = ColumnFields.SelectedItems.Cast<PivotField>().OrderBy(f => f.Index).Select(f => f.Index).ToArray(),
                Value = (ValueField.SelectedItem as PivotField)?.Index ?? -1,
                Aggregation = ((AggregationChoice)AggregationBox.SelectedItem).Value,
                NullTextIsNull = NullTextIsNull.IsChecked == true,
                FilterField = (FilterField.SelectedItem as PivotField)?.Index ?? -1,
                FilterText = FilterText.Text
            };
            var source = snapshot;
            // Clear previous results so a failed/cancelled Apply cannot look like current statistics.
            ResultGrid.ItemsSource = null;
            ResultGrid.Columns.Clear();
            Status.Text = "Calculating pivot...";
            var elapsed = Stopwatch.StartNew();
            var result = await Task.Run(() => PivotEngine.Build(source, request, token), token);
            token.ThrowIfCancellationRequested();
            if (disposed) return;
            foreach (DataColumn column in result.Table.Columns)
                ResultGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = column.Caption,
                    Binding = new Binding("[" + column.ColumnName + "]")
                        { Mode = BindingMode.OneWay, Converter = CellDisplay, TargetNullValue = "(NULL)" },
                    Width = new DataGridLength(150)
                });
            ResultGrid.FrozenColumnCount = Math.Max(1, request.Rows.Length);
            ResultGrid.ItemsSource = result.Table.DefaultView;
            Status.Text = string.Format("{0:N0} matching rows. {1:N0} pivot rows including grand total. Calculated in {2:N2}s. Select cells and press Ctrl+C to copy.",
                result.MatchedRows, result.Table.Rows.Count, elapsed.Elapsed.TotalSeconds);
        }

        private void AggregationChanged(object sender, SelectionChangedEventArgs e) => UpdateValueFields();

        private void UpdateValueFields()
        {
            if (snapshot == null || AggregationBox.SelectedItem == null) return;
            var aggregation = ((AggregationChoice)AggregationBox.SelectedItem).Value;
            var previous = ValueField.SelectedItem as PivotField;
            var fields = snapshot.Fields.Where(f => !PivotEngine.RequiresNumber(aggregation) || f.IsNumeric).ToArray();
            ValueField.ItemsSource = fields;
            ValueField.SelectedItem = fields.Contains(previous) ? previous : fields.FirstOrDefault();
            ValueField.IsEnabled = aggregation != PivotAggregation.CountRows;
        }

        private async Task RunAsync(Func<CancellationToken, Task> action)
        {
            if (disposed || operation != null) return;
            var current = new CancellationTokenSource();
            operation = current;
            SetBusy(true);
            Status.SetResourceReference(TextBlock.ForegroundProperty, "AxialThemeForegroundBrush");
            try { await action(current.Token); }
            catch (OperationCanceledException) { if (!disposed) Status.Text = "Cancelled. No partial results are shown."; }
            catch (Exception ex)
            {
                if (!disposed)
                {
                    Status.SetResourceReference(TextBlock.ForegroundProperty, "AxialThemeStatusErrorBrush");
                    Status.Text = ex is OverflowException ? "The total exceeds the decimal range. Cast or scale the values in SQL first." : ex.Message;
                }
            }
            finally
            {
                operation = null;
                current.Dispose();
                if (!disposed)
                {
                    SetBusy(false);
                    if (snapshot == null) SourceInfo.Text = "No snapshot was captured. Open Pivot Grid again from a completed result grid.";
                }
            }
        }

        private void SetBusy(bool busy)
        {
            Options.IsEnabled = FilterField.IsEnabled = FilterText.IsEnabled = ApplyButton.IsEnabled = !busy && snapshot != null;
            CancelButton.IsEnabled = busy;
        }

        private void CancelClicked(object sender, RoutedEventArgs e) => operation?.Cancel();

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            operation?.Cancel();
            snapshot = null;
            ResultGrid.ItemsSource = null;
            RowFields.ItemsSource = ColumnFields.ItemsSource = FilterField.ItemsSource = ValueField.ItemsSource = null;
            theme.Dispose();
        }
    }
}
