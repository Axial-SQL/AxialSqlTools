using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.VisualStudio.Shell;
using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

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
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                if (value == DBNull.Value) return "(NULL)";
                // Only format numeric columns; numeric-looking text remains a grouping label.
                if (!(parameter is bool numeric) || !numeric || value == null) return value;
                // Group the original digits without rounding high-precision SQL numeric labels.
                string text = System.Convert.ToString(value, CultureInfo.InvariantCulture);
                int integerEnd = text.IndexOf('.');
                if (integerEnd < 0) integerEnd = text.Length;
                int digitsStart = text.StartsWith("-", StringComparison.Ordinal) ||
                    text.StartsWith("+", StringComparison.Ordinal) ? 1 : 0;
                if (integerEnd <= digitsStart || text.IndexOfAny(new[] { 'e', 'E' }) >= 0 ||
                    !text.Substring(digitsStart, integerEnd - digitsStart).All(c => c >= '0' && c <= '9')) return value;
                var formatted = new StringBuilder(text);
                for (int i = integerEnd - 3; i > digitsStart; i -= 3) formatted.Insert(i, ',');
                return formatted.ToString();
            }
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
            var textStyle = new Style(typeof(TextBlock), DataGridTextColumn.DefaultElementStyle);
            textStyle.Setters.Add(new Setter(TextBlock.ForegroundProperty, new Binding("Foreground")
                { RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(DataGridCell), 1) }));
            var numericStyle = new Style(typeof(TextBlock), textStyle);
            numericStyle.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, TextAlignment.Right));
            numericStyle.Setters.Add(new Setter(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Stretch));
            var view = result.Table.DefaultView;
            var totalRow = view[view.Count - 1];
            var valueCellStyle = CreateCellStyle(false, totalRow);
            var groupingCellStyle = CreateCellStyle(true, totalRow);
            int rowColumnCount = Math.Max(1, request.Rows.Length);
            foreach (DataColumn column in result.Table.Columns)
            {
                bool numeric = column.Ordinal >= rowColumnCount ||
                    (column.Ordinal < request.Rows.Length && source.Fields[request.Rows[column.Ordinal]].IsNumeric);
                var gridColumn = new DataGridTextColumn
                {
                    Header = column.Caption,
                    Binding = new Binding("[" + column.ColumnName + "]")
                        { Mode = BindingMode.OneWay, Converter = CellDisplay, ConverterParameter = numeric, TargetNullValue = "(NULL)" },
                    ElementStyle = numeric ? numericStyle : textStyle,
                    CellStyle = column.Ordinal < rowColumnCount ||
                        (request.Columns.Length > 0 && column.Ordinal == result.Table.Columns.Count - 1)
                        ? groupingCellStyle : valueCellStyle,
                    Width = new DataGridLength(150)
                };
                ResultGrid.Columns.Add(gridColumn);
            }
            ResultGrid.FrozenColumnCount = rowColumnCount;
            ResultGrid.ItemsSource = view;
            Status.Text = string.Format("{0:N0} matching rows. {1:N0} pivot rows including grand total. Calculated in {2:N2}s. Select cells and press Ctrl+C to copy.",
                result.MatchedRows, result.Table.Rows.Count, elapsed.Elapsed.TotalSeconds);
        }

        private Style CreateCellStyle(bool grouping, DataRowView totalRow)
        {
            var style = new Style(typeof(DataGridCell), (Style)FindResource(typeof(DataGridCell)));
            style.Setters.Add(new Setter(Control.BackgroundProperty, grouping ? Brushes.WhiteSmoke : Brushes.White));
            style.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.Black));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, Brushes.LightGray));
            // Compare the row object, not its label: a source value can also be 'Grand total'.
            var total = new DataTrigger { Binding = new Binding(), Value = totalRow };
            total.Setters.Add(new Setter(Control.BackgroundProperty, Brushes.WhiteSmoke));
            style.Triggers.Add(total);
            // Keep selected cells distinguishable on both the white and gray backgrounds.
            var selected = new Trigger { Property = DataGridCell.IsSelectedProperty, Value = true };
            selected.Setters.Add(new Setter(Control.BackgroundProperty, SystemColors.HighlightBrush));
            selected.Setters.Add(new Setter(Control.ForegroundProperty, SystemColors.HighlightTextBrush));
            style.Triggers.Add(selected);
            return style;
        }

        private void AggregationChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) => UpdateValueFields();

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
            ResultGrid.Columns.Clear();
            RowFields.ItemsSource = ColumnFields.ItemsSource = FilterField.ItemsSource = ValueField.ItemsSource = null;
            theme.Dispose();
        }
    }
}
