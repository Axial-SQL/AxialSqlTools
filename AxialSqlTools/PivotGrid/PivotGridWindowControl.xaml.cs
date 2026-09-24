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
using System.Windows.Input;
using System.Windows.Media;

namespace AxialSqlTools.PivotGrid
{
    public partial class PivotGridWindowControl : UserControl, IDisposable
    {
        private readonly ToolWindowThemeController theme;
        private PivotSnapshot snapshot;
        private PivotResult displayedResult;
        private PivotDetailsWindow detailsWindow;
        private DataGrid activeResultGrid;
        private bool syncingScroll;
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

        internal static readonly IValueConverter CellDisplay = new CellDisplayConverter();

        private sealed class ColumnWidthConverter : IValueConverter
        {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
                value is double width && !double.IsNaN(width) && !double.IsInfinity(width)
                    ? new DataGridLength(Math.Max(0, width)) : DataGridLength.Auto;
            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
                Binding.DoNothing;
        }

        private static readonly IValueConverter ColumnWidth = new ColumnWidthConverter();

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
            displayedResult = null;
            // Clear previous results so a failed/cancelled Apply cannot look like current statistics.
            ResultGrid.ItemsSource = TotalGrid.ItemsSource = null;
            ResultGrid.Columns.Clear();
            TotalGrid.Columns.Clear();
            activeResultGrid = null;
            Status.Text = "Calculating pivot...";
            var elapsed = Stopwatch.StartNew();
            var result = await Task.Run(() => PivotEngine.Build(source, request, token), token);
            token.ThrowIfCancellationRequested();
            if (disposed) return;
            var textStyle = CreateTextStyle(false);
            var numericStyle = CreateTextStyle(true);
            var view = result.Table.DefaultView;
            var totalRow = view[view.Count - 1];
            var valueCellStyle = CreateCellStyle(this, false, totalRow);
            var groupingCellStyle = CreateCellStyle(this, true, totalRow);
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
                var totalColumn = new DataGridTextColumn
                {
                    Header = gridColumn.Header,
                    Binding = gridColumn.Binding,
                    ElementStyle = gridColumn.ElementStyle,
                    CellStyle = groupingCellStyle
                };
                // Follow rendered widths, including drag resizing and automatic sizing.
                BindingOperations.SetBinding(totalColumn, DataGridColumn.WidthProperty,
                    new Binding(nameof(DataGridColumn.ActualWidth))
                        { Source = gridColumn, Mode = BindingMode.OneWay, Converter = ColumnWidth });
                TotalGrid.Columns.Add(totalColumn);
            }
            ResultGrid.FrozenColumnCount = TotalGrid.FrozenColumnCount = rowColumnCount;
            TotalGrid.ItemsSource = new[] { totalRow };
            displayedResult = result;
            ApplyValueSort();
            Status.Text = string.Format("{0:N0} matching rows. {1:N0} pivot rows including grand total. Calculated in {2:N2}s. Double-click a value or total to see underlying rows. Ctrl+C copies selected cells.",
                result.MatchedRows, result.Table.Rows.Count, elapsed.Elapsed.TotalSeconds);
        }

        private static ScrollViewer GetScrollViewer(DataGrid grid)
        {
            grid.ApplyTemplate();
            return grid.Template?.FindName("DG_ScrollViewer", grid) as ScrollViewer;
        }

        private void ResultScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (syncingScroll || disposed || ResultGrid == null || TotalGrid == null) return;
            var sourceGrid = sender as DataGrid;
            if (sourceGrid == null) return;
            var source = GetScrollViewer(sourceGrid);
            // Ignore scroll events from child controls.
            if (source == null || e.OriginalSource != source) return;
            var target = GetScrollViewer(sourceGrid == ResultGrid ? TotalGrid : ResultGrid);
            if (target == null) return;
            syncingScroll = true;
            try { target.ScrollToHorizontalOffset(source.HorizontalOffset); }
            finally { syncingScroll = false; }
        }

        private void ResultColumnReordered(object sender, DataGridColumnEventArgs e)
        {
            if (TotalGrid.Columns.Count != ResultGrid.Columns.Count) return;
            foreach (var column in ResultGrid.Columns.OrderBy(c => c.DisplayIndex))
                TotalGrid.Columns[ResultGrid.Columns.IndexOf(column)].DisplayIndex = column.DisplayIndex;
        }

        private void ValueSortChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (displayedResult != null) ApplyValueSort();
        }

        private void ApplyValueSort()
        {
            if (displayedResult == null) return;
            // Reorder only the presentation. Keep DataRows and engine row-key indexes intact
            // so drill-down still resolves the selected group after sorting.
            var view = displayedResult.Table.DefaultView;
            int valueColumn = displayedResult.Table.Columns.Count - 1;
            var rows = view.Cast<DataRowView>().Take(view.Count - 1);
            Func<DataRowView, decimal?> value = row => row[valueColumn] == DBNull.Value
                ? (decimal?)null : Convert.ToDecimal(row[valueColumn], CultureInfo.InvariantCulture);
            if (ValueSort.SelectedIndex == 1) rows = rows.OrderBy(value);
            else if (ValueSort.SelectedIndex == 2) rows = rows.OrderByDescending(value);
            ResultGrid.ItemsSource = rows.ToArray();
        }

        internal static Style CreateTextStyle(bool numeric)
        {
            var style = new Style(typeof(TextBlock), DataGridTextColumn.DefaultElementStyle);
            style.Setters.Add(new Setter(TextBlock.ForegroundProperty, new Binding("Foreground")
                { RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(DataGridCell), 1) }));
            if (numeric)
            {
                style.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, TextAlignment.Right));
                style.Setters.Add(new Setter(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Stretch));
            }
            return style;
        }

        internal static Style CreateCellStyle(FrameworkElement owner, bool grouping, DataRowView totalRow)
        {
            var style = new Style(typeof(DataGridCell), (Style)owner.FindResource(typeof(DataGridCell)));
            style.Setters.Add(new Setter(Control.BackgroundProperty, new DynamicResourceExtension(grouping ? "AxialThemeGridHeaderBackgroundBrush" : "AxialThemeBackgroundBrush")));
            style.Setters.Add(new Setter(Control.ForegroundProperty, new DynamicResourceExtension("AxialThemeForegroundBrush")));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, new DynamicResourceExtension("AxialThemeBorderBrush")));
            // Compare the row object, not its label: a source value can also be 'Grand total'.
            var total = new DataTrigger { Binding = new Binding(), Value = totalRow };
            total.Setters.Add(new Setter(Control.BackgroundProperty, new DynamicResourceExtension("AxialThemeGridHeaderBackgroundBrush")));
            style.Triggers.Add(total);
            // Keep selected cells distinguishable on data, grouping and total backgrounds.
            var selected = new Trigger { Property = DataGridCell.IsSelectedProperty, Value = true };
            selected.Setters.Add(new Setter(Control.BackgroundProperty, new DynamicResourceExtension("AxialThemeGridSelectionBrush")));
            selected.Setters.Add(new Setter(Control.ForegroundProperty, new DynamicResourceExtension("AxialThemeGridSelectionTextBrush")));
            style.Triggers.Add(selected);
            return style;
        }

        private bool TryGetDrillCell(out int row, out int column)
        {
            row = column = -1;
            if (disposed || operation != null || displayedResult == null) return false;
            var grid = activeResultGrid ?? ResultGrid;
            var cell = grid.CurrentCell;
            var item = cell.Item as DataRowView;
            if (item == null || item.Row.Table != displayedResult.Table || cell.Column == null) return false;
            row = displayedResult.Table.Rows.IndexOf(item.Row);
            // Column collection order is stable even when the user changes DisplayIndex.
            column = grid.Columns.IndexOf(cell.Column);
            return displayedResult.CanDrillDown(row, column);
        }

        private static T FindAncestor<T>(DependencyObject source) where T : DependencyObject
        {
            while (source != null)
            {
                if (source is T found) return found;
                source = source is Visual ? VisualTreeHelper.GetParent(source) : (source as FrameworkContentElement)?.Parent;
            }
            return null;
        }

        private bool SelectClickedCell(object originalSource)
        {
            var cell = FindAncestor<DataGridCell>(originalSource as DependencyObject);
            var grid = FindAncestor<DataGrid>(cell);
            if (cell == null || (grid != ResultGrid && grid != TotalGrid)) return false;
            activeResultGrid = grid;
            grid.CurrentCell = new DataGridCellInfo(cell.DataContext, cell.Column);
            return true;
        }

        private void ResultDoubleClicked(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && SelectClickedCell(e.OriginalSource) && StartDrillDown()) e.Handled = true;
        }

        private void ResultRightClicked(object sender, MouseButtonEventArgs e)
        {
            // Do not let a click on a header/empty space reuse a previously selected value.
            activeResultGrid = (DataGrid)sender;
            if (!SelectClickedCell(e.OriginalSource)) activeResultGrid.CurrentCell = new DataGridCellInfo();
        }

        private void ResultKeyDown(object sender, KeyEventArgs e)
        {
            activeResultGrid = (DataGrid)sender;
            if (e.Key == Key.Enter && Keyboard.Modifiers == ModifierKeys.None && StartDrillDown()) e.Handled = true;
        }

        private void ResultCurrentCellChanged(object sender, EventArgs e)
        {
            activeResultGrid = (DataGrid)sender;
            if (DrillButton != null) DrillButton.IsEnabled = TryGetDrillCell(out _, out _);
        }

        private void ResultContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            activeResultGrid = (DataGrid)sender;
            DrillMenuItem.IsEnabled = TotalDrillMenuItem.IsEnabled = TryGetDrillCell(out _, out _);
        }

        private void DrillClicked(object sender, RoutedEventArgs e) => StartDrillDown();

        private bool StartDrillDown()
        {
            if (!TryGetDrillCell(out int row, out int column)) return false;
            var result = displayedResult;
            AxialSqlToolsPackage.PackageInstance.JoinableTaskFactory.RunAsync(() => RunAsync(async token =>
            {
                Status.Text = "Finding underlying rows...";
                var found = await Task.Run(() => result.GetUnderlyingRows(row, column, token), token);
                token.ThrowIfCancellationRequested();
                if (disposed) return;
                using (var dialog = new PivotDetailsWindow(found, snapshot.Fields))
                {
                    detailsWindow = dialog;
                    try
                    {
                        // The shell API sets the SSMS owner and enters proper modal state.
                        dialog.ShowModal();
                    }
                    finally { detailsWindow = null; }
                }
                if (!disposed) Status.Text = "Details closed. Double-click another value or total to inspect its source rows.";
            })).FileAndForget("AxialSqlTools/PivotGrid/DrillDown");
            return true;
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
            DrillButton.IsEnabled = !busy && TryGetDrillCell(out _, out _);
            DrillMenuItem.IsEnabled = TotalDrillMenuItem.IsEnabled = DrillButton.IsEnabled;
        }

        private void CancelClicked(object sender, RoutedEventArgs e) => operation?.Cancel();

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            operation?.Cancel();
            detailsWindow?.Close();
            snapshot = null;
            displayedResult = null;
            ResultGrid.ItemsSource = TotalGrid.ItemsSource = null;
            ResultGrid.Columns.Clear();
            TotalGrid.Columns.Clear();
            activeResultGrid = null;
            RowFields.ItemsSource = ColumnFields.ItemsSource = FilterField.ItemsSource = ValueField.ItemsSource = null;
            theme.Dispose();
        }
    }
}
