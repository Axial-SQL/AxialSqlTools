using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.VisualStudio.Shell;
using System;
using System.Data;
using System.Collections.Generic;
using System.ComponentModel;
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
        private PivotRequest appliedRequest;
        private int sortColumn = -1;
        private ListSortDirection? sortDirection;
        private string resultStatus = "";
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

        public PivotGridWindowControl()
        {
            InitializeComponent();
            theme = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));
            Builder.ConfigurationChanged += BuilderChanged;
        }

        internal Task LoadGridAsync(IGridControl grid)
        {
            return RunAsync(async token =>
            {
                snapshot = await PivotGridSnapshot.CaptureAsync(grid, token,
                    (done, total) => Status.Text = string.Format("Copying grid: {0:N0} / {1:N0} rows...", done, total));
                if (disposed) return;
                SourceInfo.Text = string.Format("{0:N0} captured rows  |  {1:N0} columns  |  Snapshot {2:t}. The source query is not re-run.",
                    snapshot.Rows.Length, snapshot.Fields.Length, DateTime.Now);
                Builder.Initialize(snapshot);
                var remembered = PivotLayoutStore.GetLast(snapshot);
                if (remembered != null) Builder.SetRequest(remembered);
                RefreshLayoutNames();
                await BuildAsync(token);
            });
        }

        private void ApplyClicked(object sender, RoutedEventArgs e) =>
            AxialSqlToolsPackage.PackageInstance.JoinableTaskFactory.RunAsync(() => RunAsync(BuildAsync))
                .FileAndForget("AxialSqlTools/PivotGrid/Apply");

        private void ControlSizeChanged(object sender, SizeChangedEventArgs e)
        {
            // Reserve room for the result even when the document is short or docked.
            if (SetupScroll != null) SetupScroll.MaxHeight = Math.Max(72, Math.Min(310, ActualHeight - 330));
        }

        private void ControlKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && Keyboard.Modifiers == ModifierKeys.Control && ApplyButton.IsEnabled)
            {
                e.Handled = true;
                ApplyClicked(sender, e);
            }
        }

        private async Task BuildAsync(CancellationToken token)
        {
            if (snapshot == null) return;
            var request = Builder.GetRequest().Copy();
            PivotEngine.Validate(snapshot, request);
            Builder.IsEnabled = false;
            var source = snapshot;
            // Keep the applied view and its identities intact until the replacement is ready.
            PendingState.Text = displayedResult == null ? "Calculating..." : "Calculating... The grid shows the last applied layout.";
            Status.Text = "Calculating pivot...";
            var elapsed = Stopwatch.StartNew();
            var result = await Task.Run(() => PivotEngine.Build(source, request, token), token);
            token.ThrowIfCancellationRequested();
            if (disposed) return;
            DisplayResult(result, request);
            resultStatus = string.Format("{0:N0} of {1:N0} source rows match. {2:N0} groups. Calculated in {3:N2}s.",
                result.MatchedRows, source.Rows.Length, result.Table.Rows.Count - 1, elapsed.Elapsed.TotalSeconds);
            Status.Text = resultStatus;
            AppliedSummary.Text = "Applied: " + DescribeRequest(request);
            if (!PivotLayoutStore.SaveLast(source, request))
                Status.Text += " Layout could not be remembered: " + PivotLayoutStore.LastError;
        }

        private void DisplayResult(PivotResult result, PivotRequest request)
        {
            var textStyle = CreateTextStyle(false);
            var numericStyle = CreateTextStyle(true);
            var view = result.Table.DefaultView;
            var totalRow = view[view.Count - 1];
            var valueCellStyle = CreateCellStyle(this, false, totalRow);
            var groupingCellStyle = CreateCellStyle(this, true, totalRow);
            var widths = ResultGrid.Columns.GroupBy(c => Convert.ToString(c.Header)).ToDictionary(g => g.Key, g => g.First().Width);
            string previousSort = sortColumn >= 0 && displayedResult != null && sortColumn < displayedResult.Table.Columns.Count
                ? displayedResult.Table.Columns[sortColumn].Caption : null;
            var columns = new List<DataGridTextColumn>();
            foreach (DataColumn column in result.Table.Columns)
            {
                bool numeric = column.Ordinal >= result.RowColumnCount ||
                    (column.Ordinal < request.Rows.Length && snapshot.Fields[request.Rows[column.Ordinal]].IsNumeric);
                columns.Add(new DataGridTextColumn
                {
                    Header = column.Caption,
                    SortMemberPath = column.ColumnName,
                    Binding = new Binding("[" + column.ColumnName + "]")
                        { Mode = BindingMode.OneWay, Converter = CellDisplay, ConverterParameter = numeric, TargetNullValue = "(NULL)" },
                    ElementStyle = numeric ? numericStyle : textStyle,
                    CellStyle = column.Ordinal < result.RowColumnCount || result.IsTotalColumn(column.Ordinal)
                        ? groupingCellStyle : valueCellStyle,
                    Width = widths.TryGetValue(column.Caption, out var width) ? width : new DataGridLength(150)
                });
            }
            // No source rows or prior grid are discarded during aggregation or validation.
            ResultGrid.ItemsSource = TotalGrid.ItemsSource = null;
            ResultGrid.Columns.Clear();
            TotalGrid.Columns.Clear();
            activeResultGrid = null;
            foreach (var gridColumn in columns)
            {
                ResultGrid.Columns.Add(gridColumn);
                var totalColumn = new DataGridTextColumn
                {
                    Header = gridColumn.Header,
                    Binding = gridColumn.Binding,
                    ElementStyle = gridColumn.ElementStyle,
                    CellStyle = groupingCellStyle
                };
                BindingOperations.SetBinding(totalColumn, DataGridColumn.WidthProperty,
                    new Binding(nameof(DataGridColumn.ActualWidth))
                        { Source = gridColumn, Mode = BindingMode.OneWay, Converter = ColumnWidth });
                TotalGrid.Columns.Add(totalColumn);
            }
            ResultGrid.FrozenColumnCount = TotalGrid.FrozenColumnCount = result.RowColumnCount;
            TotalGrid.ItemsSource = new[] { totalRow };
            displayedResult = result;
            appliedRequest = request.Copy();
            sortColumn = previousSort == null ? -1 : result.Table.Columns.Cast<DataColumn>()
                .Where(c => c.Caption == previousSort).Select(c => c.Ordinal).DefaultIfEmpty(-1).First();
            if (sortColumn < 0) sortDirection = null;
            ApplySort();
        }

        private string DescribeRequest(PivotRequest request)
        {
            string Fields(int[] indexes) => indexes.Length == 0 ? "none" : string.Join(" > ", indexes.Select(i => snapshot.Fields[i].Name));
            string filters = request.Filters.Length == 0 ? "none" : string.Join("; ", request.Filters.Select(f =>
                snapshot.Fields[f.Field].Name + " " + f.Operator +
                (f.Operator == PivotFilterOperator.In ? " (" + f.Values.Length + " selected)" :
                f.Operator == PivotFilterOperator.IsNull || f.Operator == PivotFilterOperator.IsNotNull ? "" : " " + f.Text)));
            return "Rows: " + Fields(request.Rows) + "  |  Columns: " + Fields(request.Columns) + "  |  " +
                string.Join(", ", request.Measures.Select(m => m.GetLabel(snapshot))) + "  |  Filters: " + filters;
        }

        private void BuilderChanged(object sender, EventArgs e)
        {
            if (disposed || snapshot == null) return;
            if (operation == null)
            {
                Status.SetResourceReference(TextBlock.ForegroundProperty, "AxialThemeForegroundBrush");
                Status.Text = resultStatus;
            }
            UpdatePendingState();
        }

        private void UpdatePendingState()
        {
            if (disposed || snapshot == null) return;
            string validation = null;
            PivotRequest draft = null;
            try
            {
                draft = Builder.GetRequest();
                PivotEngine.Validate(snapshot, draft);
            }
            catch (Exception ex) { validation = ex.Message; }
            bool pending = !PivotPresentation.RequestsEqual(appliedRequest, draft);
            ApplyButton.IsEnabled = operation == null && validation == null && (pending || displayedResult == null);
            PendingState.SetResourceReference(TextBlock.ForegroundProperty,
                validation == null ? "AxialThemeForegroundBrush" : "AxialThemeStatusErrorBrush");
            PendingState.Text = validation != null
                ? validation + (displayedResult == null ? "" : " The grid still shows the last applied layout.")
                : operation != null ? (displayedResult == null ? "Working..." : "Working... The grid shows the last applied layout.")
                : pending ? (displayedResult == null ? "Ready to apply." : "Changes pending. The grid shows the last applied layout.")
                : "Up to date.";
        }

        private void RefreshLayoutNames()
        {
            string name = LayoutNames.Text;
            LayoutNames.ItemsSource = PivotLayoutStore.GetLayouts(snapshot).Select(l => l.Name).ToArray();
            LayoutNames.Text = name;
        }

        private void ShowLayoutError(string fallback)
        {
            Status.SetResourceReference(TextBlock.ForegroundProperty, "AxialThemeStatusErrorBrush");
            Status.Text = string.IsNullOrEmpty(PivotLayoutStore.LastError) ? fallback : PivotLayoutStore.LastError;
        }

        private void LoadLayoutClicked(object sender, RoutedEventArgs e)
        {
            var layouts = PivotLayoutStore.GetLayouts(snapshot);
            var layout = layouts.FirstOrDefault(l => string.Equals(l.Name, LayoutNames.Text.Trim(), StringComparison.OrdinalIgnoreCase));
            if (layout == null) { ShowLayoutError("Choose a saved layout for these result columns."); return; }
            Builder.SetRequest(layout.Request.Copy());
            SetupExpander.IsExpanded = true;
        }

        private void SaveLayoutClicked(object sender, RoutedEventArgs e)
        {
            try
            {
                var request = Builder.GetRequest();
                PivotEngine.Validate(snapshot, request);
                string name = LayoutNames.Text.Trim();
                var layouts = PivotLayoutStore.GetLayouts(snapshot);
                if (layouts.Any(l => string.Equals(l.Name, name, StringComparison.OrdinalIgnoreCase)) &&
                    MessageBox.Show("Replace saved layout '" + name + "'?", "Pivot Grid", MessageBoxButton.YesNo,
                        MessageBoxImage.Question) != MessageBoxResult.Yes) return;
                if (!PivotLayoutStore.SaveNamed(snapshot, name, request)) { ShowLayoutError("Could not save the layout."); return; }
                RefreshLayoutNames();
                LayoutNames.Text = name;
                Status.SetResourceReference(TextBlock.ForegroundProperty, "AxialThemeForegroundBrush");
                Status.Text = "Layout saved. Apply to update the grid.";
            }
            catch (Exception ex)
            {
                Status.SetResourceReference(TextBlock.ForegroundProperty, "AxialThemeStatusErrorBrush");
                Status.Text = ex.Message;
            }
        }

        private void DeleteLayoutClicked(object sender, RoutedEventArgs e)
        {
            string name = LayoutNames.Text.Trim();
            if (!PivotLayoutStore.GetLayouts(snapshot).Any(l => string.Equals(l.Name, name, StringComparison.OrdinalIgnoreCase)))
            { ShowLayoutError("Choose a saved layout to delete."); return; }
            if (MessageBox.Show("Delete saved layout '" + name + "'?", "Pivot Grid", MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            if (!PivotLayoutStore.DeleteNamed(snapshot, name)) { ShowLayoutError("Could not delete the layout."); return; }
            LayoutNames.Text = "";
            RefreshLayoutNames();
            Status.SetResourceReference(TextBlock.ForegroundProperty, "AxialThemeForegroundBrush");
            Status.Text = "Saved layout deleted.";
        }

        private void SwapAxesClicked(object sender, RoutedEventArgs e)
        {
            var request = Builder.GetRequest();
            var rows = request.Rows;
            request.Rows = request.Columns;
            request.Columns = rows;
            Builder.SetRequest(request);
        }

        private void ResetClicked(object sender, RoutedEventArgs e)
        {
            Builder.SetRequest(new PivotRequest());
            LayoutNames.Text = "";
            SetupExpander.IsExpanded = true;
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

        private void ResultSorting(object sender, DataGridSortingEventArgs e)
        {
            e.Handled = true;
            if (displayedResult == null || operation != null) return;
            int column = ResultGrid.Columns.IndexOf(e.Column);
            if (sortColumn != column || !sortDirection.HasValue)
            {
                sortColumn = column;
                sortDirection = ListSortDirection.Ascending;
            }
            else if (sortDirection == ListSortDirection.Ascending) sortDirection = ListSortDirection.Descending;
            else { sortColumn = -1; sortDirection = null; }
            ApplySort();
        }

        private void ApplySort()
        {
            if (displayedResult == null) return;
            // Only presentation order changes; drill-down retains the original row and column identities.
            ResultGrid.ItemsSource = PivotPresentation.SortRows(displayedResult, snapshot, appliedRequest, sortColumn, sortDirection);
            for (int i = 0; i < ResultGrid.Columns.Count; i++)
                ResultGrid.Columns[i].SortDirection = i == sortColumn ? sortDirection : null;
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
                if (!disposed) Status.Text = resultStatus;
            })).FileAndForget("AxialSqlTools/PivotGrid/DrillDown");
            return true;
        }

        private async Task RunAsync(Func<CancellationToken, Task> action)
        {
            if (disposed || operation != null) return;
            var current = new CancellationTokenSource();
            operation = current;
            SetBusy(true);
            Status.SetResourceReference(TextBlock.ForegroundProperty, "AxialThemeForegroundBrush");
            try { await action(current.Token); }
            catch (OperationCanceledException) { if (!disposed) Status.Text = displayedResult == null ? "Cancelled." : "Cancelled. The last applied result is still shown."; }
            catch (Exception ex)
            {
                if (!disposed)
                {
                    Status.SetResourceReference(TextBlock.ForegroundProperty, "AxialThemeStatusErrorBrush");
                    Status.Text = (ex is OverflowException ? "The total exceeds the decimal range. Cast or scale the values in SQL first." : ex.Message)
                        + (displayedResult == null ? "" : " The last applied result is still shown.");
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
            Builder.IsEnabled = LayoutToolbar.IsEnabled = !busy && snapshot != null;
            CancelButton.IsEnabled = busy;
            DrillButton.IsEnabled = !busy && TryGetDrillCell(out _, out _);
            DrillMenuItem.IsEnabled = TotalDrillMenuItem.IsEnabled = DrillButton.IsEnabled;
            UpdatePendingState();
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
            appliedRequest = null;
            Builder.ConfigurationChanged -= BuilderChanged;
            Builder.Dispose();
            LayoutNames.ItemsSource = null;
            theme.Dispose();
        }
    }
}
