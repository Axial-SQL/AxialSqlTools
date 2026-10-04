using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace AxialSqlTools.PivotGrid
{
    public partial class PivotBuilderControl : UserControl, IDisposable
    {
        private readonly ToolWindowThemeController theme;
        private PivotSnapshot snapshot;
        private PivotFilterWindow filterWindow;
        private readonly ObservableCollection<PivotField> rowFields = new ObservableCollection<PivotField>();
        private readonly ObservableCollection<PivotField> columnFields = new ObservableCollection<PivotField>();
        private readonly List<PivotMeasure> measures = new List<PivotMeasure>();
        private readonly List<PivotFilter> filters = new List<PivotFilter>();
        private bool updating;
        private bool disposed;
        private Point dragStart;
        private ListBox dragSource;
        private PivotField dragField;

        internal event EventHandler ConfigurationChanged;

        private sealed class CatalogField
        {
            public PivotField Field { get; set; }
            public string Label => Field.Label;
            public string Detail { get; set; }
            public string Description => Label + " - " + Detail;
        }

        private sealed class FieldDrag
        {
            public PivotBuilderControl Owner { get; set; }
            public PivotField Field { get; set; }
            public ListBox Source { get; set; }
        }

        private sealed class AggregationChoice
        {
            public PivotAggregation Value { get; set; }
            public string Label { get; set; }
        }

        private static readonly AggregationChoice[] Aggregations =
        {
            new AggregationChoice { Value = PivotAggregation.CountRows, Label = "Count rows" },
            new AggregationChoice { Value = PivotAggregation.CountValues, Label = "Count non-null values" },
            new AggregationChoice { Value = PivotAggregation.DistinctCount, Label = "Distinct count" },
            new AggregationChoice { Value = PivotAggregation.Sum, Label = "Sum" },
            new AggregationChoice { Value = PivotAggregation.Average, Label = "Average" },
            new AggregationChoice { Value = PivotAggregation.Minimum, Label = "Minimum" },
            new AggregationChoice { Value = PivotAggregation.Maximum, Label = "Maximum" }
        };

        public PivotBuilderControl()
        {
            InitializeComponent();
            theme = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));
            RowAxis.ItemsSource = rowFields;
            ColumnAxis.ItemsSource = columnFields;
            IsEnabled = false;
        }

        internal void Initialize(PivotSnapshot source)
        {
            if (disposed) throw new ObjectDisposedException(nameof(PivotBuilderControl));
            snapshot = source ?? throw new ArgumentNullException(nameof(source));
            IsEnabled = true;
            SetRequest(new PivotRequest());
        }

        internal PivotRequest GetRequest()
        {
            return new PivotRequest
            {
                Rows = rowFields.Select(f => f.Index).ToArray(),
                Columns = columnFields.Select(f => f.Index).ToArray(),
                Measures = measures.Select(m => m.Copy()).ToArray(),
                Filters = filters.Select(f => f.Copy()).ToArray(),
                NullTextIsNull = NullTextIsNull.IsChecked == true
            };
        }

        internal void SetRequest(PivotRequest request)
        {
            if (disposed) throw new ObjectDisposedException(nameof(PivotBuilderControl));
            if (snapshot == null) throw new InvalidOperationException("Initialize the builder before restoring a layout.");
            if (request == null) throw new ArgumentNullException(nameof(request));
            updating = true;
            try
            {
                var known = snapshot.Fields.ToDictionary(f => f.Index);
                var assigned = new HashSet<int>();
                rowFields.Clear();
                columnFields.Clear();
                foreach (int index in request.Rows ?? new int[0])
                    if (known.TryGetValue(index, out var field) && assigned.Add(index)) rowFields.Add(field);
                foreach (int index in request.Columns ?? new int[0])
                    if (known.TryGetValue(index, out var field) && assigned.Add(index)) columnFields.Add(field);
                measures.Clear();
                foreach (var source in request.Measures ?? new PivotMeasure[0])
                {
                    if (source == null || !Enum.IsDefined(typeof(PivotAggregation), source.Aggregation)) continue;
                    if (source.Aggregation == PivotAggregation.CountRows)
                        measures.Add(new PivotMeasure { Aggregation = PivotAggregation.CountRows, Value = -1 });
                    else if (known.TryGetValue(source.Value, out var field) &&
                             (!PivotEngine.RequiresNumber(source.Aggregation) || field.IsNumeric)) measures.Add(source.Copy());
                }
                if (measures.Count == 0) measures.Add(new PivotMeasure { Aggregation = PivotAggregation.CountRows, Value = -1 });
                filters.Clear();
                foreach (var filter in request.Filters ?? new PivotFilter[0])
                    if (filter != null && known.ContainsKey(filter.Field) && Enum.IsDefined(typeof(PivotFilterOperator), filter.Operator))
                        filters.Add(filter.Copy());
                NullTextIsNull.IsChecked = request.NullTextIsNull;
                BuilderMessage.Text = "Drag fields or use the buttons. Filters use AND.";
                RenderMeasures();
                RenderFilters();
                RefreshCatalog();
                UpdateAxisButtons();
            }
            finally { updating = false; }
            ConfigurationChanged?.Invoke(this, EventArgs.Empty);
        }

        private void Changed(string message = null)
        {
            if (disposed || updating || snapshot == null) return;
            if (message != null) BuilderMessage.Text = message;
            RefreshCatalog();
            UpdateAxisButtons();
            ConfigurationChanged?.Invoke(this, EventArgs.Empty);
        }

        private void RefreshCatalog()
        {
            if (snapshot == null) return;
            int selected = (FieldCatalog.SelectedItem as CatalogField)?.Field.Index ?? -1;
            string search = FieldSearch.Text ?? "";
            var items = snapshot.Fields.Where(f => f.Label.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
                .Select(f =>
                {
                    var uses = new List<string>();
                    if (rowFields.Contains(f)) uses.Add("Rows");
                    if (columnFields.Contains(f)) uses.Add("Columns");
                    if (measures.Any(m => m.Value == f.Index && m.Aggregation != PivotAggregation.CountRows)) uses.Add("Values");
                    return new CatalogField
                    {
                        Field = f,
                        Detail = (f.DataType?.Name ?? (f.IsNumeric ? "Number" : "Text")) +
                            (uses.Count == 0 ? "" : " | " + string.Join(", ", uses))
                    };
                }).ToArray();
            FieldCatalog.ItemsSource = items;
            FieldCatalog.SelectedItem = items.FirstOrDefault(f => f.Field.Index == selected) ?? items.FirstOrDefault();
            UpdateCatalogButtons();
        }

        private void FieldSearchChanged(object sender, TextChangedEventArgs e) => RefreshCatalog();
        private void CatalogSelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateCatalogButtons();

        private void UpdateCatalogButtons()
        {
            if (AddRowsButton == null) return;
            bool selected = FieldCatalog.SelectedItem is CatalogField;
            AddRowsButton.IsEnabled = AddColumnsButton.IsEnabled = AddValuesButton.IsEnabled = selected;
        }

        private void AddRowsClicked(object sender, RoutedEventArgs e) => AssignSelected(RowAxis);
        private void AddColumnsClicked(object sender, RoutedEventArgs e) => AssignSelected(ColumnAxis);
        private void AddValuesClicked(object sender, RoutedEventArgs e)
        {
            if (FieldCatalog.SelectedItem is CatalogField field) AddMeasure(field.Field);
        }

        private void CatalogDoubleClicked(object sender, MouseButtonEventArgs e)
        {
            if (FindParent<ListBoxItem>(e.OriginalSource as DependencyObject) == null) return;
            AssignSelected(RowAxis);
            e.Handled = true;
        }

        private void CatalogKeyDown(object sender, KeyEventArgs e)
        {
            Key key = e.Key == Key.System ? e.SystemKey : e.Key;
            if (Keyboard.Modifiers == ModifierKeys.Alt && (key == Key.R || key == Key.C || key == Key.V))
            {
                if (key == Key.V) AddValuesClicked(sender, e);
                else AssignSelected(key == Key.R ? RowAxis : ColumnAxis);
                e.Handled = true;
            }
            else if (key == Key.Enter && Keyboard.Modifiers == ModifierKeys.None)
            {
                AssignSelected(RowAxis);
                e.Handled = true;
            }
        }

        private void AssignSelected(ListBox target)
        {
            if (FieldCatalog.SelectedItem is CatalogField field) AssignField(target, field.Field, -1);
        }

        private ObservableCollection<PivotField> FieldsFor(ListBox axis) => axis == RowAxis ? rowFields : columnFields;
        private string AxisName(ListBox axis) => axis == RowAxis ? "Rows" : "Columns";

        private void AssignField(ListBox target, PivotField field, int insertionIndex)
        {
            var fields = FieldsFor(target);
            var other = target == RowAxis ? columnFields : rowFields;
            int previous = fields.IndexOf(field);
            if (previous >= 0 && insertionIndex < 0)
            {
                target.SelectedItem = field;
                BuilderMessage.Text = field.Name + " is already in " + AxisName(target) + ".";
                return;
            }
            if (previous < 0 && !other.Contains(field) && rowFields.Count + columnFields.Count >= 6)
            {
                BuilderMessage.Text = "Choose up to six grouping fields in total. Remove a grouping field before adding another.";
                return;
            }
            bool moved = other.Remove(field);
            if (previous >= 0)
            {
                fields.RemoveAt(previous);
                if (insertionIndex > previous) insertionIndex--;
            }
            int index = insertionIndex < 0 ? fields.Count : Math.Min(insertionIndex, fields.Count);
            fields.Insert(index, field);
            target.SelectedItem = field;
            target.ScrollIntoView(field);
            Changed(moved ? "Moved " + field.Name + " to " + AxisName(target) + ". A field can be on only one grouping axis."
                : (previous >= 0 ? "Reordered " : "Added ") + field.Name + " in " + AxisName(target) + ".");
        }

        private void AxisSelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateAxisButtons();

        private void UpdateAxisButtons()
        {
            if (RowUpButton == null || ColumnUpButton == null) return;
            RowUpButton.IsEnabled = RowAxis.SelectedIndex > 0;
            RowDownButton.IsEnabled = RowAxis.SelectedIndex >= 0 && RowAxis.SelectedIndex < rowFields.Count - 1;
            RowRemoveButton.IsEnabled = RowAxis.SelectedItem != null;
            ColumnUpButton.IsEnabled = ColumnAxis.SelectedIndex > 0;
            ColumnDownButton.IsEnabled = ColumnAxis.SelectedIndex >= 0 && ColumnAxis.SelectedIndex < columnFields.Count - 1;
            ColumnRemoveButton.IsEnabled = ColumnAxis.SelectedItem != null;
            EmptyRows.Visibility = rowFields.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            EmptyColumns.Visibility = columnFields.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private ListBox TaggedAxis(object sender) => (sender as FrameworkElement)?.Tag as string == "Rows" ? RowAxis : ColumnAxis;
        private void MoveUpClicked(object sender, RoutedEventArgs e) => MoveAxis(TaggedAxis(sender), -1);
        private void MoveDownClicked(object sender, RoutedEventArgs e) => MoveAxis(TaggedAxis(sender), 1);
        private void RemoveAxisClicked(object sender, RoutedEventArgs e) => RemoveAxis(TaggedAxis(sender));

        private void MoveAxis(ListBox axis, int offset)
        {
            var fields = FieldsFor(axis);
            int current = axis.SelectedIndex;
            int next = current + offset;
            if (current < 0 || next < 0 || next >= fields.Count) return;
            var field = fields[current];
            fields.Move(current, next);
            axis.SelectedItem = field;
            axis.ScrollIntoView(field);
            Changed("Reordered " + field.Name + " in " + AxisName(axis) + ".");
        }

        private void RemoveAxis(ListBox axis)
        {
            if (!(axis.SelectedItem is PivotField field)) return;
            RemoveAxisField(axis, field);
        }

        private void RemoveAxisField(ListBox axis, PivotField field)
        {
            var fields = FieldsFor(axis);
            int index = fields.IndexOf(field);
            if (!fields.Remove(field)) return;
            axis.SelectedIndex = fields.Count == 0 ? -1 : Math.Min(index, fields.Count - 1);
            Changed("Removed " + field.Name + " from " + AxisName(axis) + ".");
        }

        private void AxisItemRemoveClicked(object sender, RoutedEventArgs e)
        {
            var button = (Button)sender;
            var axis = FindParent<ListBox>(button);
            if (axis != null && button.Tag is PivotField field) RemoveAxisField(axis, field);
            e.Handled = true;
        }

        private void AxisKeyDown(object sender, KeyEventArgs e)
        {
            var axis = (ListBox)sender;
            Key key = e.Key == Key.System ? e.SystemKey : e.Key;
            if (key == Key.Delete && Keyboard.Modifiers == ModifierKeys.None)
            {
                RemoveAxis(axis);
                e.Handled = true;
            }
            else if (Keyboard.Modifiers == ModifierKeys.Alt && (key == Key.Up || key == Key.Down))
            {
                MoveAxis(axis, key == Key.Up ? -1 : 1);
                e.Handled = true;
            }
        }

        private static T FindParent<T>(DependencyObject value) where T : DependencyObject
        {
            while (value != null)
            {
                if (value is T found) return found;
                value = value is Visual ? VisualTreeHelper.GetParent(value) : (value as FrameworkContentElement)?.Parent;
            }
            return null;
        }

        private void FieldMouseDown(object sender, MouseButtonEventArgs e)
        {
            dragSource = null;
            dragField = null;
            if (FindParent<ButtonBase>(e.OriginalSource as DependencyObject) != null) return;
            var item = FindParent<ListBoxItem>(e.OriginalSource as DependencyObject);
            if (item == null) return;
            dragField = item.DataContext is CatalogField catalog ? catalog.Field : item.DataContext as PivotField;
            dragSource = sender as ListBox;
            dragStart = e.GetPosition(this);
        }

        private void FieldMouseMove(object sender, MouseEventArgs e)
        {
            if (dragSource == null || dragField == null || e.LeftButton != MouseButtonState.Pressed) return;
            var point = e.GetPosition(this);
            if (Math.Abs(point.X - dragStart.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(point.Y - dragStart.Y) < SystemParameters.MinimumVerticalDragDistance) return;
            var source = dragSource;
            var payload = new FieldDrag { Owner = this, Field = dragField, Source = source };
            dragSource = null;
            dragField = null;
            DragDrop.DoDragDrop(source, new DataObject(typeof(FieldDrag), payload), DragDropEffects.Copy | DragDropEffects.Move);
            e.Handled = true;
        }

        private FieldDrag GetDrag(DragEventArgs e)
        {
            var value = e.Data.GetDataPresent(typeof(FieldDrag)) ? e.Data.GetData(typeof(FieldDrag)) as FieldDrag : null;
            return value?.Owner == this ? value : null;
        }

        private void AxisDragOver(object sender, DragEventArgs e)
        {
            var drag = GetDrag(e);
            e.Effects = drag == null ? DragDropEffects.None : drag.Source == FieldCatalog ? DragDropEffects.Copy : DragDropEffects.Move;
            e.Handled = true;
        }

        private void AxisDrop(object sender, DragEventArgs e)
        {
            var drag = GetDrag(e);
            if (drag == null) return;
            var axis = (ListBox)sender;
            var targetItem = FindParent<ListBoxItem>(e.OriginalSource as DependencyObject);
            int index = FieldsFor(axis).Count;
            if (targetItem?.DataContext is PivotField targetField)
            {
                index = FieldsFor(axis).IndexOf(targetField);
                if (e.GetPosition(targetItem).Y > targetItem.ActualHeight / 2) index++;
            }
            AssignField(axis, drag.Field, index);
            e.Handled = true;
        }

        private void ValuesDragOver(object sender, DragEventArgs e)
        {
            e.Effects = GetDrag(e) == null ? DragDropEffects.None : DragDropEffects.Copy;
            e.Handled = true;
        }

        private void ValuesDrop(object sender, DragEventArgs e)
        {
            var drag = GetDrag(e);
            if (drag != null) AddMeasure(drag.Field);
            e.Handled = true;
        }

        private void AddMeasureClicked(object sender, RoutedEventArgs e)
        {
            if (snapshot == null) return;
            AddMeasure((FieldCatalog.SelectedItem as CatalogField)?.Field ?? snapshot.Fields.FirstOrDefault(f => f.IsNumeric) ?? snapshot.Fields.FirstOrDefault());
        }

        private void AddCountClicked(object sender, RoutedEventArgs e)
        {
            if (measures.Any(m => m.Aggregation == PivotAggregation.CountRows))
            {
                BuilderMessage.Text = "Count rows is already included. Add another field to compare additional measures.";
                return;
            }
            measures.Add(new PivotMeasure { Value = -1, Aggregation = PivotAggregation.CountRows });
            RenderMeasures();
            Changed("Added Count rows.");
        }

        private void AddMeasure(PivotField field)
        {
            if (field == null) return;
            var choices = field.IsNumeric
                ? new[] { PivotAggregation.Sum, PivotAggregation.Average, PivotAggregation.Minimum, PivotAggregation.Maximum, PivotAggregation.CountValues, PivotAggregation.DistinctCount }
                : new[] { PivotAggregation.CountValues, PivotAggregation.DistinctCount };
            // The same field can have different aggregations, but identical measures add no information.
            var available = choices.Where(a => !measures.Any(m => m.Value == field.Index && m.Aggregation == a)).ToArray();
            if (available.Length == 0)
            {
                BuilderMessage.Text = field.Name + " is already in Values. Change an existing aggregation to compare another statistic.";
                return;
            }
            measures.Add(new PivotMeasure { Value = field.Index, Aggregation = available[0] });
            RenderMeasures();
            Changed("Added a measure for " + field.Name + ".");
        }

        private void RenderMeasures()
        {
            MeasureRows.Children.Clear();
            foreach (var measure in measures)
            {
                var row = new Grid { Margin = new Thickness(0, 0, 0, 5) };
                row.ColumnDefinitions.Add(new ColumnDefinition());
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                var inputs = new StackPanel();
                var aggregation = new ComboBox
                {
                    ItemsSource = Aggregations, DisplayMemberPath = "Label", Margin = new Thickness(0, 0, 3, 2),
                    SelectedItem = Aggregations.First(a => a.Value == measure.Aggregation),
                    ToolTip = "Aggregation for this measure"
                };
                AutomationProperties.SetName(aggregation, "Measure aggregation");
                var value = new ComboBox { DisplayMemberPath = "Label", Margin = new Thickness(0, 0, 3, 0), ToolTip = "Source field for this measure" };
                AutomationProperties.SetName(value, "Measure value field");
                bool refreshing = false;
                Action refresh = () =>
                {
                    refreshing = true;
                    try
                    {
                        var available = snapshot.Fields.Where(f => !PivotEngine.RequiresNumber(measure.Aggregation) || f.IsNumeric).ToArray();
                        value.ItemsSource = available;
                        value.IsEnabled = measure.Aggregation != PivotAggregation.CountRows;
                        value.Visibility = value.IsEnabled ? Visibility.Visible : Visibility.Collapsed;
                        value.SelectedItem = available.FirstOrDefault(f => f.Index == measure.Value) ?? available.FirstOrDefault();
                        measure.Value = measure.Aggregation == PivotAggregation.CountRows ? -1 : (value.SelectedItem as PivotField)?.Index ?? -1;
                    }
                    finally { refreshing = false; }
                };
                refresh();
                aggregation.SelectionChanged += (s, e) =>
                {
                    if (!(aggregation.SelectedItem is AggregationChoice choice)) return;
                    measure.Aggregation = choice.Value;
                    refresh();
                    Changed("Measure changed. Apply to update the result.");
                };
                value.SelectionChanged += (s, e) =>
                {
                    if (refreshing) return;
                    measure.Value = (value.SelectedItem as PivotField)?.Index ?? -1;
                    Changed("Measure field changed. Apply to update the result.");
                };
                inputs.Children.Add(aggregation);
                inputs.Children.Add(value);
                row.Children.Add(inputs);
                var remove = new Button { Content = "\u00D7", Padding = new Thickness(4, 0, 4, 0), MinWidth = 22, VerticalAlignment = VerticalAlignment.Top, ToolTip = "Remove this measure" };
                AutomationProperties.SetName(remove, "Remove measure");
                remove.Click += (s, e) =>
                {
                    measures.Remove(measure);
                    if (measures.Count == 0) measures.Add(new PivotMeasure { Aggregation = PivotAggregation.CountRows, Value = -1 });
                    RenderMeasures();
                    Changed("Measure removed. At least one measure is required.");
                };
                Grid.SetColumn(remove, 1);
                row.Children.Add(remove);
                MeasureRows.Children.Add(row);
            }
        }

        private void AddFilterClicked(object sender, RoutedEventArgs e) => EditFilter(null);

        private void EditFilter(PivotFilter current)
        {
            if (snapshot == null) return;
            using (var dialog = new PivotFilterWindow(snapshot, NullTextIsNull.IsChecked == true, current))
            {
                filterWindow = dialog;
                bool? accepted;
                try { accepted = dialog.ShowModal(); }
                finally { filterWindow = null; }
                if (disposed || accepted != true || dialog.Filter == null) return;
                int index = current == null ? -1 : filters.IndexOf(current);
                if (index < 0) filters.Add(dialog.Filter.Copy());
                else filters[index] = dialog.Filter.Copy();
                RenderFilters();
                Changed("Filters changed. Every filter must match (AND).");
            }
        }

        private void ClearFiltersClicked(object sender, RoutedEventArgs e)
        {
            if (filters.Count == 0) return;
            filters.Clear();
            RenderFilters();
            Changed("All filters cleared.");
        }

        private void RenderFilters()
        {
            FilterChips.Children.Clear();
            ClearFiltersButton.IsEnabled = filters.Count > 0;
            if (filters.Count == 0)
            {
                FilterChips.Children.Add(new TextBlock { Text = "All source rows", Opacity = 0.7, VerticalAlignment = VerticalAlignment.Center });
                return;
            }
            foreach (var filter in filters)
            {
                var chip = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 5, 3) };
                string label = FilterLabel(filter);
                var edit = new Button
                {
                    Content = new TextBlock { Text = label, TextTrimming = TextTrimming.CharacterEllipsis },
                    MaxWidth = 290, ToolTip = label + "\nClick to edit this filter", Padding = new Thickness(6, 3, 6, 3)
                };
                AutomationProperties.SetName(edit, "Edit filter: " + label);
                edit.Click += (s, e) => EditFilter(filter);
                chip.Children.Add(edit);
                var remove = new Button { Content = "\u00D7", ToolTip = "Remove filter: " + label, Padding = new Thickness(5, 3, 5, 3) };
                AutomationProperties.SetName(remove, "Remove filter: " + label);
                remove.Click += (s, e) => { filters.Remove(filter); RenderFilters(); Changed("Filter removed."); };
                chip.Children.Add(remove);
                FilterChips.Children.Add(chip);
            }
        }

        private string FilterLabel(PivotFilter filter)
        {
            var field = snapshot.Fields.FirstOrDefault(f => f.Index == filter.Field);
            string name = field == null ? "Field" : string.IsNullOrEmpty(field.Name) ||
                snapshot.Fields.Count(f => string.Equals(f.Name, field.Name, StringComparison.OrdinalIgnoreCase)) > 1
                ? field.Label : field.Name;
            switch (filter.Operator)
            {
                case PivotFilterOperator.In: return name + " is one of " + (filter.Values?.Length ?? 0).ToString(CultureInfo.InvariantCulture) + " values";
                case PivotFilterOperator.IsNull: return name + " is null";
                case PivotFilterOperator.IsNotNull: return name + " is not null";
                case PivotFilterOperator.Equals: return name + " = " + filter.Text;
                case PivotFilterOperator.NotEquals: return name + " != " + filter.Text;
                case PivotFilterOperator.GreaterThan: return name + " > " + filter.Text;
                case PivotFilterOperator.GreaterThanOrEqual: return name + " >= " + filter.Text;
                case PivotFilterOperator.LessThan: return name + " < " + filter.Text;
                case PivotFilterOperator.LessThanOrEqual: return name + " <= " + filter.Text;
                default: return name + " contains " + filter.Text;
            }
        }

        private void NullHandlingChanged(object sender, RoutedEventArgs e)
        {
            if (NullTextIsNull != null) Changed("NULL handling changed. Apply to update the result.");
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            updating = true;
            filterWindow?.Close();
            filterWindow = null;
            snapshot = null;
            rowFields.Clear();
            columnFields.Clear();
            measures.Clear();
            filters.Clear();
            MeasureRows.Children.Clear();
            FilterChips.Children.Clear();
            FieldCatalog.ItemsSource = null;
            RowAxis.ItemsSource = ColumnAxis.ItemsSource = null;
            dragSource = null;
            dragField = null;
            ConfigurationChanged = null;
            theme.Dispose();
        }
    }
}
