using Microsoft.VisualStudio.PlatformUI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace AxialSqlTools.PivotGrid
{
    public partial class PivotFilterWindow : DialogWindow, IDisposable
    {
        private readonly PivotSnapshot snapshot;
        private readonly bool nullTextIsNull;
        private readonly ToolWindowThemeController theme;
        private readonly DispatcherTimer searchTimer;
        private readonly HashSet<string> selectedValues = new HashSet<string>(StringComparer.Ordinal);
        private readonly PivotFilter existing;
        private CancellationTokenSource valueCancellation;
        private List<ValueChoice> values = new List<ValueChoice>();
        private ICollectionView valuesView;
        private string searchText = "";
        private int loadedField = -1;
        private bool updating;
        private bool initialized;
        private bool loading;
        private bool changingSelection;
        private bool disposed;
        private double valuesHeight = 570;

        internal PivotFilter Filter { get; private set; }

        internal PivotFilterWindow(PivotSnapshot snapshot, bool nullTextIsNull, PivotFilter existing = null)
        {
            this.snapshot = snapshot ?? throw new ArgumentNullException(nameof(snapshot));
            this.nullTextIsNull = nullTextIsNull;
            this.existing = existing;
            InitializeComponent();
            theme = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));
            searchTimer = new DispatcherTimer(DispatcherPriority.Background, Dispatcher)
            {
                Interval = TimeSpan.FromMilliseconds(180)
            };
            searchTimer.Tick += SearchTimerTick;
            FieldPicker.ItemsSource = snapshot.Fields;
            FieldPicker.SelectedItem = snapshot.Fields.FirstOrDefault(f => f.Index == existing?.Field)
                ?? snapshot.Fields.FirstOrDefault();
            TextValue.Text = existing?.Text ?? "";
            SetOperators(existing?.Operator ?? PivotFilterOperator.In);
            initialized = true;
            UpdateEditor();
        }

        private PivotField SelectedField => FieldPicker.SelectedItem as PivotField;
        private PivotFilterOperator SelectedOperator => OperatorPicker.SelectedValue is PivotFilterOperator value
            ? value : PivotFilterOperator.In;

        private void SetOperators(PivotFilterOperator preferred)
        {
            updating = true;
            var operators = new List<OperatorChoice>
            {
                new OperatorChoice(PivotFilterOperator.In, "Selected values"),
                new OperatorChoice(PivotFilterOperator.Contains, "Contains"),
                new OperatorChoice(PivotFilterOperator.Equals, "Equals"),
                new OperatorChoice(PivotFilterOperator.NotEquals, "Does not equal"),
                new OperatorChoice(PivotFilterOperator.IsNull, "Is NULL"),
                new OperatorChoice(PivotFilterOperator.IsNotNull, "Is not NULL")
            };
            if (SelectedField?.IsNumeric == true)
            {
                operators.Add(new OperatorChoice(PivotFilterOperator.GreaterThan, "Greater than"));
                operators.Add(new OperatorChoice(PivotFilterOperator.GreaterThanOrEqual, "Greater than or equal"));
                operators.Add(new OperatorChoice(PivotFilterOperator.LessThan, "Less than"));
                operators.Add(new OperatorChoice(PivotFilterOperator.LessThanOrEqual, "Less than or equal"));
            }
            OperatorPicker.ItemsSource = operators;
            OperatorPicker.SelectedValue = operators.Any(o => o.Value == preferred) ? preferred : PivotFilterOperator.In;
            updating = false;
        }

        private void FieldChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!initialized || updating || disposed) return;
            valueCancellation?.Cancel();
            loadedField = -1;
            loading = false;
            values = new List<ValueChoice>();
            valuesView = null;
            ValueList.ItemsSource = null;
            selectedValues.Clear();
            ValueSearch.Text = "";
            SetOperators(SelectedOperator);
            UpdateEditor();
        }

        private void OperatorChanged(object sender, SelectionChangedEventArgs e)
        {
            if (initialized && !updating && !disposed) UpdateEditor();
        }

        private void UpdateEditor()
        {
            ClearValidation();
            bool selectValues = SelectedOperator == PivotFilterOperator.In;
            bool isNull = SelectedOperator == PivotFilterOperator.IsNull || SelectedOperator == PivotFilterOperator.IsNotNull;
            bool wasValues = ValuesPanel.Visibility == Visibility.Visible;
            if (wasValues && !selectValues) valuesHeight = Height;
            ValuesPanel.Visibility = selectValues ? Visibility.Visible : Visibility.Collapsed;
            TextPanel.Visibility = !selectValues && !isNull ? Visibility.Visible : Visibility.Collapsed;
            NullHint.Visibility = isNull ? Visibility.Visible : Visibility.Collapsed;
            MinHeight = selectValues ? 440 : 280;
            if (wasValues != selectValues) Height = selectValues ? Math.Max(440, valuesHeight) : 300;
            TextHint.Text = RequiresNumber(SelectedField, SelectedOperator)
                ? "Enter a number using a period for decimals. Scientific notation is supported."
                : SelectedOperator == PivotFilterOperator.Contains
                    ? "Match text anywhere in the value. Matching is case-insensitive."
                    : "Compare the complete value, ignoring case. Use 'Selected values' to pick exact source values, including blanks.";
            NullHint.Text = SelectedOperator == PivotFilterOperator.IsNull
                ? "Include rows where this field is NULL."
                : "Include rows where this field has a non-NULL value.";
            ApplyButton.IsEnabled = SelectedField != null && (!selectValues || !loading);
            if (selectValues && SelectedField != null && loadedField != SelectedField.Index && !loading)
                LoadValuesAsync(SelectedField);
        }

        private async void LoadValuesAsync(PivotField field)
        {
            valueCancellation?.Cancel();
            var cancellation = new CancellationTokenSource();
            valueCancellation = cancellation;
            loading = true;
            ApplyButton.IsEnabled = false;
            SelectVisibleButton.IsEnabled = ClearVisibleButton.IsEnabled = false;
            SelectionInfo.Text = "Loading distinct values...";
            try
            {
                bool normalizeNull = nullTextIsNull || field.IsNumeric;
                var initialSelection = existing != null && existing.Field == field.Index && existing.Operator == PivotFilterOperator.In
                    ? new HashSet<string>((existing.Values ?? new string[0]).Select(v => Normalize(v, normalizeNull)), StringComparer.Ordinal)
                    : new HashSet<string>(StringComparer.Ordinal);
                var distinct = await Task.Run(() =>
                {
                    var unique = new HashSet<string>(StringComparer.Ordinal);
                    for (int row = 0; row < snapshot.Rows.Length; row++)
                    {
                        if ((row & 2047) == 0) cancellation.Token.ThrowIfCancellationRequested();
                        unique.Add(Normalize(snapshot.Rows[row][field.Index], normalizeNull));
                    }
                    cancellation.Token.ThrowIfCancellationRequested();
                    var choices = unique.OrderBy(v => v, StringComparer.CurrentCultureIgnoreCase)
                        .ThenBy(v => v, StringComparer.Ordinal)
                        .Select(v => new ValueChoice(v, initialSelection.Contains(v), SelectionChanged)).ToList();
                    cancellation.Token.ThrowIfCancellationRequested();
                    return choices;
                }, cancellation.Token);
                if (disposed || cancellation.IsCancellationRequested || !ReferenceEquals(valueCancellation, cancellation)) return;
                values = distinct;
                loadedField = field.Index;
                selectedValues.Clear();
                foreach (var choice in values)
                    if (choice.IsSelected) selectedValues.Add(choice.Value);
                valuesView = new ListCollectionView(values);
                valuesView.Filter = item => MatchesSearch((ValueChoice)item);
                ValueList.ItemsSource = valuesView;
                ApplySearch();
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                if (!disposed && ReferenceEquals(valueCancellation, cancellation))
                {
                    SelectionInfo.Text = "Distinct values could not be loaded.";
                    ShowValidation("Unable to load values: " + ex.Message);
                }
            }
            finally
            {
                if (!disposed && ReferenceEquals(valueCancellation, cancellation))
                {
                    valueCancellation = null;
                    loading = false;
                    bool loaded = loadedField == field.Index;
                    SelectVisibleButton.IsEnabled = ClearVisibleButton.IsEnabled = loaded;
                    ApplyButton.IsEnabled = SelectedField != null && (SelectedOperator != PivotFilterOperator.In || loaded);
                }
                cancellation.Dispose();
            }
        }

        private static string Normalize(string value, bool normalizeNull) => normalizeNull && value == "NULL" ? null : value;

        private static bool IsNumericComparison(PivotFilterOperator operation) =>
            operation == PivotFilterOperator.GreaterThan || operation == PivotFilterOperator.GreaterThanOrEqual ||
            operation == PivotFilterOperator.LessThan || operation == PivotFilterOperator.LessThanOrEqual;

        private static bool RequiresNumber(PivotField field, PivotFilterOperator operation) => field?.IsNumeric == true &&
            (IsNumericComparison(operation) || operation == PivotFilterOperator.Equals || operation == PivotFilterOperator.NotEquals);

        private void SearchChanged(object sender, TextChangedEventArgs e)
        {
            if (!initialized || disposed) return;
            searchTimer.Stop();
            searchTimer.Start();
        }

        private void SearchTimerTick(object sender, EventArgs e)
        {
            searchTimer.Stop();
            ApplySearch();
        }

        private void ApplySearch()
        {
            searchText = ValueSearch.Text ?? "";
            valuesView?.Refresh();
            UpdateSelectionInfo();
        }

        private bool MatchesSearch(ValueChoice choice) => string.IsNullOrEmpty(searchText) ||
            (choice.Value ?? "(NULL)").IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
            choice.Display.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0;

        private void SelectionChanged(ValueChoice choice)
        {
            if (choice.IsSelected) selectedValues.Add(choice.Value);
            else selectedValues.Remove(choice.Value);
            if (!changingSelection)
            {
                ClearValidation();
                UpdateSelectionInfo();
            }
        }

        private void UpdateSelectionInfo()
        {
            if (loading && loadedField < 0) return;
            int shown = valuesView is ListCollectionView view ? view.Count : 0;
            SelectionInfo.Text = string.Format(CultureInfo.CurrentCulture, "{0:N0} selected of {1:N0} values. {2:N0} shown. Search does not clear selections.",
                selectedValues.Count, values.Count, shown);
        }

        private void SelectVisibleClicked(object sender, RoutedEventArgs e) => SelectVisible(true);
        private void ClearVisibleClicked(object sender, RoutedEventArgs e) => SelectVisible(false);

        private void SelectVisible(bool select)
        {
            searchTimer.Stop();
            ApplySearch();
            changingSelection = true;
            try
            {
                foreach (var choice in values)
                    if (MatchesSearch(choice)) choice.IsSelected = select;
            }
            finally { changingSelection = false; }
            ClearValidation();
            UpdateSelectionInfo();
        }

        private void ValueListKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Space || Keyboard.Modifiers != ModifierKeys.None) return;
            // The checkbox's own Space handling remains available when it holds focus.
            if (e.OriginalSource is CheckBox) return;
            if (ValueList.SelectedItem is ValueChoice choice)
            {
                choice.IsSelected = !choice.IsSelected;
                e.Handled = true;
            }
        }

        private void TextValueChanged(object sender, TextChangedEventArgs e)
        {
            if (initialized) ClearValidation();
        }

        private void ApplyClicked(object sender, RoutedEventArgs e)
        {
            var field = SelectedField;
            if (field == null) { ShowValidation("Choose a field."); return; }
            var operation = SelectedOperator;
            if (operation == PivotFilterOperator.In && (loading || loadedField != field.Index)) return;
            if (operation == PivotFilterOperator.In && selectedValues.Count == 0)
            {
                ShowValidation("Select at least one value, or cancel to leave this filter unchanged.");
                ValueList.Focus();
                return;
            }
            string comparison = TextValue.Text ?? "";
            if (RequiresNumber(field, operation) && (comparison != "NULL" || IsNumericComparison(operation)))
            {
                decimal number;
                // Use the engine's precision check so high-precision SQL decimals cannot silently round here.
                comparison = comparison.Trim();
                if (!PivotEngine.TryParseDecimalExact(comparison, out number))
                {
                    ShowValidation("Enter an exactly representable decimal number, using a period for decimals. Round or cast higher-precision values in SQL first.");
                    TextValue.Focus();
                    TextValue.SelectAll();
                    return;
                }
            }
            Filter = new PivotFilter
            {
                Field = field.Index,
                Operator = operation,
                Text = comparison,
                Values = operation == PivotFilterOperator.In
                    ? values.Where(v => v.IsSelected).Select(v => v.Value).ToArray() : new string[0]
            };
            DialogResult = true;
        }

        private void ShowValidation(string message)
        {
            ValidationMessage.Text = message;
            ValidationMessage.Visibility = Visibility.Visible;
        }

        private void ClearValidation()
        {
            ValidationMessage.Text = "";
            ValidationMessage.Visibility = Visibility.Collapsed;
        }

        protected override void OnClosed(EventArgs e)
        {
            Dispose();
            base.OnClosed(e);
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            valueCancellation?.Cancel();
            searchTimer.Stop();
            searchTimer.Tick -= SearchTimerTick;
            ValueList.ItemsSource = null;
            valuesView = null;
            values.Clear();
            selectedValues.Clear();
            theme.Dispose();
        }

        private sealed class OperatorChoice
        {
            public PivotFilterOperator Value { get; }
            public string Label { get; }
            public OperatorChoice(PivotFilterOperator value, string label) { Value = value; Label = label; }
        }

        private sealed class ValueChoice : INotifyPropertyChanged
        {
            private readonly Action<ValueChoice> selectionChanged;
            private bool isSelected;
            public string Value { get; }
            public string Display { get; }
            public bool IsSelected
            {
                get => isSelected;
                set
                {
                    if (isSelected == value) return;
                    isSelected = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                    selectionChanged(this);
                }
            }
            public event PropertyChangedEventHandler PropertyChanged;
            public ValueChoice(string value, bool selected, Action<ValueChoice> selectionChanged)
            {
                Value = value;
                Display = value == null ? "(NULL)" : value.Length == 0 ? "(Empty string)" :
                    value == "(NULL)" || value == "(Empty string)" ? value + " (text)" :
                    value.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
                isSelected = selected;
                this.selectionChanged = selectionChanged;
            }
        }
    }
}
