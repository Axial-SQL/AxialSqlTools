using Microsoft.VisualStudio.Shell;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using NReco.PivotData;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace AxialSqlTools
{
    internal static class PivotTableTabManager
    {
        private static TabPage _pivotTab;
        private static PivotTableView _pivotView;

        public static void ShowPivotTab(DataTable dataTable)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var sqlResultsControl = GridAccess.GetSQLResultsControl();
            if (sqlResultsControl == null)
                return;

            var tabControl = GridAccess.GetTabControl(sqlResultsControl, "MarshallingControl") as TabControl;
            if (tabControl == null)
                return;

            if (_pivotTab == null || _pivotTab.IsDisposed)
            {
                _pivotTab = new TabPage("Pivot Table");
                _pivotView = new PivotTableView
                {
                    Dock = DockStyle.Fill
                };

                _pivotTab.Controls.Add(_pivotView);
                tabControl.TabPages.Add(_pivotTab);
            }

            _pivotView?.LoadData(dataTable);
            tabControl.SelectedTab = _pivotTab;
        }
    }

    internal sealed class PivotTableView : UserControl
    {
        private readonly SplitContainer _splitContainer;
        private readonly CheckedListBox _rowsList;
        private readonly CheckedListBox _columnsList;
        private readonly ComboBox _valueFieldCombo;
        private readonly ComboBox _aggregatorCombo;
        private readonly Button _applyButton;
        private readonly Label _statusLabel;
        private readonly WebView2 _webView;

        private DataTable _dataTable;
        private Task _webViewInitTask;

        private enum AggregationType
        {
            Sum,
            Count,
            Average,
            Min,
            Max
        }

        public PivotTableView()
        {
            _splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 280,
                FixedPanel = FixedPanel.Panel1
            };

            _rowsList = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true
            };

            _columnsList = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true
            };

            _valueFieldCombo = new ComboBox
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            _aggregatorCombo = new ComboBox
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            _applyButton = new Button
            {
                Text = "Build Pivot",
                Dock = DockStyle.Top,
                Height = 30
            };

            _applyButton.Click += async (s, e) => await RenderPivotAsync();

            _statusLabel = new Label
            {
                Dock = DockStyle.Top,
                Padding = new Padding(4),
                ForeColor = Color.DimGray,
                AutoSize = false,
                Height = 36
            };

            _webView = new WebView2
            {
                Dock = DockStyle.Fill
            };

            _splitContainer.Panel1.Controls.Add(CreateLayoutPanel());
            _splitContainer.Panel2.Controls.Add(_webView);

            Controls.Add(_splitContainer);

            PopulateAggregatorOptions();
        }

        public void LoadData(DataTable dataTable)
        {
            _dataTable = dataTable ?? throw new ArgumentNullException(nameof(dataTable));

            _rowsList.Items.Clear();
            _columnsList.Items.Clear();
            _valueFieldCombo.Items.Clear();

            foreach (DataColumn column in _dataTable.Columns)
            {
                _rowsList.Items.Add(column.ColumnName, false);
                _columnsList.Items.Add(column.ColumnName, false);
                _valueFieldCombo.Items.Add(column.ColumnName);
            }

            if (_dataTable.Columns.Count > 0)
            {
                _rowsList.SetItemChecked(0, true);
                if (_dataTable.Columns.Count > 1)
                    _columnsList.SetItemChecked(1, true);

                var defaultValueColumn = GetDefaultValueColumn();
                if (!string.IsNullOrEmpty(defaultValueColumn))
                    _valueFieldCombo.SelectedItem = defaultValueColumn;
                else
                    _valueFieldCombo.SelectedIndex = 0;
            }

            _aggregatorCombo.SelectedIndex = 0;
            _ = RenderPivotAsync();
        }

        private TableLayoutPanel CreateLayoutPanel()
        {
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 12,
                Padding = new Padding(6)
            };

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label { Text = "Rows", Dock = DockStyle.Top, Font = new Font(Font, FontStyle.Bold) });
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 35));
            layout.Controls.Add(_rowsList);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label { Text = "Columns", Dock = DockStyle.Top, Font = new Font(Font, FontStyle.Bold) });
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 35));
            layout.Controls.Add(_columnsList);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label { Text = "Values", Dock = DockStyle.Top, Font = new Font(Font, FontStyle.Bold) });
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(_valueFieldCombo);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label { Text = "Aggregate", Dock = DockStyle.Top, Font = new Font(Font, FontStyle.Bold) });
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(_aggregatorCombo);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(_applyButton);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(_statusLabel);

            return layout;
        }

        private void PopulateAggregatorOptions()
        {
            _aggregatorCombo.Items.Clear();
            foreach (var aggr in Enum.GetValues(typeof(AggregationType)))
                _aggregatorCombo.Items.Add(aggr.ToString());
        }

        private string GetDefaultValueColumn()
        {
            foreach (DataColumn column in _dataTable.Columns)
            {
                if (IsNumeric(column.DataType))
                    return column.ColumnName;
            }
            return _dataTable.Columns.Count > 0 ? _dataTable.Columns[0].ColumnName : string.Empty;
        }

        private async Task EnsureWebViewAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (_webViewInitTask == null)
            {
                var userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "AxialSqlTools",
                    "WebView2");

                Directory.CreateDirectory(userDataFolder);

                _webView.CreationProperties = new CoreWebView2CreationProperties
                {
                    UserDataFolder = userDataFolder
                };

                _webViewInitTask = _webView.EnsureCoreWebView2Async();
            }
        }

        private async Task RenderPivotAsync()
        {
            if (_dataTable == null || _dataTable.Columns.Count == 0)
                return;

            await EnsureWebViewAsync();

            if (_webView?.CoreWebView2 == null)
            {
                _statusLabel.Text = "Unable to initialize WebView2. Please reopen the Pivot Table tab.";
                return;
            }

            var rowDimensions = _rowsList.CheckedItems.Cast<string>().ToArray();
            var columnDimensions = _columnsList.CheckedItems.Cast<string>().ToArray();
            var valueField = _valueFieldCombo.SelectedItem as string ?? _dataTable.Columns[0].ColumnName;

            var aggregationType = AggregationType.Sum;
            if (_aggregatorCombo.SelectedItem is string selectedAggr &&
                Enum.TryParse(selectedAggr, out AggregationType parsedAggr))
            {
                aggregationType = parsedAggr;
            }

            var aggregatorFactory = CreateAggregatorFactory(aggregationType, valueField);
            var pivotData = new PivotData(rowDimensions.Concat(columnDimensions).ToArray(), aggregatorFactory, true);
            pivotData.ProcessData(new DataTableReader(_dataTable));

            var pivotTable = new PivotTable(rowDimensions, columnDimensions, pivotData);
            var html = BuildPivotHtml(pivotTable, valueField, aggregationType.ToString());

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            _webView.CoreWebView2.NavigateToString(html);
            _statusLabel.Text = $"Rows: {rowDimensions.Length}, Columns: {columnDimensions.Length}, Field: {valueField}, Aggr: {aggregationType}";
        }

        private static IAggregatorFactory CreateAggregatorFactory(AggregationType aggregationType, string valueField)
        {
            switch (aggregationType)
            {
                case AggregationType.Count:
                    return new CountAggregatorFactory();
                case AggregationType.Average:
                    return new AverageAggregatorFactory(valueField);
                case AggregationType.Min:
                    return new MinAggregatorFactory(valueField);
                case AggregationType.Max:
                    return new MaxAggregatorFactory(valueField);
                default:
                    return new SumAggregatorFactory(valueField);
            }
        }

        private static bool IsNumeric(Type type)
        {
            if (type == null)
                return false;

            var numericTypes = new HashSet<Type>
            {
                typeof(byte), typeof(short), typeof(int), typeof(long), typeof(float), typeof(double), typeof(decimal), typeof(sbyte), typeof(ushort), typeof(uint), typeof(ulong)
            };

            return numericTypes.Contains(Nullable.GetUnderlyingType(type) ?? type);
        }

        private static string BuildPivotHtml(PivotTable pivotTable, string valueField, string aggregationName)
        {
            var sb = new StringBuilder();
            sb.Append("<!DOCTYPE html><html><head><meta charset=\"UTF-8\">");
            sb.Append("<style>");
            sb.Append("body{font-family:'Segoe UI',Arial,sans-serif;margin:0;padding:12px;background:#f7f7f7;}");
            sb.Append("table{border-collapse:collapse;width:100%;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.1);}");
            sb.Append("th,td{border:1px solid #e0e0e0;padding:6px 10px;text-align:right;font-size:12px;}");
            sb.Append("th{background:#fafafa;font-weight:600;color:#555;}");
            sb.Append("td:first-child,th:first-child{text-align:left;white-space:nowrap;}");
            sb.Append(".total{background:#f0f4ff;font-weight:700;}");
            sb.Append(".meta{margin-bottom:8px;color:#666;font-size:12px;}");
            sb.Append("</style></head><body>");

            sb.Append($"<div class='meta'>Aggregation: {aggregationName} on {valueField}</div>");
            sb.Append("<table>");

            sb.Append("<thead><tr><th>");
            sb.Append(string.Join(" / ", pivotTable.Rows));
            sb.Append("</th>");

            foreach (var colKey in pivotTable.ColumnKeys)
            {
                sb.Append("<th>");
                sb.Append(FormatKey(colKey));
                sb.Append("</th>");
            }

            sb.Append("<th class='total'>Total</th></tr></thead><tbody>");

            for (int r = 0; r < pivotTable.RowKeys.Length; r++)
            {
                sb.Append("<tr><th>");
                sb.Append(FormatKey(pivotTable.RowKeys[r]));
                sb.Append("</th>");

                for (int c = 0; c < pivotTable.ColumnKeys.Length; c++)
                {
                    sb.Append("<td>");
                    sb.Append(FormatAggregatorValue(pivotTable[r, c]));
                    sb.Append("</td>");
                }

                sb.Append("<td class='total'>");
                sb.Append(FormatAggregatorValue(pivotTable[r, null]));
                sb.Append("</td></tr>");
            }

            sb.Append("<tr class='total'><th>Total</th>");
            for (int c = 0; c < pivotTable.ColumnKeys.Length; c++)
            {
                sb.Append("<td class='total'>");
                sb.Append(FormatAggregatorValue(pivotTable[null, c]));
                sb.Append("</td>");
            }

            sb.Append("<td class='total'>");
            sb.Append(FormatAggregatorValue(pivotTable[null, null]));
            sb.Append("</td></tr>");

            sb.Append("</tbody></table></body></html>");
            return sb.ToString();
        }

        private static string FormatKey(ValueKey key)
        {
            if (key == null || key.DimKeys == null || key.DimKeys.Length == 0)
                return "(All)";

            var parts = key.DimKeys.Select(v =>
            {
                if (v == null)
                    return "(null)";
                if (v.Equals(Key.Empty))
                    return "(All)";
                return Convert.ToString(v, CultureInfo.InvariantCulture);
            });

            return string.Join(" / ", parts);
        }

        private static string FormatAggregatorValue(IAggregator aggregator)
        {
            if (aggregator == null || aggregator.Count == 0)
                return string.Empty;

            var value = aggregator.Value;
            if (value is object[] values)
            {
                var formatted = values.Select(v => FormatValue(v));
                return string.Join(" | ", formatted);
            }

            return FormatValue(value);
        }

        private static string FormatValue(object value)
        {
            if (value == null)
                return string.Empty;

            if (value is IFormattable formattable)
                return formattable.ToString("G", CultureInfo.InvariantCulture);

            return value.ToString();
        }
    }
}
