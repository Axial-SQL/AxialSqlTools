using Microsoft.VisualStudio.PlatformUI;
using System;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace AxialSqlTools.PivotGrid
{
    public partial class PivotDetailsWindow : DialogWindow, IDisposable
    {
        private readonly ToolWindowThemeController theme;
        private readonly PivotDetails details;
        private readonly PivotField[] fields;
        private int pageIndex;
        private bool disposed;

        internal PivotDetailsWindow(PivotDetails details, PivotField[] fields)
        {
            InitializeComponent();
            this.details = details;
            this.fields = fields;
            theme = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));
            DetailsTitle.Text = details.Description;
            ShowPage();
        }

        private void ShowPage()
        {
            var page = details.GetPage(pageIndex);
            if (DetailsGrid.Columns.Count == 0)
            {
                var textStyle = PivotGridWindowControl.CreateTextStyle(false);
                var numericStyle = PivotGridWindowControl.CreateTextStyle(true);
                var cellStyle = PivotGridWindowControl.CreateCellStyle(this, false, null);
                foreach (DataColumn column in page.Columns)
                {
                    // Detail columns contain display strings, so use the original SQL type metadata.
                    bool numeric = column.Ordinal == 0 || fields[column.Ordinal - 1].IsNumeric;
                    DetailsGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = column.Caption,
                        Binding = new Binding("[" + column.ColumnName + "]")
                        {
                            Mode = BindingMode.OneWay, Converter = PivotGridWindowControl.CellDisplay,
                            ConverterParameter = numeric, TargetNullValue = "(NULL)"
                        },
                        ElementStyle = numeric ? numericStyle : textStyle,
                        CellStyle = cellStyle,
                        Width = new DataGridLength(column.Ordinal == 0 ? 100 : 180)
                    });
                }
                DetailsGrid.FrozenColumnCount = 1;
            }
            DetailsGrid.ItemsSource = page.DefaultView;
            int first = details.Count == 0 ? 0 : pageIndex * details.PageSize + 1;
            int last = Math.Min(details.Count, (pageIndex + 1) * details.PageSize);
            DetailsPageInfo.Text = string.Format("Rows {0:N0}-{1:N0} of {2:N0}. Page {3:N0}/{4:N0}. Ctrl+C copies selected cells on this page.",
                first, last, details.Count, pageIndex + 1, details.PageCount);
            PreviousPageButton.IsEnabled = pageIndex > 0;
            NextPageButton.IsEnabled = pageIndex + 1 < details.PageCount;
        }

        private void PreviousPageClicked(object sender, RoutedEventArgs e)
        {
            if (!disposed && pageIndex > 0) { pageIndex--; ShowPage(); }
        }

        private void NextPageClicked(object sender, RoutedEventArgs e)
        {
            if (!disposed && pageIndex + 1 < details.PageCount) { pageIndex++; ShowPage(); }
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
            DetailsGrid.ItemsSource = null;
            DetailsGrid.Columns.Clear();
            theme.Dispose();
        }
    }
}
