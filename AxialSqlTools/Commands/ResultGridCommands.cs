using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Spreadsheet;
using EnvDTE;
using EnvDTE80;
using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{

    internal abstract class ResultGridCommandBase
    {
        private const string SQL_RESULT_GRID_CONTEXT_NAME = "SQL Results Grid Tab Context";
        protected static CommandBar GridCommandBar { get; }

        static ResultGridCommandBase()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
            GridCommandBar = ((CommandBars)dte.CommandBars)[SQL_RESULT_GRID_CONTEXT_NAME];
        }
    }

    internal sealed class ResultGridCopyAsInsertCommand : ResultGridCommandBase
    {

        static ResultGridCopyAsInsertCommand()
        {
            
        }

        public static async Task InitializeAsync(AxialSqlToolsPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var copyAllPopup = (CommandBarPopup)GridCommandBar.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
            copyAllPopup.Visible = true;
            copyAllPopup.Caption = "Copy All As ...";

            var btnControlAllInsert = (CommandBarButton)copyAllPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlAllInsert.Visible = true;
            btnControlAllInsert.Caption = "INSERT";
            btnControlAllInsert.Click += OnClick_CopyAllAsInsert;

            var btnControlAllCsv = (CommandBarButton)copyAllPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlAllCsv.Visible = true;
            btnControlAllCsv.Caption = "CSV";
            btnControlAllCsv.Click += OnClick_CopyAllAsCsv;

            var btnControlAllJson = (CommandBarButton)copyAllPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlAllJson.Visible = true;
            btnControlAllJson.Caption = "JSON";
            btnControlAllJson.Click += OnClick_CopyAllAsJson;

            var btnControlAllXml = (CommandBarButton)copyAllPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlAllXml.Visible = true;
            btnControlAllXml.Caption = "XML";
            btnControlAllXml.Click += OnClick_CopyAllAsXml;

            var btnControlAllHtml = (CommandBarButton)copyAllPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlAllHtml.Visible = true;
            btnControlAllHtml.Caption = "HTML";
            btnControlAllHtml.Click += OnClick_CopyAllAsHtml;

            var btnControlAllInClauseList = (CommandBarButton)copyAllPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlAllInClauseList.Visible = true;
            btnControlAllInClauseList.Caption = "IN Clause List";
            btnControlAllInClauseList.Click += OnClick_CopyAllAsInClauseList;

            var copySelectedPopup = (CommandBarPopup)GridCommandBar.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
            copySelectedPopup.Visible = true;
            copySelectedPopup.Caption = "Copy Selected As ...";

            var btnControlSelectedInsert = (CommandBarButton)copySelectedPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlSelectedInsert.Visible = true;
            btnControlSelectedInsert.Caption = "INSERT";
            btnControlSelectedInsert.Click += OnClick_CopySelectedAsInsert;

            var btnControlSelectedCsv = (CommandBarButton)copySelectedPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlSelectedCsv.Visible = true;
            btnControlSelectedCsv.Caption = "CSV";
            btnControlSelectedCsv.Click += OnClick_CopySelectedAsCsv;

            var btnControlSelectedJson = (CommandBarButton)copySelectedPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlSelectedJson.Visible = true;
            btnControlSelectedJson.Caption = "JSON";
            btnControlSelectedJson.Click += OnClick_CopySelectedAsJson;

            var btnControlSelectedXml = (CommandBarButton)copySelectedPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlSelectedXml.Visible = true;
            btnControlSelectedXml.Caption = "XML";
            btnControlSelectedXml.Click += OnClick_CopySelectedAsXml;

            var btnControlSelectedHtml = (CommandBarButton)copySelectedPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlSelectedHtml.Visible = true;
            btnControlSelectedHtml.Caption = "HTML";
            btnControlSelectedHtml.Click += OnClick_CopySelectedAsHtml;

            var btnControlSelectedInClauseList = (CommandBarButton)copySelectedPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlSelectedInClauseList.Visible = true;
            btnControlSelectedInClauseList.Caption = "IN Clause List";
            btnControlSelectedInClauseList.Click += OnClick_CopySelectedAsInClauseList;

            var btnControlCCN = (CommandBarButton)GridCommandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlCCN.Visible = true;
            btnControlCCN.Caption = "Copy Selected Column Names";
            btnControlCCN.Click += OnClick_CopySelectedColumnNames;

            var btnControlCCNA = (CommandBarButton)GridCommandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlCCNA.Visible = true;
            btnControlCCNA.Caption = "Copy All Column Names";
            btnControlCCNA.Click += OnClick_CopyAllColumnNames ;

        }

        private static void OnClick_CopyAllAsInsert(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.All, CopyFormat.Insert);
        private static void OnClick_CopyAllAsCsv(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.All, CopyFormat.Csv);
        private static void OnClick_CopyAllAsJson(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.All, CopyFormat.Json);
        private static void OnClick_CopyAllAsXml(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.All, CopyFormat.Xml);
        private static void OnClick_CopyAllAsHtml(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.All, CopyFormat.Html);
        private static void OnClick_CopyAllAsInClauseList(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.All, CopyFormat.InClauseList);
        private static void OnClick_CopySelectedAsInsert(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.Selected, CopyFormat.Insert);
        private static void OnClick_CopySelectedAsCsv(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.Selected, CopyFormat.Csv);
        private static void OnClick_CopySelectedAsJson(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.Selected, CopyFormat.Json);
        private static void OnClick_CopySelectedAsXml(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.Selected, CopyFormat.Xml);
        private static void OnClick_CopySelectedAsHtml(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.Selected, CopyFormat.Html);
        private static void OnClick_CopySelectedAsInClauseList(CommandBarButton Ctrl, ref bool CancelDefault) => CopyValues(CopyScope.Selected, CopyFormat.InClauseList);

        private static void OnClick_CopySelectedColumnNames(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            CopyColumnNames(all: false);
        }

        private static void OnClick_CopyAllColumnNames(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            CopyColumnNames(all: true);
        }

        private static void CopyColumnNames(bool all)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var focusGridControl = GridAccess.GetFocusGridControl();
            using (var gridResultControl = new ResultGridControlAdaptor(focusGridControl))
            {

                var columnNames = new List<string>();

                if (all)                
                {
                    for (int i = 1; i < gridResultControl.ColumnCount; i++)
                    {
                        columnNames.Add("[" + gridResultControl.GetColumnName(i) + "]");
                    }
                } else
                {
                    var selectedCells = focusGridControl.SelectedCells;

                    var addedColIndexes = new HashSet<int>();

                    foreach (BlockOfCells cell in selectedCells)
                    {
                        int leftColNumber = cell.X;
                        int rightColNumber = cell.Right;

                        for (int i = leftColNumber; i <= rightColNumber; i++)
                        {
                            if (addedColIndexes.Add(i)) // ignore dups
                                columnNames.Add("[" + gridResultControl.GetColumnName(i) + "]");
                        }
                    }

                }

                if (columnNames.Any())
                {
                    string columnNamesStr = string.Join(",", columnNames);

                    SetClipboardText(columnNamesStr);

                    ServiceCache.ExtensibilityModel.StatusBar.Text = "Copied Column Names";
                }
                else
                {
                    ServiceCache.ExtensibilityModel.StatusBar.Text = "No Column Names to Copy";
                }

            }
            
        }



        private enum CopyScope
        {
            All,
            Selected
        }

        private enum CopyFormat
        {
            Insert,
            Csv,
            Json,
            Xml,
            Html,
            InClauseList
        }

        private static void CopyValues(CopyScope scope, CopyFormat format)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var focusGridControl = GridAccess.GetFocusGridControl();
            using (var gridResultControl = new ResultGridControlAdaptor(focusGridControl))
            {
                if (format == CopyFormat.Insert)
                {
                    CopyInsert(scope, gridResultControl);
                    return;
                }

                if (format == CopyFormat.InClauseList)
                {
                    CopyInClauseList(scope, gridResultControl);
                    return;
                }

                var datatable = scope == CopyScope.All
                    ? gridResultControl.GridFocusAsDatatable()
                    : gridResultControl.GridSelectedAsDataTable();

                if (datatable.Columns.Count == 0 || datatable.Rows.Count == 0)
                {
                    ServiceCache.ExtensibilityModel.StatusBar.Text = scope == CopyScope.All
                        ? "No data to copy"
                        : "No cells selected to copy";
                    return;
                }

                string resultText;

                switch (format)
                {
                    case CopyFormat.Csv:
                        resultText = datatable.ToCsv();
                        break;
                    case CopyFormat.Json:
                        resultText = datatable.ToJson();
                        break;
                    case CopyFormat.Xml:
                        resultText = datatable.ToXml();
                        break;
                    case CopyFormat.Html:
                        resultText = datatable.ToHtml();
                        break;
                    default:
                        resultText = string.Empty;
                        break;
                }

                if (!string.IsNullOrWhiteSpace(resultText))
                {
                    SetClipboardText(resultText);
                    ServiceCache.ExtensibilityModel.StatusBar.Text = "Copied";
                }
            }
        }

        private static void CopyInsert(CopyScope scope, ResultGridControlAdaptor gridResultControl)
        {
            var columnHeaders = string.Empty;
            IEnumerable<string> contentRows = Enumerable.Empty<string>();

            if (scope == CopyScope.Selected)
            {
                var gridResultSelected = gridResultControl.GridSelectedAsQuerySql().ToList();
                if (gridResultSelected.Count <= 1)
                {
                    ServiceCache.ExtensibilityModel.StatusBar.Text = "No cells selected to copy";
                    return;
                }

                columnHeaders = gridResultSelected.First();
                contentRows = gridResultSelected.Skip(1).Select(r => $"({r})");
            }
            else
            {
                columnHeaders = string.Join(", ", gridResultControl.GetBracketColumns());
                contentRows = gridResultControl.GridAsQuerySql().Select(row => $"({string.Join(", ", row)})");

                if (!contentRows.Any())
                {
                    ServiceCache.ExtensibilityModel.StatusBar.Text = "No data to copy";
                    return;
                }
            }

            var valuesText = string.Join(",\r\n", contentRows);
            var resultText = $"INSERT INTO [table] ({columnHeaders}) VALUES\r\n" + valuesText;

            SetClipboardText(resultText);
            ServiceCache.ExtensibilityModel.StatusBar.Text = "Copied";
        }

        private static void CopyInClauseList(CopyScope scope, ResultGridControlAdaptor gridResultControl)
        {
            var columnIndex = gridResultControl.GetCurrentColumnIndex();
            if (!columnIndex.HasValue)
            {
                ServiceCache.ExtensibilityModel.StatusBar.Text = "No column selected to copy";
                return;
            }

            var column = columnIndex.Value;
            var values = new List<object>();

            if (scope == CopyScope.All)
            {
                if (gridResultControl.RowCount == 0)
                {
                    ServiceCache.ExtensibilityModel.StatusBar.Text = "No data to copy";
                    return;
                }

                for (var rowIndex = 0L; rowIndex < gridResultControl.RowCount; rowIndex++)
                {
                    values.Add(gridResultControl.GetCellValueAsString(rowIndex, column));
                }
            }
            else
            {
                var selectedRows = gridResultControl.GetSelectedRowIndexesForColumn(column).ToList();
                if (selectedRows.Count == 0)
                {
                    ServiceCache.ExtensibilityModel.StatusBar.Text = "No cells selected to copy";
                    return;
                }

                foreach (var rowIndex in selectedRows)
                {
                    values.Add(gridResultControl.GetCellValueAsString(rowIndex, column));
                }
            }

            var resultText = $"({string.Join(", ", values)})";
            SetClipboardText(resultText);
            ServiceCache.ExtensibilityModel.StatusBar.Text = "Copied";
        }

        [STAThread]
        public static void SetClipboardText(string value)
        {
            try
            {
                System.Windows.Forms.Clipboard.SetText(value);
            }
            catch // (Exception ex) //when (ex is TException)
            {

            }
        }


    }
}
