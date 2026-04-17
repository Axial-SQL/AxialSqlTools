using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;
using static AxialSqlTools.AxialSqlToolsPackage;

namespace AxialSqlTools
{
    public static class GridAccess
    {
        public static void SetPropertyValue(object targetObj, string fieldName, object fieldValue)
        {
            targetObj.GetType().GetProperty(fieldName).SetValue(targetObj, fieldValue);
        }

        public static object GetProperty(object obj, string field)
        {
            if (obj is null)
            {
                return null;
            }

            return obj.GetType().GetProperty(field, BindingFlags.Public | BindingFlags.Instance)?.GetValue(obj);
        }

        public static object GetNonPublicField(object obj, string field)
        {
            if (obj is null)
            {
                return null;
            }

            return obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(obj);
        }

        public static FieldInfo GetNonPublicFieldInfo(object obj, string field)
            => obj?.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance);

        public static object GetSQLResultsControl()
        {
            try
            {
                var factoryType = ServiceCache.ScriptFactory.GetType();

                var method = factoryType.GetMethod(
                    "GetCurrentlyActiveFrameDocView",
                    BindingFlags.NonPublic | BindingFlags.Instance);

                if (method == null)
                {
                    _logger?.Error("GetCurrentlyActiveFrameDocView method not found");
                    return null;
                }

                var docView = method.Invoke(
                    ServiceCache.ScriptFactory,
                    new object[] { ServiceCache.VSMonitorSelection, false, null });

                // It might be ObjectExplorerTool, a designer, etc.
                if (!(docView is SqlScriptEditorControl scriptEditor))
                {
                    // No active query window - handle as you like
                    // e.g. return null, or throw a more descriptive exception
                    return null;
                }

                return GetNonPublicField(scriptEditor, "m_sqlResultsControl");
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in GetSQLResultsControl");
                return null;
            }
        }

        public static QEStatusBarManager GetStatusBarManager()
        {
            var objType = ServiceCache.ScriptFactory.GetType();
            var method1 = objType.GetMethod("GetCurrentlyActiveFrameDocView", BindingFlags.NonPublic | BindingFlags.Instance);
            var Result = (SqlScriptEditorControl)method1.Invoke(ServiceCache.ScriptFactory, new object[] { ServiceCache.VSMonitorSelection, false, null });

            return Result.StatusBarManager;
        }

        public static CollectionBase GetGridContainers()
        {
            try
            {
                var SQLResultsControl = GetSQLResultsControl();
                if (SQLResultsControl is null)
                {
                    return null;
                }

                var m_gridResultsPage = GetNonPublicField(SQLResultsControl, "m_gridResultsPage");
                if (m_gridResultsPage is null)
                {
                    return null;
                }

                var gridContainersObj = GetNonPublicField(m_gridResultsPage, "m_gridContainers");
                if (gridContainersObj is null)
                {
                    return null;
                }

                // SSMS 21: Return CollectionBase directly
                if (gridContainersObj is CollectionBase gridContainers)
                {
                    return gridContainers;
                }

                // SSMS 22: Wrap List<ResultSetAndGridContainer> as CollectionBase
                if (gridContainersObj is ICollection collection)
                {
                    var wrapper = new GridContainerCollection();
                    foreach (var item in collection)
                    {
                        wrapper.AddItem(item);
                    }
                    return wrapper;
                }

                _logger?.Error($"Unexpected grid containers type: {gridContainersObj.GetType().FullName}");
                return null;
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in GetGridContainers");
                return null;
            }
        }

        // Wrapper for SSMS 22's List<T> to match SSMS 21's CollectionBase
        private class GridContainerCollection : CollectionBase
        {
            public void AddItem(object item)
            {
                List.Add(item);
            }
        }

        public static void ChangeStatusBarContent(int OpenTranCount, bool isColumnEncryptionSettingOn, string ActualElapsedTime)
        {
            QEStatusBarManager statusBarManager = GetStatusBarManager();

            if (OpenTranCount > 0)
            {
                var msg = "One transaction is still open!";
                if (OpenTranCount > 1)
                {
                    msg = $"{OpenTranCount} transactions are still open!";
                }

                var currentMsg = statusBarManager.StatusText;
                statusBarManager.SetKnownState(QEStatusBarKnownStates.Executing);
                statusBarManager.StatusText = currentMsg + " | " + msg;
            }

            if (isColumnEncryptionSettingOn)
            {
                var oeMsg = "Column Encryption Setting is ON";
                statusBarManager.StatusText = statusBarManager.StatusText + " | " + oeMsg;
            }

            var statusBarManager_executionTimePanel = GetNonPublicField(statusBarManager, "executionTimePanel");
            SetPropertyValue(statusBarManager_executionTimePanel, "Text", ActualElapsedTime);

            var statusBarManager_completedTimePanel = GetNonPublicField(statusBarManager, "completedTimePanel");
            //statusBarManager_completedTimePanel.Visible = true;

            var generalPanel = GetNonPublicField(statusBarManager, "generalPanel");
            if (OpenTranCount > 0)
            {
                SetPropertyValue(generalPanel, "ForeColor", Color.Red);
            }
            else
            {
                SetPropertyValue(generalPanel, "ForeColor", Color.Black);
            }

            // TODO - need to contract font from existing property..
            Font defaultFont = new Font("Segoe UI", 9);
            Font boldFont = new Font("Segoe UI", 10, FontStyle.Bold);

            var statusStrip = GetNonPublicField(statusBarManager, "statusStrip");
            if (OpenTranCount > 0 || isColumnEncryptionSettingOn)
            {
                SetPropertyValue(statusStrip, "Font", boldFont);
            }
            else
            {
                SetPropertyValue(statusStrip, "Font", defaultFont);
            }
        }

        public static Color? FindMatchingConnectionColor(string serverName, string databaseName)
        {
            var rules = SettingsManager.GetConnectionColorRules();

            foreach (var rule in rules)
            {
                if (!rule.IsEnabled)
                    continue;

                bool hasServerPattern = !string.IsNullOrEmpty(rule.ServerNamePattern);
                bool hasDatabasePattern = !string.IsNullOrEmpty(rule.DatabaseNamePattern);

                if (!hasServerPattern && !hasDatabasePattern)
                    continue;

                bool serverMatch = !hasServerPattern ||
                    (!string.IsNullOrEmpty(serverName) && serverName.IndexOf(rule.ServerNamePattern, StringComparison.OrdinalIgnoreCase) >= 0);

                bool databaseMatch = !hasDatabasePattern ||
                    (!string.IsNullOrEmpty(databaseName) && databaseName.IndexOf(rule.DatabaseNamePattern, StringComparison.OrdinalIgnoreCase) >= 0);

                if (serverMatch && databaseMatch)
                {
                    try
                    {
                        return ColorTranslator.FromHtml(rule.StatusBarColor);
                    }
                    catch { }
                }
            }

            return null;
        }

        public static void ApplyConnectionColor(string serverName, string databaseName)
        {
            // Status bar coloring - only for query windows
            try
            {
                var SQLResultsControl = GetSQLResultsControl();
                if (SQLResultsControl == null)
                    return;

                Color? matchedColor = FindMatchingConnectionColor(serverName, databaseName);

                var statusBarManager = GetStatusBarManager();
                if (statusBarManager == null)
                    return;

                var statusStripObj = GetNonPublicField(statusBarManager, "statusStrip");
                if (statusStripObj is StatusStrip statusStrip)
                {
                    if (matchedColor != null)
                    {
                        statusStrip.BackColor = matchedColor.Value;
                    }
                    else
                    {
                        statusStrip.BackColor = SystemColors.Control;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in ApplyConnectionColor");
            }
        }

        private static System.Windows.Threading.DispatcherTimer _reapplyTimer;
        private static int _reapplyRetryCount;

        public static void ColorAllDocumentTabs()
        {
            try
            {
                var wpfApp = System.Windows.Application.Current;
                if (wpfApp == null) return;
                var mainWindow = wpfApp.MainWindow;
                if (mainWindow == null) return;

                ColorAllTabs(mainWindow);
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in ColorAllDocumentTabs");
            }
        }

        private static void ColorAllTabs(System.Windows.DependencyObject mainWindow)
        {
            var tabItems = new List<System.Windows.DependencyObject>();
            FindElementsByTypeName(mainWindow, "DocumentTabItem", tabItems);

            if (tabItems.Count > 0)
            {
                foreach (var tabItem in tabItems)
                {
                    var headerCtrl = tabItem as System.Windows.Controls.HeaderedContentControl;
                    string header = headerCtrl?.Header?.ToString();
                    if (string.IsNullOrEmpty(header)) continue;

                    var parts = header.Split(new[] { " - " }, 2, StringSplitOptions.None);
                    if (parts.Length < 2)
                    {
                        var border0 = FindChildByTypeName(tabItem, "SimpleCurvedBorder")
                                   ?? FindChildByTypeName(tabItem, "TopCurvedBorder");
                        ApplyColorToElement(border0 ?? tabItem, null);
                        continue;
                    }

                    string connectionPart = parts[1].Trim();
                    int parenIdx = connectionPart.IndexOf(" (");
                    if (parenIdx > 0)
                        connectionPart = connectionPart.Substring(0, parenIdx);

                    Color? color = FindMatchingConnectionColor(connectionPart, connectionPart);

                    var curvedBorder = FindChildByTypeName(tabItem, "SimpleCurvedBorder")
                                   ?? FindChildByTypeName(tabItem, "TopCurvedBorder");
                    ApplyColorToElement(curvedBorder ?? tabItem, color);
                }
            }
            else
            {
                ColorSingleTabFallback(mainWindow);
            }
        }

        private static void ColorSingleTabFallback(System.Windows.DependencyObject root)
        {
            try
            {
                var docGroups = new List<System.Windows.DependencyObject>();
                FindElementsByTypeName(root, "DocumentGroupControl", docGroups);
                if (docGroups.Count == 0) return;

                foreach (var docGroup in docGroups)
                {
                    var header = FindChildByTypeName(docGroup, "DragUndockHeader");
                    if (header == null) continue;

                    var fe = header as System.Windows.FrameworkElement;
                    if (fe == null || fe.ActualHeight <= 0 || fe.Visibility != System.Windows.Visibility.Visible)
                        continue;

                    string headerTitle = fe.DataContext?.ToString();
                    Color? color = null;

                    if (!string.IsNullOrEmpty(headerTitle))
                    {
                        var parts = headerTitle.Split(new[] { " - " }, 2, StringSplitOptions.None);
                        if (parts.Length >= 2)
                        {
                            string connectionPart = parts[1].Trim();
                            int parenIdx = connectionPart.IndexOf(" (");
                            if (parenIdx > 0)
                                connectionPart = connectionPart.Substring(0, parenIdx);
                            color = FindMatchingConnectionColor(connectionPart, connectionPart);
                        }
                    }

                    ApplyColorToElement(header, color);
                    return;
                }
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in ColorSingleTabFallback");
            }
        }

        private static void ApplyColorToElement(System.Windows.DependencyObject element, Color? matchedColor)
        {
            if (element is System.Windows.Controls.Control ctrl)
            {
                if (matchedColor.HasValue)
                {
                    var wpfColor = System.Windows.Media.Color.FromRgb(
                        matchedColor.Value.R, matchedColor.Value.G, matchedColor.Value.B);
                    var brush = new System.Windows.Media.SolidColorBrush(wpfColor);
                    brush.Freeze();
                    ctrl.Background = brush;
                }
                else
                {
                    ctrl.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
                }
            }
            else if (element is System.Windows.Controls.Border border)
            {
                if (matchedColor.HasValue)
                {
                    var wpfColor = System.Windows.Media.Color.FromRgb(
                        matchedColor.Value.R, matchedColor.Value.G, matchedColor.Value.B);
                    var brush = new System.Windows.Media.SolidColorBrush(wpfColor);
                    brush.Freeze();
                    border.Background = brush;
                }
                else
                {
                    border.ClearValue(System.Windows.Controls.Border.BackgroundProperty);
                }
            }
        }

        private static System.Windows.DependencyObject FindChildByTypeName(System.Windows.DependencyObject parent, string typeName)
        {
            if (parent == null) return null;

            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child.GetType().Name == typeName)
                    return child;
            }
            for (int i = 0; i < count; i++)
            {
                var result = FindChildByTypeName(System.Windows.Media.VisualTreeHelper.GetChild(parent, i), typeName);
                if (result != null) return result;
            }
            return null;
        }

        private static void FindElementsByTypeName(System.Windows.DependencyObject obj, string typeName, List<System.Windows.DependencyObject> results)
        {
            if (obj == null) return;

            if (obj.GetType().Name == typeName)
            {
                results.Add(obj);
            }

            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(obj);
            for (int i = 0; i < count; i++)
            {
                FindElementsByTypeName(System.Windows.Media.VisualTreeHelper.GetChild(obj, i), typeName, results);
            }
        }

        public static void ScheduleReapplyAllTabColors()
        {
            if (_reapplyTimer == null)
            {
                _reapplyTimer = new System.Windows.Threading.DispatcherTimer();
                _reapplyTimer.Interval = TimeSpan.FromMilliseconds(150);
                _reapplyTimer.Tick += (s, e) =>
                {
                    _reapplyRetryCount++;
                    try { ColorAllDocumentTabs(); } catch { }

                    if (_reapplyRetryCount >= 4)
                    {
                        _reapplyTimer.Stop();
                    }
                };
            }

            _reapplyRetryCount = 0;
            _reapplyTimer.Stop();
            _reapplyTimer.Start();
        }

        public static string GetColumnSqlType(DataRow schemaRow)
        {
            int sqlDataColumnSize = (int)schemaRow[2];
            int? NumericPrecision = schemaRow[3] != DBNull.Value ? Convert.ToInt32(schemaRow[3]) : (int?)null;
            int? NumericScale = schemaRow[4] != DBNull.Value ? Convert.ToInt32(schemaRow[4]) : (int?)null;

            string sqlDataTypeName = ((string)schemaRow[24]).ToUpper();

            if (sqlDataTypeName == "NVARCHAR" || sqlDataTypeName == "NCHAR"
                || sqlDataTypeName == "VARCHAR" || sqlDataTypeName == "CHAR"
                || sqlDataTypeName == "VARBINARY" || sqlDataTypeName == "BINARY")
            {
                return sqlDataTypeName + "(" + (sqlDataColumnSize == 2147483647 ? "MAX" : sqlDataColumnSize.ToString()) + ")";
            }
            else if (sqlDataTypeName == "DECIMAL" || sqlDataTypeName == "NUMERIC")
            {
                return sqlDataTypeName + "(" + NumericPrecision + "," + NumericScale + ")";
            }
            else if (sqlDataTypeName == "DATETIME2")
            {
                return sqlDataTypeName + "(" + NumericScale + ")";
            }


            //else if (sqlDataTypeName == "DECIMAL")
            //    sqlDataTypeName = sqlDataTypeName;
            //TODO ... list all other types that need additional handling 

            return sqlDataTypeName;
        }

        public static IGridControl GetFocusGridControl()
        {
            object outVsWindowFrame = null;
            ServiceCache.VSMonitorSelection.GetCurrentElementValue((int)VSConstants.VSSELELEMID.SEID_WindowFrame, out outVsWindowFrame);

            var vsWindowFrame = outVsWindowFrame as IVsWindowFrame;
            vsWindowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out var outControl);

            var control = (Control)outControl;
            return (GridControl)((ContainerControl)((ContainerControl)control).ActiveControl).ActiveControl;
        }

        public static List<DataTable> GetDataTables()
        {
            List<DataTable> dataTables = new List<DataTable>();

            try
            {
                var SQLResultsControl = GetSQLResultsControl();
                if (SQLResultsControl is null)
                {
                    return dataTables;
                }

                var m_gridResultsPage = GetNonPublicField(SQLResultsControl, "m_gridResultsPage");
                if (m_gridResultsPage is null)
                {
                    return dataTables;
                }

                var gridContainersObj = GetNonPublicField(m_gridResultsPage, "m_gridContainers");
                if (gridContainersObj is null)
                {
                    return dataTables;
                }

                // Handle both SSMS 21 (CollectionBase) and SSMS 22 (List<T>)
                if (!(gridContainersObj is IEnumerable gridContainers))
                {
                    _logger?.Error($"Unexpected grid containers type: {gridContainersObj.GetType().FullName}");
                    return dataTables;
                }

                foreach (var gridContainer in gridContainers)
                {
                    if (!(GetNonPublicField(gridContainer, "m_grid") is GridControl grid))
                    {
                        continue;
                    }

                    var gridStorage = grid.GridStorage;
                    if (gridStorage is null)
                    {
                        continue;
                    }

                    if (!(GetNonPublicField(gridStorage, "m_schemaTable") is DataTable schemaTable))
                    {
                        continue;
                    }

                    List<int> columnSizes = new List<int>();
                    if (GetNonPublicField(grid, "m_Columns") is GridColumnCollection gridColumns)
                    {
                        foreach (GridColumn gridColumn in gridColumns)
                        {
                            columnSizes.Add(gridColumn.WidthInPixels);
                        }
                    }

                    var data = new DataTable();

                    for (int c = 0; c < schemaTable.Rows.Count; c++)
                    {
                        string columnNameInt = "Column_" + c.ToString();
                        var columnType = schemaTable.Rows[c][12] as Type;

                        if (columnType == typeof(Guid)
                                || columnType == typeof(DateTime)
                                || columnType == typeof(DateTimeOffset)
                                || columnType == typeof(byte[]))
                        {
                            columnType = typeof(string);
                        }

                        string sqlDataTypeName = GetColumnSqlType(schemaTable.Rows[c]);

                        DataColumn newColumn = data.Columns.Add(columnNameInt, columnType);

                        var columnName = schemaTable.Rows[c][0].ToString();
                        if (string.IsNullOrEmpty(columnName))
                        {
                            columnName = columnNameInt;
                        }

                        newColumn.ExtendedProperties.Add("columnName", columnName);
                        newColumn.ExtendedProperties.Add("sqlType", sqlDataTypeName);
                        if (columnSizes.Count > c + 1)
                        {
                            newColumn.ExtendedProperties.Add("columnWidthInPixels", columnSizes[c + 1]);
                        }
                    }

                    for (long i = 0; i < gridStorage.NumRows(); i++)
                    {
                        var rowItems = new List<object>();

                        for (int c = 0; c < schemaTable.Rows.Count; c++)
                        {
                            var columnType = schemaTable.Rows[c][12] as Type;
                            var cellData = gridStorage.GetCellDataAsString(i, c + 1);

                            if (cellData == "NULL")
                            {
                                rowItems.Add(null);
                                continue;
                            }

                            if (columnType == typeof(bool))
                            {
                                cellData = cellData == "0" ? "False" : "True";
                            }

                            // leave some types as strings because the conversion from string fails
                            if (columnType == typeof(Guid)
                                || columnType == typeof(TimeSpan)
                                || columnType == typeof(DateTime)
                                || columnType == typeof(DateTimeOffset)
                                || columnType == typeof(byte[]))
                            {
                                columnType = typeof(string);
                            }

                            var typedValue = Convert.ChangeType(cellData, columnType, CultureInfo.InvariantCulture);

                            rowItems.Add(typedValue);
                        }

                        data.Rows.Add(rowItems.ToArray());
                    }

                    data.AcceptChanges();

                    dataTables.Add(data);
                }
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in GetDataTables");
            }

            return dataTables;
        }
    }
}