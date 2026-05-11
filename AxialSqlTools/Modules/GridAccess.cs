using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
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

        public static string TryGetStatisticsMessagesText(object executionContext = null)
        {
            var visited = new HashSet<object>(new ReferenceEqualityComparer());

            if (TryFindStatisticsText(executionContext, 0, visited, out var statisticsText))
            {
                return statisticsText;
            }

            var sqlResultsControl = GetSQLResultsControl();
            if (!ReferenceEquals(sqlResultsControl, executionContext)
                && TryFindStatisticsText(sqlResultsControl, 0, visited, out statisticsText))
            {
                return statisticsText;
            }

            return null;
        }

        public static void TryFlushStatisticsMessages()
        {
            var sqlResultsControl = GetSQLResultsControl();
            if (sqlResultsControl is null)
            {
                return;
            }

            TryInvokeParameterlessDelegateField(sqlResultsControl, "m_FlushTextWritersInvoker");
            TryInvokeParameterlessMethod(sqlResultsControl, "FlushTextWriters");

            TryFlushWriterLike(GetNonPublicField(sqlResultsControl, "m_messagesWriter"));
            TryFlushWriterLike(GetNonPublicField(sqlResultsControl, "m_resultsWriter"));
            TryFlushWriterLike(GetNonPublicField(sqlResultsControl, "m_errorsWriter"));
        }

        private static bool TryFindStatisticsText(object candidate, int depth, HashSet<object> visited, out string statisticsText)
        {
            statisticsText = null;

            if (candidate is null || depth > 4)
            {
                return false;
            }

            if (candidate is string candidateText)
            {
                if (LooksLikeStatisticsOutput(candidateText))
                {
                    statisticsText = candidateText;
                    return true;
                }

                return false;
            }

            if (candidate is TextWriter textWriter)
            {
                var writerText = SafeToString(textWriter);
                if (LooksLikeStatisticsOutput(writerText))
                {
                    statisticsText = writerText;
                    return true;
                }
            }

            if (candidate is IVsTextView textView
                && TryGetTextFromVsTextView(textView, out var textViewText)
                && LooksLikeStatisticsOutput(textViewText))
            {
                statisticsText = textViewText;
                return true;
            }

            if (candidate is IVsTextLines textLines
                && TryGetTextFromVsTextLines(textLines, out var textLinesText)
                && LooksLikeStatisticsOutput(textLinesText))
            {
                statisticsText = textLinesText;
                return true;
            }

            var candidateType = candidate.GetType();
            if (candidateType.IsPrimitive
                || candidate is decimal
                || candidate is DateTime
                || candidate is TimeSpan
                || candidate is Enum)
            {
                return false;
            }

            if (!candidateType.IsValueType && !visited.Add(candidate))
            {
                return false;
            }

            if (candidate is Control control)
            {
                if (LooksLikeStatisticsOutput(control.Text))
                {
                    statisticsText = control.Text;
                    return true;
                }

                foreach (Control childControl in control.Controls)
                {
                    if (TryFindStatisticsText(childControl, depth + 1, visited, out statisticsText))
                    {
                        return true;
                    }
                }
            }

            foreach (var member in GetRelevantMembers(candidate))
            {
                if (TryFindStatisticsText(member.Value, depth + 1, visited, out statisticsText))
                {
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<KeyValuePair<string, object>> GetRelevantMembers(object obj)
        {
            foreach (var field in EnumerateInstanceFields(obj.GetType()))
            {
                if (!IsRelevantMember(field.Name, field.FieldType))
                {
                    continue;
                }

                object value = null;
                try
                {
                    value = field.GetValue(obj);
                }
                catch
                {
                }

                if (value != null)
                {
                    yield return new KeyValuePair<string, object>(field.Name, value);
                }
            }

            var seenProperties = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in EnumerateInstanceProperties(obj.GetType()))
            {
                if (!property.CanRead || property.GetIndexParameters().Length > 0 || !IsRelevantMember(property.Name, property.PropertyType))
                {
                    continue;
                }

                if (!seenProperties.Add(property.Name))
                {
                    continue;
                }

                object value = null;
                try
                {
                    value = property.GetValue(obj);
                }
                catch
                {
                }

                if (value != null)
                {
                    yield return new KeyValuePair<string, object>(property.Name, value);
                }
            }
        }

        private static IEnumerable<FieldInfo> EnumerateInstanceFields(Type type)
        {
            const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly;

            for (var currentType = type; currentType != null; currentType = currentType.BaseType)
            {
                foreach (var field in currentType.GetFields(bindingFlags))
                {
                    yield return field;
                }
            }
        }

        private static IEnumerable<PropertyInfo> EnumerateInstanceProperties(Type type)
        {
            const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly;

            for (var currentType = type; currentType != null; currentType = currentType.BaseType)
            {
                foreach (var property in currentType.GetProperties(bindingFlags))
                {
                    yield return property;
                }
            }
        }

        private static bool IsRelevantMember(string memberName, Type memberType)
        {
            return LooksLikeRelevantMemberName(memberName) || IsRelevantMemberType(memberType);
        }

        private static bool IsRelevantMemberType(Type memberType)
        {
            if (memberType is null)
            {
                return false;
            }

            return typeof(Control).IsAssignableFrom(memberType)
                || typeof(TextWriter).IsAssignableFrom(memberType)
                || typeof(IVsTextView).IsAssignableFrom(memberType)
                || typeof(IVsTextLines).IsAssignableFrom(memberType);
        }

        private static bool LooksLikeRelevantMemberName(string memberName)
        {
            if (string.IsNullOrEmpty(memberName))
            {
                return false;
            }

            memberName = memberName.ToLowerInvariant();
            return memberName.Contains("message")
                || memberName.Contains("text")
                || memberName.Contains("writer")
                || memberName.Contains("buffer")
                || memberName.Contains("document")
                || memberName.Contains("editor")
                || memberName.Contains("view")
                || memberName.Contains("output")
                || memberName.Contains("result")
                || memberName.Contains("pane")
                || memberName.Contains("page")
                || memberName.Contains("sql")
                || memberName.Contains("exec");
        }

        private static string SafeToString(object value)
        {
            try
            {
                return value?.ToString();
            }
            catch
            {
                return null;
            }
        }


        private static void TryFlushWriterLike(object writer)
        {
            if (writer is null)
            {
                return;
            }

            if (writer is TextWriter textWriter)
            {
                try
                {
                    textWriter.Flush();
                }
                catch
                {
                }
            }

            TryInvokeParameterlessMethod(writer, "Flush");
        }

        private static bool TryInvokeParameterlessDelegateField(object target, string fieldName)
        {
            if (target is null || string.IsNullOrWhiteSpace(fieldName))
            {
                return false;
            }

            try
            {
                if (GetNonPublicField(target, fieldName) is Delegate callback)
                {
                    callback.DynamicInvoke();
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool TryInvokeParameterlessMethod(object target, string methodName)
        {
            if (target is null || string.IsNullOrWhiteSpace(methodName))
            {
                return false;
            }

            try
            {
                var method = target.GetType().GetMethod(methodName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, Type.EmptyTypes, null);
                if (method == null)
                {
                    return false;
                }

                method.Invoke(target, null);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetTextFromVsTextView(IVsTextView textView, out string text)
        {
            text = null;

            try
            {
                if (textView == null || textView.GetBuffer(out IVsTextLines textLines) != VSConstants.S_OK)
                {
                    return false;
                }

                return TryGetTextFromVsTextLines(textLines, out text);
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetTextFromVsTextLines(IVsTextLines textLines, out string text)
        {
            text = null;

            try
            {
                if (textLines == null || textLines.GetLastLineIndex(out var lastLine, out var lastIndex) != VSConstants.S_OK)
                {
                    return false;
                }

                if (textLines.GetLineText(0, 0, lastLine, lastIndex, out var fullText) != VSConstants.S_OK)
                {
                    return false;
                }

                text = fullText;
                return !string.IsNullOrWhiteSpace(text);
            }
            catch
            {
                return false;
            }
        }

        public static bool LooksLikeStatisticsOutput(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return text.IndexOf("SQL Server Execution Times:", StringComparison.OrdinalIgnoreCase) >= 0
                || (text.IndexOf("logical reads", StringComparison.OrdinalIgnoreCase) >= 0
                    && text.IndexOf("Scan count", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
        {
            public new bool Equals(object x, object y) => ReferenceEquals(x, y);

            public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
        }
    }
}