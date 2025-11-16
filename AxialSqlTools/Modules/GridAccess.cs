using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;
using System.Windows.Forms;

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
            if (obj == null) return null;
            return obj.GetType().GetProperty(field, BindingFlags.Public | BindingFlags.Instance)?.GetValue(obj);
        }
        public static object GetNonPublicField(object obj, string field)
        {
            if (obj == null) return null;
            return obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(obj);
        }
        public static FieldInfo GetNonPublicFieldInfo(object obj, string field)
        {
            return obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance);
        }

        public static object GetSQLResultsControl()
        {
            var factoryType = ServiceCache.ScriptFactory.GetType();
            var method = factoryType.GetMethod(
                "GetCurrentlyActiveFrameDocView",
                BindingFlags.NonPublic | BindingFlags.Instance);

            var docView = method.Invoke(
                ServiceCache.ScriptFactory,
                new object[] { ServiceCache.VSMonitorSelection, false, null });

            // It might be ObjectExplorerTool, a designer, etc.
            var scriptEditor = docView as SqlScriptEditorControl;
            if (scriptEditor == null)
            {
                // No active query window - handle as you like
                // e.g. return null, or throw a more descriptive exception
                return null;
            }

            var field = typeof(SqlScriptEditorControl)
                .GetField("m_sqlResultsControl",
                    BindingFlags.NonPublic | BindingFlags.Instance);

            var sqlResultsControl = field.GetValue(scriptEditor);
            return sqlResultsControl;
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
            var SQLResultsControl = GetSQLResultsControl();

            var m_gridResultsPage = GetNonPublicField(SQLResultsControl, "m_gridResultsPage");
            CollectionBase gridContainers = GetNonPublicField(m_gridResultsPage, "m_gridContainers") as CollectionBase;

            return gridContainers;
        }

        public static void ChangeStatusBarContent(int OpenTranCount, string ActualElapsedTime)
        {
            QEStatusBarManager statusBarManager = GetStatusBarManager();

            if (OpenTranCount > 0 )
            {
                var msg = "One transaction is still open!";
                if (OpenTranCount > 1)
                    msg = $"{OpenTranCount} transactions are still open!";

                var currentMsg = statusBarManager.StatusText;
                statusBarManager.SetKnownState(QEStatusBarKnownStates.Executing);
                statusBarManager.StatusText = currentMsg + " | " + msg;                

            }

            var statusBarManager_executionTimePanel = GetNonPublicField(statusBarManager, "executionTimePanel");
            SetPropertyValue(statusBarManager_executionTimePanel, "Text", ActualElapsedTime);

            var statusBarManager_completedTimePanel = GetNonPublicField(statusBarManager, "completedTimePanel");
            //statusBarManager_completedTimePanel.Visible = true;

            // Can't express enough how much I don't like this...
            var generalPanel = GetNonPublicField(statusBarManager, "generalPanel");
            if (OpenTranCount > 0)
                SetPropertyValue(generalPanel, "ForeColor", Color.Red);
            else
                SetPropertyValue(generalPanel, "ForeColor", Color.Black);
                        
            //TODO - need to contract font from existing property..
            Font defaultFont = new Font("Segoe UI", 9);
            Font boldFont = new Font("Segoe UI", 10, FontStyle.Bold);

            var statusStrip = GetNonPublicField(statusBarManager, "statusStrip");
            if (OpenTranCount > 0)
                SetPropertyValue(statusStrip, "Font", boldFont);
            else
                SetPropertyValue(statusStrip, "Font", defaultFont);            
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
                sqlDataTypeName = sqlDataTypeName + "(" + (sqlDataColumnSize == 2147483647 ? "MAX" : sqlDataColumnSize.ToString()) + ")";
            }
            else if (sqlDataTypeName == "DECIMAL" || sqlDataTypeName == "NUMERIC")
            {
                sqlDataTypeName = sqlDataTypeName + "(" + NumericPrecision + "," + NumericScale + ")";
            }
            else if (sqlDataTypeName == "DATETIME2")
            {
                sqlDataTypeName = sqlDataTypeName + "(" + NumericScale + ")";
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

            CollectionBase gridContainers = GetGridContainers();

            if (gridContainers == null) return dataTables;

            foreach (var gridContainer in gridContainers)
            {
                var grid = GetNonPublicField(gridContainer, "m_grid") as GridControl;
                var gridStorage = grid.GridStorage;
                var schemaTable = GetNonPublicField(gridStorage, "m_schemaTable") as DataTable;

                List<int> columnSizes = new List<int> { };
                var gridColumns = GridAccess.GetNonPublicField(grid, "m_Columns") as GridColumnCollection;
                if (gridColumns != null)
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
                            || columnType == typeof(byte[])
                            )
                        columnType = typeof(string);

                    string sqlDataTypeName = GetColumnSqlType(schemaTable.Rows[c]);

                    DataColumn newColumn = data.Columns.Add(columnNameInt, columnType);

                    var columnName = schemaTable.Rows[c][0].ToString();
                    if (string.IsNullOrEmpty(columnName))
                        columnName = columnNameInt;

                    newColumn.ExtendedProperties.Add("columnName", columnName);
                    newColumn.ExtendedProperties.Add("sqlType", sqlDataTypeName);
                    newColumn.ExtendedProperties.Add("columnWidthInPixels", columnSizes[c + 1]);

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
                            cellData = cellData == "0" ? "False" : "True";

                        // leave some types as strings because the conversation from string fails
                        if (columnType == typeof(Guid) 
                            || columnType == typeof(TimeSpan)
                            || columnType == typeof(DateTime)
                            || columnType == typeof(DateTimeOffset)
                            || columnType == typeof(byte[])
                            )
                            columnType = typeof(string);

                        var typedValue = Convert.ChangeType(cellData, columnType, CultureInfo.InvariantCulture);

                        rowItems.Add(typedValue);
                    }

                    data.Rows.Add(rowItems.ToArray());
                }

                data.AcceptChanges();

                dataTables.Add(data);

            }

            return dataTables;

        }
    }
}
