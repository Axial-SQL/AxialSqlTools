using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    public static class GridAccess
    {
        public static object GetNonPublicField(object obj, string field)
        {
            FieldInfo f = obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance);

            return f.GetValue(obj);
        }
        public static FieldInfo GetNonPublicFieldInfo(object obj, string field)
        {
            FieldInfo f = obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance);

            return f;
        }

        public static object GetSQLResultsControl()
        {
            var objType = ServiceCache.ScriptFactory.GetType();
            var method1 = objType.GetMethod("GetCurrentlyActiveFrameDocView", BindingFlags.NonPublic | BindingFlags.Instance);
            var Result = method1.Invoke(ServiceCache.ScriptFactory, new object[] { ServiceCache.VSMonitorSelection, false, null });


            var objType2 = Result.GetType();
            var field = objType2.GetField("m_sqlResultsControl", BindingFlags.NonPublic | BindingFlags.Instance);
            var SQLResultsControl = field.GetValue(Result);

            return SQLResultsControl;
        }

        public static CollectionBase GetGridContainers()
        {
            var SQLResultsControl = GetSQLResultsControl();

            var m_gridResultsPage = GetNonPublicField(SQLResultsControl, "m_gridResultsPage");
            CollectionBase gridContainers = GetNonPublicField(m_gridResultsPage, "m_gridContainers") as CollectionBase;

            return gridContainers;
        }

        public static Microsoft.Data.SqlClient.SqlConnection GetCurrentWindowSqlConnection()
        {
            var SQLResultsControl = GetSQLResultsControl();

            var m_SqlExec = GetNonPublicField(SQLResultsControl, "m_sqlExec");
            Microsoft.Data.SqlClient.SqlConnection m_conn = GetNonPublicField(m_SqlExec, "m_conn") as Microsoft.Data.SqlClient.SqlConnection;

            return m_conn;
        }

        public static void ChangeCurrentWindowTitle(string TransactionMsg)
        {
            var SQLResultsControl = GetSQLResultsControl();

            var m_rawSP = GetNonPublicField(SQLResultsControl, "m_rawSP");
            var frame = GetNonPublicField(m_rawSP, "frame");

            //var rf = frame.GetType().GetProperty("RootFrame");
            //var TitleObj = rv.GetType().GetProperty("Title");

            // doesn't let me change those...
            //var propertyInfo_ToolTip = frame.GetType().GetProperty("ToolTip");
            //var ToolTip = propertyInfo_ToolTip.GetValue(frame, null);


            //var propertyInfo_Title = rf.GetType().GetProperty("Title");
            //var propertyInfo_Title = frame.GetType().GetProperty("AnnotatedTitle");

            //string NewTitle = (string)ToolTip;
            //if (!string.IsNullOrEmpty(TransactionMsg))
            //{
            //    NewTitle = TransactionMsg + NewTitle;
            //}
            //propertyInfo_Title.SetValue(frame, NewTitle);

            var a = 0;

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
            else if (sqlDataTypeName == "DECIMAL" || sqlDataTypeName == "DECIMAL")
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

        public static List<DataTable> GetDataTables()
        {
            List<DataTable> dataTables = new List<DataTable>();

            CollectionBase gridContainers = GetGridContainers();

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
