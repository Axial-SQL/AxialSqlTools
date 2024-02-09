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

        public static List<DataTable> GetDataTables()
        {
            List<DataTable> dataTables = new List<DataTable>();

            CollectionBase gridContainers = GetGridContainers();

            foreach (var gridContainer in gridContainers)
            {
                var grid = GetNonPublicField(gridContainer, "m_grid") as GridControl;
                var gridStorage = grid.GridStorage;
                var schemaTable = GetNonPublicField(gridStorage, "m_schemaTable") as DataTable;

                var data = new DataTable();

                for (int c = 0; c < schemaTable.Rows.Count; c++)
                {
                    string columnNameInt = "Column_" + c.ToString();
                    var columnType = schemaTable.Rows[c][12] as Type;

                    int sqlDataColumnSize = (int)schemaTable.Rows[c][2];
                    string sqlDataTypeName = ((string)schemaTable.Rows[c][24]).ToUpper();

                    DataColumn newColumn = data.Columns.Add(columnNameInt, columnType);

                    if (sqlDataTypeName == "NVARCHAR" || sqlDataTypeName == "NCHAR" || sqlDataTypeName == "VARCHAR" || sqlDataTypeName == "CHAR")
                        sqlDataTypeName = sqlDataTypeName + "(" + sqlDataColumnSize + ")";
                    //else if (sqlDataTypeName == "DECIMAL")
                    //    sqlDataTypeName = sqlDataTypeName;
                    //TODO ... list all other types that need additional handling 

                    var columnName = schemaTable.Rows[c][0].ToString();
                    if (string.IsNullOrEmpty(columnName))
                        columnName = columnNameInt;
                    newColumn.ExtendedProperties.Add("columnName", columnName);
                    newColumn.ExtendedProperties.Add("sqlType", sqlDataTypeName);
                    
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
                        if (columnType == typeof(Guid) || columnType == typeof(DateTimeOffset))
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
