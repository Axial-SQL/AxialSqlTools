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
        private static object GetNonPublicField(object obj, string field)
        {
            FieldInfo f = obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance);

            return f.GetValue(obj);
        }
        private static FieldInfo GetNonPublicFieldInfo(object obj, string field)
        {
            FieldInfo f = obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance);

            return f;
        }

        public static List<DataTable> GetDataTables()
        {
            List<DataTable> dataTables = new List<DataTable>();

            var objType = ServiceCache.ScriptFactory.GetType();
            var method1 = objType.GetMethod("GetCurrentlyActiveFrameDocView", BindingFlags.NonPublic | BindingFlags.Instance);
            var Result = method1.Invoke(ServiceCache.ScriptFactory, new object[] { ServiceCache.VSMonitorSelection, false, null });


            var objType2 = Result.GetType();
            var field = objType2.GetField("m_sqlResultsControl", BindingFlags.NonPublic | BindingFlags.Instance);
            var SQLResultsControl = field.GetValue(Result);


            var m_gridResultsPage = GetNonPublicField(SQLResultsControl, "m_gridResultsPage");
            CollectionBase gridContainers = GetNonPublicField(m_gridResultsPage, "m_gridContainers") as CollectionBase;

            foreach (var gridContainer in gridContainers)
            {
                var grid = GetNonPublicField(gridContainer, "m_grid") as GridControl;
                var gridStorage = grid.GridStorage;
                var schemaTable = GetNonPublicField(gridStorage, "m_schemaTable") as DataTable;

                //// TEST----
                //var gridColumns = GetNonPublicField(grid, "m_Columns") as GridColumnCollection;
                //if (gridColumns != null)
                //{
                //    //var gridColumnCollection = columnsProperty.GetValue(grid) as IEnumerable; // Replace IEnumerable with the actual type of the Columns collection

                //    // Iterate through the gridColumnCollection to access individual GridColumn objects

                //    grid.GridColumnsInfo.Clear();

                //    foreach (GridColumnInfo gridColumn in grid.GridColumnsInfo)
                //    {

                //        //var textAlignField = GetNonPublicFieldInfo(gridColumn, "ColumnAlignment");
                //        //if (textAlignField != null)
                //        //{
                //        //    textAlignField.SetValue(gridColumn, System.Windows.Forms.HorizontalAlignment.Right);
                //        //}

                //        var id = grid.GridColumnsInfo.IndexOf(gridColumn);
                //        grid.GridColumnsInfo.Remove(gridColumn);

                //        gridColumn.ColumnAlignment = HorizontalAlignment.Right;                        

                //        grid.GridColumnsInfo.Add(gridColumn);

                //    }

                //    //foreach (var gridColumn in gridColumns)
                //    //{
                //    //    var textAlignField = GetNonPublicFieldInfo(gridColumn, "TextAlign");
                //    //    if (textAlignField != null)
                //    //    {
                //    //        textAlignField.SetValue(gridColumn, System.Windows.Forms.HorizontalAlignment.Right);
                //    //    }
                //    //    var textAlignField2 = GetNonPublicFieldInfo(gridColumn, "m_myAlign");
                //    //    if (textAlignField2 != null)
                //    //    {
                //    //        textAlignField2.SetValue(gridColumn, System.Windows.Forms.HorizontalAlignment.Right);
                //    //    }
                //    //}
                //}
                //grid.Refresh();
                //// TEST----

                var data = new DataTable();

                for (long i = 0; i < gridStorage.NumRows(); i++)
                {
                    var rowItems = new List<object>();

                    for (int c = 0; c < schemaTable.Rows.Count; c++)
                    {
                        var columnName = schemaTable.Rows[c][0].ToString();
                        var columnType = schemaTable.Rows[c][12] as Type;

                        if (!data.Columns.Contains(columnName))
                        {
                            data.Columns.Add(columnName, columnType);
                        }

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

                        //Console.WriteLine($"Parsing {columnName} with '{cellData}'");

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
