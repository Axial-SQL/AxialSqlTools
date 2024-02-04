using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace AxialSqlTools
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ExportGridToAsInsertsCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4138;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExportGridToAsInsertsCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ExportGridToAsInsertsCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ExportGridToAsInsertsCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ExportGridToAsInsertsCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new ExportGridToAsInsertsCommand(package, commandService);
        }

        public object GetNonPublicField(object obj, string field)
        {
            FieldInfo f = obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance);

            return f.GetValue(obj);
        }
        public FieldInfo GetNonPublicFieldInfo(object obj, string field)
        {
            FieldInfo f = obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance);

            return f;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();


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



                StringBuilder buffer = new StringBuilder();
                // Define new temp table

                StringBuilder columnList = new StringBuilder();
                for (int i = 0; i < data.Columns.Count; i++)
                {
                    if (i > 0) columnList.Append(", ");
                    columnList.AppendFormat("[{0}] SQL_VARIANT", data.Columns[i].ColumnName);
                }

                buffer.AppendFormat("CREATE TABLE #tempBuffer ({0});\nGO", columnList.ToString());
                buffer.AppendLine("");

                string prefix = "INSERT INTO #tempBuffer VALUES";

                // generate INSERT statements
                int j = 1;
                foreach (DataRow row in data.Rows)
                {
                    StringBuilder values = new StringBuilder();
                    values.Append("(");
                    for (int i = 0; i < data.Columns.Count; i++)
                    {
                        if (i > 0) values.Append(", ");

                        Type dataType = data.Columns[i].DataType;

                        if (row.IsNull(i))
                        {
                            values.Append("NULL");
                        }
                        else if (dataType == typeof(int) ||
                            dataType == typeof(decimal) ||
                            dataType == typeof(long) ||
                            dataType == typeof(double) ||
                            dataType == typeof(float) ||
                            dataType == typeof(byte))
                        {
                            values.Append(row[i].ToString());
                        }
                        else if (dataType == typeof(byte[]))
                        {
                            values.Append("0x");
                            foreach (byte b in (byte[])row[i])
                            {
                                values.Append(b.ToString("x2"));
                            }
                        }
                        else
                        {
                            values.AppendFormat("N'{0}'", row[i].ToString().Replace("'", "''"));
                        }
                    }
                    values.AppendFormat(")");

                    if (j == 1)
                    {
                        buffer.AppendLine(prefix);
                    }
                    if (j > 1)
                    {
                        buffer.Append(",");
                        buffer.AppendLine("");
                    }

                    buffer.Append(values.ToString());

                    if (j % 100 == 0)
                    {
                        buffer.AppendLine(";");
                        buffer.AppendLine("GO");
                        j = 0;
                    }

                    j += 1;

                }

                if (j % 100 == 1) buffer.AppendLine("GO");
                buffer.AppendLine("SELECT * FROM #tempBuffer;");

                UIConnectionInfo connection = ServiceCache.ScriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo;

                ServiceCache.ScriptFactory.CreateNewBlankScript(ScriptType.Sql, connection, null);

                // insert SQL definition to document
                EnvDTE.TextDocument doc = (EnvDTE.TextDocument)ServiceCache.ExtensibilityModel.Application.ActiveDocument.Object(null);

                doc.EndPoint.CreateEditPoint().Insert(buffer.ToString());

            }
        }
    }
}
