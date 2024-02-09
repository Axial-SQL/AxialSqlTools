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

            List<DataTable> dataTables = GridAccess.GetDataTables();

            int jj = 0;

            StringBuilder globalBuffer = new StringBuilder();

            foreach (DataTable dataTable in dataTables)
            {

                StringBuilder buffer = new StringBuilder();
                // Define new temp table

                if (jj > 0) buffer.AppendLine("-------------------------------------------");

                StringBuilder columnList = new StringBuilder();
                int ii = 0;
                foreach (DataColumn column in dataTable.Columns)
                {
                    if (ii > 0)
                    {
                        columnList.Append(", \n");
                        columnList.Append("         ");
                    }
                    string columnType = "SQL_VARIANT";
                    if (column.ExtendedProperties.ContainsKey("sqlType"))
                        columnType = (string)column.ExtendedProperties["sqlType"];

                    string columnName = column.ColumnName;
                    if (column.ExtendedProperties.ContainsKey("columnName"))
                        columnName = (string)column.ExtendedProperties["columnName"];

                    columnList.AppendFormat("[{0}] {1}", columnName, columnType);

                    ii += 1;
                }

                buffer.AppendLine("IF OBJECT_ID('tempdb..#tempBuffer') IS NOT NULL DROP TABLE #tempBuffer;");
                buffer.AppendLine("GO");
                // TODO - probably should format it too
                buffer.AppendFormat("CREATE TABLE #tempBuffer ({0});\n", columnList.ToString());
                buffer.AppendLine("GO");
                buffer.AppendLine("-------------------------------------------");

                string prefix = "INSERT INTO #tempBuffer VALUES";

                // generate INSERT statements
                int j = 1;
                foreach (DataRow row in dataTable.Rows)
                {
                    StringBuilder values = new StringBuilder();
                    values.Append("(");
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (i > 0) values.Append(", ");

                        Type dataType = dataTable.Columns[i].DataType;

                        if (row.IsNull(i))
                        {
                            values.Append("NULL");
                        }
                        else if (dataType == typeof(bool))
                        {
                            values.Append((bool)row[i] ? 1 : 0);
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

                // the last batch was less than 100 records
                if (j % 100 > 1)
                {
                    buffer.AppendLine(";");
                    buffer.AppendLine("GO");
                }

                buffer.AppendLine("SELECT * FROM #tempBuffer;");

                globalBuffer.Append(buffer.ToString());

                jj += 1;

            }

            UIConnectionInfo connection = ServiceCache.ScriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo;

            ServiceCache.ScriptFactory.CreateNewBlankScript(ScriptType.Sql, connection, null);

            // insert SQL definition to document
            EnvDTE.TextDocument doc = (EnvDTE.TextDocument)ServiceCache.ExtensibilityModel.Application.ActiveDocument.Object(null);

            doc.EndPoint.CreateEditPoint().Insert(globalBuffer.ToString());

        }
    }
}
