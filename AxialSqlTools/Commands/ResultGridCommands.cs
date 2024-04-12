using EnvDTE;
using EnvDTE80;
using Microsoft.SqlServer.Management.UI.Grid;
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

    public interface IResultGridControl
    {
        string GetQueryText();

        object GetGridControl();

        CollectionBase GetGridContainers();

        void ChangeWindowTitle(string text);

        IGridControl GetFocusGridControl();
    }


    internal abstract class ResultGridCommandBase
    {
        private const string SQL_RESULT_GRID_CONTEXT_NAME = "SQL Results Grid Tab Context";
        protected static CommandBar GridCommandBar { get; }

        //protected static readonly IResultGridControl GridControl;

        static ResultGridCommandBase()
        {
            //GridControl = new ResultGridControl();
            var dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
            GridCommandBar = ((CommandBars)dte.CommandBars)[SQL_RESULT_GRID_CONTEXT_NAME];
        }
    }

    internal sealed class ResultGridCopyAsInsertCommand : ResultGridCommandBase
    {
        private static string templates;
        private readonly AsyncPackage package;
        //private static readonly IClipboardService _clipboardService;

        static ResultGridCopyAsInsertCommand()
        {
            // _clipboardService = new ClipboardService();
        }

        public static async Task InitializeAsync(AxialSqlToolsPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            // templates = File.ReadAllText($"{sqlAsyncPackage.ExtensionInstallationDirectory}/Templates/SQL.INSERT.INTO.sql");
            // GridCommandBar.AddButton("Copy As #INSERT", "", OnClick);

            var caption = "Copy As #INSERT";

            var btnControl = (CommandBarButton)GridCommandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControl.Visible = true;
            btnControl.Caption = caption;
            btnControl.Click += OnClick;

        }

        private static void OnClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Show a message box to prove we were here

            MessageBox.Show("OnClick", "Done");

            //var focusGridControl = GridControl.GetFocusGridControl();
            //using (var gridResultControl = new ResultGridControlAdaptor(focusGridControl))
            //{
            //    var gridResultSelected = gridResultControl.GridSelectedAsQuerySql();
            //    var columnHeaders = gridResultSelected.ElementAt(0);
            //    var contentRows = gridResultSelected.Skip(1);

            //    var rows = string.Join($",{Environment.NewLine}\t", contentRows.Select(r => $"({string.Join(", ", r)})"));
            //    var sqlQuery = templates.Replace("{columnHeaders}", columnHeaders).Replace("{rows}", rows);

            //    _clipboardService.Set(sqlQuery);
            //    ServiceCache.ExtensibilityModel.StatusBar.Text = "Copied";
            //}

        }
    }
}
