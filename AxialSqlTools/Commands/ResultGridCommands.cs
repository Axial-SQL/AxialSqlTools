﻿using EnvDTE;
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
        private static string templates;
        private readonly AsyncPackage package;

        static ResultGridCopyAsInsertCommand()
        {
            
        }

        public static async Task InitializeAsync(AxialSqlToolsPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var btnControlAsInsert = (CommandBarButton)GridCommandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlAsInsert.Visible = true;
            btnControlAsInsert.Caption = "Copy As #INSERT";
            btnControlAsInsert.Click += OnClick_CopyAsInsert;

            var btnControlAsCsv = (CommandBarButton)GridCommandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            btnControlAsCsv.Visible = true;
            btnControlAsCsv.Caption = "Copy As CSV";
            btnControlAsCsv.Click += OnClick_CopyAsCSV;

        }

        private static void OnClick_CopyAsInsert(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            CopySelectedValues("SQLINSERT");
        }

        private static void OnClick_CopyAsCSV(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            CopySelectedValues("CSV");
        }

        
        private static void CopySelectedValues (string CopyType)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var focusGridControl = GridAccess.GetFocusGridControl();
            using (var gridResultControl = new ResultGridControlAdaptor(focusGridControl))
            {      
                if (CopyType == "SQLINSERT")
                {
                    var gridResultSelected = gridResultControl.GridSelectedAsQuerySql();
                    var columnHeaders = gridResultSelected.ElementAt(0);
                    var contentRows = gridResultSelected.Skip(1);

                    var resultText = string.Join(",\r\n", contentRows.Select(r => $"({string.Join(", ", r)})"));
                    resultText = resultText.TrimEnd(',');

                    resultText = $"INSERT INTO [table] ({columnHeaders}) VALUES\r\n" + resultText;

                    SetClipboardText(resultText);

                }
                else if(CopyType == "CSV")
                {
                    // TODO - this should copy selected cells only
                    var datatable = gridResultControl.GridFocusAsDatatable();
                    var resultText = datatable.ToCsv();

                    SetClipboardText(resultText);

                }

                ServiceCache.ExtensibilityModel.StatusBar.Text = "Copied";
            }

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
