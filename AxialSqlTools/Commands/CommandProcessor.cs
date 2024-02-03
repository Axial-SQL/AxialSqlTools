// Copyright (C) 2006-2010 Jim Tilander. See COPYING for and README for more details.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Reflection;
using EnvDTE;
using EnvDTE80;
using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using AxialSqlTools;
using System.Windows.Input;

namespace Aurora
{
    class CommandProcessor : CommandBase
    {
        public CommandProcessor(Plugin plugin, string canonicalName, string buttonText, string toolTip)
            : base(buttonText, canonicalName, plugin, toolTip)
        {
        }

        public AsyncPackage package;
        public string FullFileName;

        override public int ScriptIconIndex { get { return 0; } }

        public override bool OnCommand()
        {
            ThreadHelper.ThrowIfNotOnUIThread();


            try
            {

                if (System.IO.File.Exists(FullFileName))
                {
                    string fileContent = System.IO.File.ReadAllText(FullFileName);

                    DTE dte = Package.GetGlobalService(typeof(DTE)) as DTE;

                    bool isShiftPressed = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);

                    if (isShiftPressed || dte?.ActiveDocument == null) {
                        ServiceCache.ScriptFactory.CreateNewBlankScript(ScriptType.Sql);
                    }

                    if (dte?.ActiveDocument != null)
                    {
                        TextSelection selection = dte.ActiveDocument.Selection as TextSelection;
                        selection.Delete();
                        selection.Insert(fileContent.Trim());
                    }
                }
                else
                {
                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        "File " + FullFileName + " doesn't exist!",
                        "Error",
                        OLEMSGICON.OLEMSGICON_WARNING,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }
            }
            catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    ex.Message,
                    "Error",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }

            return true;

        }


        public override bool IsEnabled()
        {
            return true;
        }
    }

}
