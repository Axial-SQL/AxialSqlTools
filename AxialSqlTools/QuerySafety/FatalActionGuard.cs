using EnvDTE;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Windows.Interop;

namespace AxialSqlTools.QuerySafety
{
    internal static class FatalActionGuard
    {
        // Keep the COM document wrapper stable until WindowClosing, then release it.
        private static readonly Dictionary<object, FatalActionSession> sessions = new Dictionary<object, FatalActionSession>();
        private static bool dialogOpen;

        internal static void ResetApprovals() { sessions.Clear(); }
        internal static void ForgetDocument(object document) { if (document != null) sessions.Remove(document); }

        internal static bool ShouldCancel(DTE dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (dialogOpen) return true;
            if (!SettingsManager.GetWarnWhenRunningFatalAction()) return false;

            try
            {
                var document = dte?.ActiveDocument;
                if (document == null) return false;
                var textDocument = document.Object("TextDocument") as TextDocument;
                var selection = document.Selection as TextSelection;
                if (textDocument == null || selection == null)
                    throw new InvalidOperationException("The selected query text is unavailable.");

                // Query.Execute runs the selection when present, otherwise the whole editor buffer.
                string sql = selection.IsEmpty ? textDocument.StartPoint.CreateEditPoint().GetText(textDocument.EndPoint) : selection.Text;
                string context;
                string connection = ConnectionKey(dte, out context);
                if (!sessions.TryGetValue(document, out FatalActionSession session))
                {
                    session = new FatalActionSession();
                    sessions.Add(document, session);
                }
                if (session.IsApproved(sql, connection)) return false;

                var analysis = FatalActionAnalyzer.Analyze(sql);
                if (!analysis.RequiresConfirmation) return false;
                bool canRemember = connection != null && analysis.Limitation == null;
                bool allowRepeats;
                bool allow = Confirm(analysis, context, canRemember, out allowRepeats);
                if (allow && canRemember && allowRepeats) session.Approve(sql, connection);
                return !allow;
            }
            catch (Exception ex)
            {
                AxialSqlToolsPackage._logger?.Error(ex, "Could not check query execution for fatal actions.");
                bool ignored;
                return !Confirm(new FatalActionAnalysis {
                    Limitation = "The query could not be checked. Review it before running. " + ex.Message
                }, "Query check unavailable", false, out ignored);
            }
        }

        private static bool Confirm(FatalActionAnalysis analysis, string context, bool canRemember, out bool allowRepeats)
        {
            allowRepeats = false;
            dialogOpen = true;
            try
            {
                var dialog = new FatalActionWarningDialog(analysis, context, canRemember);
                var shell = Package.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
                if (shell != null && shell.GetDialogOwnerHwnd(out var hwnd) == 0 && hwnd != IntPtr.Zero)
                    new WindowInteropHelper(dialog).Owner = hwnd;
                bool allow = dialog.ShowDialog() == true;
                allowRepeats = allow && dialog.AllowRepeats;
                return allow;
            }
            catch (Exception ex)
            {
                // A failure to show a confirmation must not authorize execution.
                AxialSqlToolsPackage._logger?.Error(ex, "Could not display the fatal action warning. Execution cancelled.");
                VsShellUtilities.ShowMessageBox(Microsoft.VisualStudio.Shell.ServiceProvider.GlobalProvider,
                    "Query execution was cancelled because its warning could not be displayed. " + ex.Message,
                    "Axial SQL Tools", OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return false;
            }
            finally { dialogOpen = false; }
        }

        private static string ConnectionKey(DTE dte, out string description)
        {
            description = "Active query window (line numbers are relative to the executed selection or script)";
            try
            {
                var ui = ServiceCache.ScriptFactory?.CurrentlyActiveWndConnectionInfo?.UIConnectionInfo;
                string server = ui?.ServerName;
                string database = ui?.AdvancedOptions?.Get("DATABASE");
                if (!string.IsNullOrWhiteSpace(server))
                    description = "Server: " + server + " | Starting database: " + (database ?? "default") +
                        "\nLine numbers are relative to the executed selection or script.";

                // Reuse the editor's connection. Never create a connection or execute a probe query.
                object connection = GridAccess.GetNonPublicField(dte.ActiveWindow?.Object, "m_connection");
                if (connection == null)
                {
                    object executor = GridAccess.GetNonPublicField(GridAccess.GetSQLResultsControl(), "m_sqlExec");
                    connection = GridAccess.GetNonPublicField(executor, "m_conn");
                }
                if (!string.Equals(GridAccess.GetProperty(connection, "State")?.ToString(), "Open", StringComparison.Ordinal)) return null;
                if (!(GridAccess.GetProperty(connection, "ClientConnectionId") is Guid id) || id == Guid.Empty) return null;
                string actualServer = GridAccess.GetProperty(connection, "DataSource") as string;
                string actualDatabase = GridAccess.GetProperty(connection, "Database") as string;
                // Include both the live database and the toolbar selection, which can change before execution.
                if (string.IsNullOrWhiteSpace(actualServer) || string.IsNullOrWhiteSpace(actualDatabase) ||
                    string.IsNullOrWhiteSpace(server) || string.IsNullOrWhiteSpace(database)) return null;
                return id.ToString("D") + "\n" + actualServer + "\n" + actualDatabase + "\n" + server + "\n" + database;
            }
            catch
            {
                // SSMS private APIs vary. Checking SQL still works, but repeat approval needs a known session.
                return null;
            }
        }
    }
}
