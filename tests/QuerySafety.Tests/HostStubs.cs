// Stand-ins for the SSMS host and modal UI. Tests run the production guard, analyzer and approval cache.
using System;
using System.Collections.Specialized;
using System.Reflection;
using AxialSqlTools.QuerySafety;

namespace EnvDTE
{
    public class DTE { public Document ActiveDocument { get; set; } public Window ActiveWindow { get; set; } }
    public class Window { public object Object { get; set; } }
    public class Document
    {
        public TextDocument TextDocument = new TextDocument();
        public object Selection { get; set; } = new TextSelection();
        public object Object(string name) => TextDocument;
    }
    public class TextDocument
    {
        public EditPoint StartPoint { get; } = new EditPoint();
        public EditPoint EndPoint => StartPoint;
    }
    public class EditPoint
    {
        public string Text { get; set; }
        public EditPoint CreateEditPoint() => this;
        public string GetText(EditPoint end) => Text;
    }
    public class TextSelection { public bool IsEmpty { get; set; } = true; public string Text { get; set; } }
}
namespace Microsoft.VisualStudio.Shell
{
    public static class ThreadHelper { public static void ThrowIfNotOnUIThread() { } }
    public static class ServiceProvider { public static object GlobalProvider; }
    public static class Package { public static object GetGlobalService(Type type) => null; }
    public static class VsShellUtilities { public static void ShowMessageBox(params object[] args) { } }
}
namespace Microsoft.VisualStudio.Shell.Interop
{
    public class SVsUIShell { }
    public interface IVsUIShell { int GetDialogOwnerHwnd(out IntPtr hwnd); }
    public enum OLEMSGICON { OLEMSGICON_WARNING }
    public enum OLEMSGBUTTON { OLEMSGBUTTON_OK }
    public enum OLEMSGDEFBUTTON { OLEMSGDEFBUTTON_FIRST }
}
namespace System.Windows.Interop
{
    public class WindowInteropHelper { public WindowInteropHelper(object window) { } public IntPtr Owner { get; set; } }
}
namespace Microsoft.SqlServer.Management.UI.VSIntegration
{
    public static class ServiceCache { public static Factory ScriptFactory = new Factory(); }
    public class Factory { public ConnectionInfo CurrentlyActiveWndConnectionInfo = new ConnectionInfo(); }
    public class ConnectionInfo { public UiConnection UIConnectionInfo = new UiConnection(); }
    public class UiConnection { public string ServerName { get; set; } = "server"; public NameValueCollection AdvancedOptions = new NameValueCollection { { "DATABASE", "db" } }; }
}
namespace AxialSqlTools
{
    public static class SettingsManager { public static bool Enabled = true; public static bool GetWarnWhenRunningFatalAction() => Enabled; }
    public static class AxialSqlToolsPackage { public static Logger _logger = new Logger(); }
    public class Logger { public void Error(Exception error, string message) { } }
    public static class GridAccess
    {
        public static object GetNonPublicField(object obj, string field) => obj?.GetType().GetField(field, BindingFlags.Instance | BindingFlags.NonPublic)?.GetValue(obj);
        public static object GetProperty(object obj, string name) => obj?.GetType().GetProperty(name)?.GetValue(obj);
        public static object GetSQLResultsControl() => null;
    }
}
namespace AxialSqlTools.QuerySafety
{
    internal class FatalActionWarningDialog
    {
        internal static Func<bool?> Decision;
        internal static bool Remember;
        internal static int Prompts;
        internal static bool CanRemember;
        internal static FatalActionAnalysis LastAnalysis;
        internal bool AllowRepeats => Remember;
        internal FatalActionWarningDialog(FatalActionAnalysis analysis, string context, bool canRemember)
        { LastAnalysis = analysis; CanRemember = canRemember; }
        internal bool? ShowDialog() { Prompts++; return Decision(); }
    }
}
