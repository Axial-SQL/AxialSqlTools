using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using static AxialSqlTools.AxialSqlToolsPackage;

namespace AxialSqlTools
{
    public static class GridAccess
    {
        public static void SetPropertyValue(object targetObj, string fieldName, object fieldValue)
        {
            targetObj.GetType().GetProperty(fieldName).SetValue(targetObj, fieldValue);
        }

        public static object GetProperty(object obj, string field)
        {
            if (obj is null)
            {
                return null;
            }

            return obj.GetType().GetProperty(field, BindingFlags.Public | BindingFlags.Instance)?.GetValue(obj);
        }

        public static object GetNonPublicField(object obj, string field)
        {
            if (obj is null)
            {
                return null;
            }

            return obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(obj);
        }

        public static FieldInfo GetNonPublicFieldInfo(object obj, string field)
            => obj?.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance);

        public static object GetSQLResultsControl()
        {
            try
            {
                var factoryType = ServiceCache.ScriptFactory.GetType();

                var method = factoryType.GetMethod(
                    "GetCurrentlyActiveFrameDocView",
                    BindingFlags.NonPublic | BindingFlags.Instance);

                if (method == null)
                {
                    _logger?.Error("GetCurrentlyActiveFrameDocView method not found");
                    return null;
                }

                var docView = method.Invoke(
                    ServiceCache.ScriptFactory,
                    new object[] { ServiceCache.VSMonitorSelection, false, null });

                // It might be ObjectExplorerTool, a designer, etc.
                if (!(docView is SqlScriptEditorControl scriptEditor))
                {
                    // No active query window - handle as you like
                    // e.g. return null, or throw a more descriptive exception
                    return null;
                }

                return GetNonPublicField(scriptEditor, "m_sqlResultsControl");
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in GetSQLResultsControl");
                return null;
            }
        }

        public static QEStatusBarManager GetStatusBarManager()
        {
            var objType = ServiceCache.ScriptFactory.GetType();
            var method1 = objType.GetMethod("GetCurrentlyActiveFrameDocView", BindingFlags.NonPublic | BindingFlags.Instance);
            var Result = (SqlScriptEditorControl)method1.Invoke(ServiceCache.ScriptFactory, new object[] { ServiceCache.VSMonitorSelection, false, null });

            return Result.StatusBarManager;
        }

        public static CollectionBase GetGridContainers()
        {
            try
            {
                var SQLResultsControl = GetSQLResultsControl();
                if (SQLResultsControl is null)
                {
                    return null;
                }

                var m_gridResultsPage = GetNonPublicField(SQLResultsControl, "m_gridResultsPage");
                if (m_gridResultsPage is null)
                {
                    return null;
                }

                var gridContainersObj = GetNonPublicField(m_gridResultsPage, "m_gridContainers");
                if (gridContainersObj is null)
                {
                    return null;
                }

                // SSMS 21: Return CollectionBase directly
                if (gridContainersObj is CollectionBase gridContainers)
                {
                    return gridContainers;
                }

                // SSMS 22: Wrap List<ResultSetAndGridContainer> as CollectionBase
                if (gridContainersObj is ICollection collection)
                {
                    var wrapper = new GridContainerCollection();
                    foreach (var item in collection)
                    {
                        wrapper.AddItem(item);
                    }
                    return wrapper;
                }

                _logger?.Error($"Unexpected grid containers type: {gridContainersObj.GetType().FullName}");
                return null;
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in GetGridContainers");
                return null;
            }
        }

        // Wrapper for SSMS 22's List<T> to match SSMS 21's CollectionBase
        private class GridContainerCollection : CollectionBase
        {
            public void AddItem(object item)
            {
                List.Add(item);
            }
        }

        public static void ChangeStatusBarContent(int OpenTranCount, bool isColumnEncryptionSettingOn, string ActualElapsedTime)
        {
            QEStatusBarManager statusBarManager = GetStatusBarManager();

            if (OpenTranCount > 0)
            {
                var msg = "One transaction is still open!";
                if (OpenTranCount > 1)
                {
                    msg = $"{OpenTranCount} transactions are still open!";
                }

                var currentMsg = statusBarManager.StatusText;
                statusBarManager.SetKnownState(QEStatusBarKnownStates.Executing);
                statusBarManager.StatusText = currentMsg + " | " + msg;
            }

            if (isColumnEncryptionSettingOn)
            {
                var oeMsg = "Column Encryption Setting is ON";
                statusBarManager.StatusText = statusBarManager.StatusText + " | " + oeMsg;
            }

            var statusBarManager_executionTimePanel = GetNonPublicField(statusBarManager, "executionTimePanel");
            SetPropertyValue(statusBarManager_executionTimePanel, "Text", ActualElapsedTime);

            var statusBarManager_completedTimePanel = GetNonPublicField(statusBarManager, "completedTimePanel");
            //statusBarManager_completedTimePanel.Visible = true;

            var generalPanel = GetNonPublicField(statusBarManager, "generalPanel");
            if (OpenTranCount > 0)
            {
                SetPropertyValue(generalPanel, "ForeColor", Color.Red);
            }
            else
            {
                SetPropertyValue(generalPanel, "ForeColor", Color.Black);
            }

            // TODO - need to contract font from existing property..
            Font defaultFont = new Font("Segoe UI", 9);
            Font boldFont = new Font("Segoe UI", 10, FontStyle.Bold);

            var statusStrip = GetNonPublicField(statusBarManager, "statusStrip");
            if (OpenTranCount > 0 || isColumnEncryptionSettingOn)
            {
                SetPropertyValue(statusStrip, "Font", boldFont);
            }
            else
            {
                SetPropertyValue(statusStrip, "Font", defaultFont);
            }
        }

        public static Color? FindMatchingConnectionColor(string serverName, string databaseName)
        {
            var rules = SettingsManager.GetConnectionColorRules();

            foreach (var rule in rules)
            {
                if (!rule.IsEnabled)
                    continue;

                bool hasServerPattern = !string.IsNullOrEmpty(rule.ServerNamePattern);
                bool hasDatabasePattern = !string.IsNullOrEmpty(rule.DatabaseNamePattern);

                if (!hasServerPattern && !hasDatabasePattern)
                    continue;

                bool serverMatch = !hasServerPattern ||
                    (!string.IsNullOrEmpty(serverName) && serverName.IndexOf(rule.ServerNamePattern, StringComparison.OrdinalIgnoreCase) >= 0);

                bool databaseMatch = !hasDatabasePattern ||
                    (!string.IsNullOrEmpty(databaseName) && databaseName.IndexOf(rule.DatabaseNamePattern, StringComparison.OrdinalIgnoreCase) >= 0);

                if (serverMatch && databaseMatch)
                {
                    try
                    {
                        return ColorTranslator.FromHtml(rule.StatusBarColor);
                    }
                    catch { }
                }
            }

            return null;
        }

        private sealed class StatusStripColorController
        {
            private readonly StatusStrip _statusStrip;
            private readonly Dictionary<ToolStripItem, ToolStripItemColorSnapshot> _itemSnapshots =
                new Dictionary<ToolStripItem, ToolStripItemColorSnapshot>();
            private bool _isApplied;

            public StatusStripColorController(StatusStrip statusStrip)
            {
                _statusStrip = statusStrip;
                OriginalBackColor = statusStrip.BackColor;
                OriginalForeColor = statusStrip.ForeColor;
            }

            public Color OriginalBackColor { get; }
            public Color OriginalForeColor { get; }

            public void ApplyColor(Color color)
            {
                Color foreColor = ContrastColor(color);

                foreach (ToolStripItem item in _statusStrip.Items)
                {
                    if (!_itemSnapshots.ContainsKey(item))
                    {
                        _itemSnapshots[item] = new ToolStripItemColorSnapshot
                        {
                            BackColor = item.BackColor,
                            ForeColor = item.ForeColor
                        };
                    }

                    item.BackColor = color;
                    item.ForeColor = foreColor;
                }

                _statusStrip.BackColor = color;
                _statusStrip.ForeColor = foreColor;
                _isApplied = true;
            }

            public void Restore()
            {
                if (!_isApplied)
                {
                    return;
                }

                _statusStrip.BackColor = OriginalBackColor;
                _statusStrip.ForeColor = OriginalForeColor;

                foreach (var itemSnapshot in _itemSnapshots)
                {
                    if (itemSnapshot.Key == null || itemSnapshot.Key.IsDisposed)
                    {
                        continue;
                    }

                    itemSnapshot.Key.BackColor = itemSnapshot.Value.BackColor;
                    itemSnapshot.Key.ForeColor = itemSnapshot.Value.ForeColor;
                }

                _itemSnapshots.Clear();
                _isApplied = false;
            }
        }

        private sealed class ToolStripItemColorSnapshot
        {
            public Color BackColor { get; set; }
            public Color ForeColor { get; set; }
        }

        private sealed class OpenDocumentTabInfo
        {
            public IVsWindowFrame Frame { get; set; }
            public string Caption { get; set; }
        }

        private static readonly Dictionary<StatusStrip, StatusStripColorController> _statusStripColorControllers =
            new Dictionary<StatusStrip, StatusStripColorController>();

        private static StatusStripColorController GetStatusStripColorController(StatusStrip statusStrip)
        {
            if (!_statusStripColorControllers.TryGetValue(statusStrip, out StatusStripColorController controller))
            {
                controller = new StatusStripColorController(statusStrip);
                _statusStripColorControllers[statusStrip] = controller;
            }

            return controller;
        }

        private static void TryRestoreStatusBarColor(object statusBarManager, StatusStrip statusStrip)
        {
            if (statusStrip == null)
            {
                return;
            }

            if (!_statusStripColorControllers.TryGetValue(statusStrip, out StatusStripColorController controller))
            {
                return;
            }

            controller.Restore();
            _statusStripColorControllers.Remove(statusStrip);
            TrySetNativeServerBackground(statusBarManager, controller.OriginalBackColor);
        }

        private static void TrySetNativeServerBackground(object statusBarManager, Color color)
        {
            if (statusBarManager == null)
            {
                return;
            }

            try
            {
                MethodInfo method = statusBarManager.GetType().GetMethod(
                    "SetServerBackground",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                    null,
                    new[] { typeof(Color) },
                    null);

                method?.Invoke(statusBarManager, new object[] { color });
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error setting native status bar background");
            }
        }

        public static void ApplyConnectionColor(string serverName, string databaseName)
        {
            // Connection coloring applies only to query windows.
            try
            {
                var SQLResultsControl = GetSQLResultsControl();
                if (SQLResultsControl == null)
                    return;

                Color? matchedColor = FindMatchingConnectionColor(serverName, databaseName);

                var statusBarManager = GetStatusBarManager();
                if (statusBarManager == null)
                    return;

                var statusStrip = GetNonPublicField(statusBarManager, "statusStrip") as StatusStrip;

                if (matchedColor.HasValue)
                {
                    TrySetNativeServerBackground(statusBarManager, matchedColor.Value);

                    if (statusStrip != null)
                    {
                        GetStatusStripColorController(statusStrip).ApplyColor(matchedColor.Value);
                    }
                }
                else
                {
                    TryRestoreStatusBarColor(statusBarManager, statusStrip);
                }
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in ApplyConnectionColor");
            }
        }

        private static System.Windows.Threading.DispatcherTimer _reapplyTimer;
        private static int _reapplyRetryCount;

        public static void ColorAllDocumentTabs()
        {
            try
            {
                var wpfApp = System.Windows.Application.Current;
                if (wpfApp == null) return;

                var mainWindow = wpfApp.MainWindow;
                if (mainWindow == null) return;

                ColorAllTabs(mainWindow);
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in ColorAllDocumentTabs");
            }
        }

        private static void ColorAllTabs(System.Windows.DependencyObject mainWindow)
        {
            var tabItems = new List<System.Windows.DependencyObject>();
            FindElementsByTypeName(mainWindow, "DocumentTabItem", tabItems);
            var openDocuments = GetOpenDocumentTabInfos();

            if (tabItems.Count > 0)
            {
                foreach (var tabItem in tabItems)
                {
                    var headerCtrl = tabItem as System.Windows.Controls.HeaderedContentControl;
                    string header = headerCtrl?.Header?.ToString();
                    Color? color = FindMatchingConnectionColorForTab(tabItem, header, openDocuments);
                    var curvedBorder = FindChildByTypeName(tabItem, "SimpleCurvedBorder")
                                   ?? FindChildByTypeName(tabItem, "TopCurvedBorder");
                    ApplyColorToElement(curvedBorder ?? tabItem, color);
                    if (curvedBorder != null)
                    {
                        ApplyForegroundToElement(tabItem, color);
                    }
                }
            }
            else
            {
                ColorSingleTabFallback(mainWindow, openDocuments);
            }
        }

        private static void ColorSingleTabFallback(System.Windows.DependencyObject root, List<OpenDocumentTabInfo> openDocuments)
        {
            try
            {
                var docGroups = new List<System.Windows.DependencyObject>();
                FindElementsByTypeName(root, "DocumentGroupControl", docGroups);
                if (docGroups.Count == 0)
                    return;

                foreach (var docGroup in docGroups)
                {
                    var header = FindChildByTypeName(docGroup, "DragUndockHeader");
                    if (header == null)
                        continue;

                    var fe = header as System.Windows.FrameworkElement;
                    if (fe == null || fe.ActualHeight <= 0 || fe.Visibility != System.Windows.Visibility.Visible)
                        continue;

                    string headerTitle = fe.DataContext?.ToString();
                    Color? color = FindMatchingConnectionColorForTab(header, headerTitle, openDocuments);
                    ApplyColorToElement(header, color);
                    return;
                }
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in ColorSingleTabFallback");
            }
        }

        private static void ApplyColorToElement(System.Windows.DependencyObject element, Color? matchedColor)
        {
            if (element is System.Windows.Controls.Control ctrl)
            {
                if (matchedColor.HasValue)
                {
                    var wpfColor = System.Windows.Media.Color.FromRgb(
                        matchedColor.Value.R, matchedColor.Value.G, matchedColor.Value.B);
                    var brush = new System.Windows.Media.SolidColorBrush(wpfColor);
                    brush.Freeze();
                    var foregroundBrush = new System.Windows.Media.SolidColorBrush(ToWpfColor(ContrastColor(matchedColor.Value)));
                    foregroundBrush.Freeze();
                    ctrl.Background = brush;
                    ctrl.Foreground = foregroundBrush;
                }
                else
                {
                    ctrl.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
                    ctrl.ClearValue(System.Windows.Controls.Control.ForegroundProperty);
                }
            }
            else if (element is System.Windows.Controls.Border border)
            {
                if (matchedColor.HasValue)
                {
                    var wpfColor = System.Windows.Media.Color.FromRgb(
                        matchedColor.Value.R, matchedColor.Value.G, matchedColor.Value.B);
                    var brush = new System.Windows.Media.SolidColorBrush(wpfColor);
                    brush.Freeze();
                    border.Background = brush;
                }
                else
                {
                    border.ClearValue(System.Windows.Controls.Border.BackgroundProperty);
                }
            }
        }

        private static void ApplyForegroundToElement(System.Windows.DependencyObject element, Color? matchedColor)
        {
            if (!(element is System.Windows.Controls.Control ctrl))
                return;

            if (matchedColor.HasValue)
            {
                var brush = new System.Windows.Media.SolidColorBrush(ToWpfColor(ContrastColor(matchedColor.Value)));
                brush.Freeze();
                ctrl.Foreground = brush;
            }
            else
            {
                ctrl.ClearValue(System.Windows.Controls.Control.ForegroundProperty);
            }
        }

        private static Color? FindMatchingConnectionColorForTab(
            System.Windows.DependencyObject tabElement,
            string header,
            List<OpenDocumentTabInfo> openDocuments)
        {
            if (TryGetConnectionInfoFromTab(tabElement, header, openDocuments, out string server, out string database))
            {
                return FindMatchingConnectionColor(server, database);
            }

            return FindMatchingConnectionColorFromCaption(header);
        }

        private static Color? FindMatchingConnectionColorFromCaption(string caption)
        {
            if (string.IsNullOrEmpty(caption))
                return null;

            var parts = caption.Split(new[] { " - " }, 2, StringSplitOptions.None);
            if (parts.Length < 2)
                return null;

            string connectionPart = parts[1].Trim();
            int parenIdx = connectionPart.IndexOf(" (");
            if (parenIdx > 0)
                connectionPart = connectionPart.Substring(0, parenIdx);

            return FindMatchingConnectionColor(connectionPart, connectionPart);
        }

        private static bool TryGetConnectionInfoFromTab(
            System.Windows.DependencyObject tabElement,
            string header,
            List<OpenDocumentTabInfo> openDocuments,
            out string server,
            out string database)
        {
            server = null;
            database = null;

            if (TryFindWindowFrame(tabElement, out IVsWindowFrame frame)
                && TryGetConnectionInfo(frame, out server, out database))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(header) && openDocuments != null)
            {
                OpenDocumentTabInfo match = null;
                int matchCount = 0;

                foreach (var doc in openDocuments)
                {
                    if (string.Equals(doc.Caption, header, StringComparison.OrdinalIgnoreCase))
                    {
                        match = doc;
                        matchCount++;
                    }
                }

                if (matchCount == 1 && TryGetConnectionInfo(match.Frame, out server, out database))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetConnectionInfo(IVsWindowFrame frame, out string server, out string database)
        {
            server = null;
            database = null;

            try
            {
                if (frame == null
                    || frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out object docView) != VSConstants.S_OK
                    || docView == null)
                {
                    return false;
                }

                object connection = GetNonPublicField(docView, "m_connection");
                if (connection == null)
                    return false;

                server = GetProperty(connection, "DataSource") as string;
                database = GetProperty(connection, "Database") as string;
                return !string.IsNullOrWhiteSpace(server);
            }
            catch
            {
                return false;
            }
        }

        private static bool TryFindWindowFrame(object root, out IVsWindowFrame frame)
        {
            frame = null;

            if (root == null)
                return false;

            var visited = new HashSet<object>(new ReferenceEqualityComparer());
            var queue = new Queue<Tuple<object, int>>();
            queue.Enqueue(Tuple.Create(root, 0));
            visited.Add(root);

            while (queue.Count > 0 && visited.Count < 300)
            {
                var current = queue.Dequeue();
                object obj = current.Item1;
                int depth = current.Item2;

                if (obj is IVsWindowFrame foundFrame)
                {
                    frame = foundFrame;
                    return true;
                }

                if (depth >= 3 || IsSimpleProbeType(obj.GetType()))
                    continue;

                foreach (object child in EnumerateProbeValues(obj))
                {
                    if (child != null && visited.Add(child))
                    {
                        queue.Enqueue(Tuple.Create(child, depth + 1));
                    }
                }
            }

            return false;
        }

        private static IEnumerable<object> EnumerateProbeValues(object obj)
        {
            const BindingFlags Flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            foreach (FieldInfo field in obj.GetType().GetFields(Flags))
            {
                object value = null;
                try { value = field.GetValue(obj); } catch { }
                if (value != null)
                    yield return value;
            }

            foreach (PropertyInfo property in obj.GetType().GetProperties(Flags))
            {
                if (!property.CanRead || property.GetIndexParameters().Length != 0)
                    continue;

                object value = null;
                try { value = property.GetValue(obj, null); } catch { }
                if (value != null)
                    yield return value;
            }
        }

        private static bool IsSimpleProbeType(Type type)
        {
            return type == null
                || type.IsPrimitive
                || type.IsEnum
                || type == typeof(string)
                || type == typeof(decimal)
                || type == typeof(DateTime)
                || type == typeof(TimeSpan)
                || type == typeof(Guid)
                || type == typeof(IntPtr)
                || type == typeof(UIntPtr)
                || type == typeof(Color)
                || typeof(System.Windows.Media.Brush).IsAssignableFrom(type);
        }

        private static List<OpenDocumentTabInfo> GetOpenDocumentTabInfos()
        {
            var documents = new List<OpenDocumentTabInfo>();

            try
            {
                var rdt = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
                var shellOpenDoc = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SVsUIShellOpenDocument)) as IVsUIShellOpenDocument;
                if (rdt == null || shellOpenDoc == null)
                    return documents;

                rdt.GetRunningDocumentsEnum(out IEnumRunningDocuments enumDocs);
                if (enumDocs == null)
                    return documents;

                uint[] cookies = new uint[1];
                while (enumDocs.Next(1, cookies, out uint fetched) == VSConstants.S_OK && fetched == 1)
                {
                    if (rdt.GetDocumentInfo(
                        cookies[0],
                        out uint _,
                        out uint _,
                        out uint _,
                        out string moniker,
                        out IVsHierarchy _,
                        out uint _,
                        out IntPtr _) != VSConstants.S_OK)
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(moniker) || !moniker.EndsWith(".sql", StringComparison.OrdinalIgnoreCase))
                        continue;

                    Guid logicalView = Guid.Empty;
                    uint[] itemid = new uint[1];
                    if (shellOpenDoc.IsDocumentOpen(null, 0, moniker, ref logicalView, 0, out IVsUIHierarchy _, itemid, out IVsWindowFrame frame, out int isOpen) != VSConstants.S_OK
                        || isOpen == 0
                        || frame == null)
                    {
                        continue;
                    }

                    string caption = null;
                    if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_Caption, out object captionObject) == VSConstants.S_OK)
                    {
                        caption = captionObject as string;
                    }

                    documents.Add(new OpenDocumentTabInfo { Frame = frame, Caption = caption });
                }
            }
            catch
            {
            }

            return documents;
        }

        private static Color ContrastColor(Color color)
        {
            double luma = ((0.299 * color.R) + (0.587 * color.G) + (0.114 * color.B)) / 255;
            return luma > 0.5 ? Color.Black : Color.White;
        }

        private static System.Windows.Media.Color ToWpfColor(Color color)
        {
            return System.Windows.Media.Color.FromRgb(color.R, color.G, color.B);
        }

        private static System.Windows.DependencyObject FindChildByTypeName(System.Windows.DependencyObject parent, string typeName)
        {
            if (parent == null) return null;

            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child.GetType().Name == typeName)
                    return child;
            }
            for (int i = 0; i < count; i++)
            {
                var result = FindChildByTypeName(System.Windows.Media.VisualTreeHelper.GetChild(parent, i), typeName);
                if (result != null) return result;
            }
            return null;
        }

        private static void FindElementsByTypeName(System.Windows.DependencyObject obj, string typeName, List<System.Windows.DependencyObject> results)
        {
            if (obj == null) return;

            if (obj.GetType().Name == typeName)
            {
                results.Add(obj);
            }

            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(obj);
            for (int i = 0; i < count; i++)
            {
                FindElementsByTypeName(System.Windows.Media.VisualTreeHelper.GetChild(obj, i), typeName, results);
            }
        }

        public static void ScheduleReapplyAllTabColors()
        {
            if (_reapplyTimer == null)
            {
                _reapplyTimer = new System.Windows.Threading.DispatcherTimer();
                _reapplyTimer.Interval = TimeSpan.FromMilliseconds(150);
                _reapplyTimer.Tick += (s, e) =>
                {
                    _reapplyRetryCount++;
                    try { ColorAllDocumentTabs(); } catch { }

                    if (_reapplyRetryCount >= 4)
                    {
                        _reapplyTimer.Stop();
                    }
                };
            }

            _reapplyRetryCount = 0;
            _reapplyTimer.Stop();
            _reapplyTimer.Start();
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
                return sqlDataTypeName + "(" + (sqlDataColumnSize == 2147483647 ? "MAX" : sqlDataColumnSize.ToString()) + ")";
            }
            else if (sqlDataTypeName == "DECIMAL" || sqlDataTypeName == "NUMERIC")
            {
                return sqlDataTypeName + "(" + NumericPrecision + "," + NumericScale + ")";
            }
            else if (sqlDataTypeName == "DATETIME2")
            {
                return sqlDataTypeName + "(" + NumericScale + ")";
            }


            //else if (sqlDataTypeName == "DECIMAL")
            //    sqlDataTypeName = sqlDataTypeName;
            //TODO ... list all other types that need additional handling 

            return sqlDataTypeName;
        }

        public static IGridControl GetFocusGridControl()
        {
            object outVsWindowFrame = null;
            ServiceCache.VSMonitorSelection.GetCurrentElementValue((int)VSConstants.VSSELELEMID.SEID_WindowFrame, out outVsWindowFrame);

            var vsWindowFrame = outVsWindowFrame as IVsWindowFrame;
            vsWindowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out var outControl);

            var control = (Control)outControl;
            return (GridControl)((ContainerControl)((ContainerControl)control).ActiveControl).ActiveControl;
        }

        public static List<DataTable> GetDataTables()
        {
            List<DataTable> dataTables = new List<DataTable>();

            try
            {
                var SQLResultsControl = GetSQLResultsControl();
                if (SQLResultsControl is null)
                {
                    return dataTables;
                }

                var m_gridResultsPage = GetNonPublicField(SQLResultsControl, "m_gridResultsPage");
                if (m_gridResultsPage is null)
                {
                    return dataTables;
                }

                var gridContainersObj = GetNonPublicField(m_gridResultsPage, "m_gridContainers");
                if (gridContainersObj is null)
                {
                    return dataTables;
                }

                // Handle both SSMS 21 (CollectionBase) and SSMS 22 (List<T>)
                if (!(gridContainersObj is IEnumerable gridContainers))
                {
                    _logger?.Error($"Unexpected grid containers type: {gridContainersObj.GetType().FullName}");
                    return dataTables;
                }

                foreach (var gridContainer in gridContainers)
                {
                    if (!(GetNonPublicField(gridContainer, "m_grid") is GridControl grid))
                    {
                        continue;
                    }

                    var gridStorage = grid.GridStorage;
                    if (gridStorage is null)
                    {
                        continue;
                    }

                    if (!(GetNonPublicField(gridStorage, "m_schemaTable") is DataTable schemaTable))
                    {
                        continue;
                    }

                    List<int> columnSizes = new List<int>();
                    if (GetNonPublicField(grid, "m_Columns") is GridColumnCollection gridColumns)
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
                                || columnType == typeof(byte[]))
                        {
                            columnType = typeof(string);
                        }

                        string sqlDataTypeName = GetColumnSqlType(schemaTable.Rows[c]);

                        DataColumn newColumn = data.Columns.Add(columnNameInt, columnType);

                        var columnName = schemaTable.Rows[c][0].ToString();
                        if (string.IsNullOrEmpty(columnName))
                        {
                            columnName = columnNameInt;
                        }

                        newColumn.ExtendedProperties.Add("columnName", columnName);
                        newColumn.ExtendedProperties.Add("sqlType", sqlDataTypeName);
                        if (columnSizes.Count > c + 1)
                        {
                            newColumn.ExtendedProperties.Add("columnWidthInPixels", columnSizes[c + 1]);
                        }
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
                            {
                                cellData = cellData == "0" ? "False" : "True";
                            }

                            // leave some types as strings because the conversion from string fails
                            if (columnType == typeof(Guid)
                                || columnType == typeof(TimeSpan)
                                || columnType == typeof(DateTime)
                                || columnType == typeof(DateTimeOffset)
                                || columnType == typeof(byte[]))
                            {
                                columnType = typeof(string);
                            }

                            var typedValue = Convert.ChangeType(cellData, columnType, CultureInfo.InvariantCulture);

                            rowItems.Add(typedValue);
                        }

                        data.Rows.Add(rowItems.ToArray());
                    }

                    data.AcceptChanges();

                    dataTables.Add(data);
                }
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Error in GetDataTables");
            }

            return dataTables;
        }

        public static string TryGetStatisticsMessagesText()
        {
            var sqlResultsControl = GetSQLResultsControl();
            return TryFindStatisticsTextFromResultsControl(sqlResultsControl, out var statisticsText) ? statisticsText : null;
        }

        private static bool TryFindStatisticsTextFromResultsControl(object sqlResultsControl, out string statisticsText)
        {
            statisticsText = null;

            var resultsTabControl = GetNonPublicField(sqlResultsControl, "m_resultsTabCtrl") as TabControl;
            if (resultsTabControl == null)
            {
                return false;
            }

            var selectedIndex = resultsTabControl.SelectedIndex;
            var messagesIndex = FindMessagesTabIndex(resultsTabControl);

            var visited = new HashSet<object>(new ReferenceEqualityComparer());
            if (TryFindStatisticsText(resultsTabControl, 0, visited, out statisticsText))
            {
                return true;
            }

            if (messagesIndex < 0 || messagesIndex >= resultsTabControl.TabPages.Count)
            {
                return false;
            }

            try
            {
                if (resultsTabControl.SelectedIndex != messagesIndex)
                {
                    resultsTabControl.SelectedIndex = messagesIndex;
                    resultsTabControl.Update();
                    resultsTabControl.Refresh();
                    Application.DoEvents();
                    TryFlushStatisticsMessages();
                }

                visited.Clear();
                if (TryFindStatisticsText(resultsTabControl.TabPages[messagesIndex], 0, visited, out statisticsText))
                {
                    return true;
                }

                visited.Clear();
                return TryFindStatisticsText(resultsTabControl, 0, visited, out statisticsText);
            }
            finally
            {
                if (selectedIndex >= 0
                    && selectedIndex < resultsTabControl.TabPages.Count
                    && resultsTabControl.SelectedIndex != selectedIndex)
                {
                    try
                    {
                        resultsTabControl.SelectedIndex = selectedIndex;
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static int FindMessagesTabIndex(TabControl tabControl)
        {
            if (tabControl == null)
            {
                return -1;
            }

            for (var index = 0; index < tabControl.TabPages.Count; index++)
            {
                var tabPageText = tabControl.TabPages[index]?.Text;
                if (!string.IsNullOrWhiteSpace(tabPageText)
                    && tabPageText.IndexOf("messages", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return index;
                }
            }

            return tabControl.TabPages.Count > 1 ? 1 : -1;
        }

        public static void TryFlushStatisticsMessages()
        {
            var sqlResultsControl = GetSQLResultsControl();
            if (sqlResultsControl is null)
            {
                return;
            }

            TryInvokeParameterlessDelegateField(sqlResultsControl, "m_FlushTextWritersInvoker");
            TryInvokeParameterlessMethod(sqlResultsControl, "FlushTextWriters");

            TryFlushWriterLike(GetNonPublicField(sqlResultsControl, "m_messagesWriter"));
            TryFlushWriterLike(GetNonPublicField(sqlResultsControl, "m_resultsWriter"));
            TryFlushWriterLike(GetNonPublicField(sqlResultsControl, "m_errorsWriter"));
        }

        private static bool TryFindStatisticsText(object candidate, int depth, HashSet<object> visited, out string statisticsText)
        {
            statisticsText = null;

            if (candidate is null || depth > 4)
            {
                return false;
            }

            if (candidate is string candidateText)
            {
                if (LooksLikeStatisticsOutput(candidateText))
                {
                    statisticsText = candidateText;
                    return true;
                }

                return false;
            }

            if (candidate is TextWriter textWriter)
            {
                var writerText = SafeToString(textWriter);
                if (LooksLikeStatisticsOutput(writerText))
                {
                    statisticsText = writerText;
                    return true;
                }
            }

            if (candidate is IVsTextView textView
                && TryGetTextFromVsTextView(textView, out var textViewText)
                && LooksLikeStatisticsOutput(textViewText))
            {
                statisticsText = textViewText;
                return true;
            }

            if (candidate is IVsTextLines textLines
                && TryGetTextFromVsTextLines(textLines, out var textLinesText)
                && LooksLikeStatisticsOutput(textLinesText))
            {
                statisticsText = textLinesText;
                return true;
            }

            var candidateType = candidate.GetType();
            if (candidateType.IsPrimitive
                || candidate is decimal
                || candidate is DateTime
                || candidate is TimeSpan
                || candidate is Enum)
            {
                return false;
            }

            if (!candidateType.IsValueType && !visited.Add(candidate))
            {
                return false;
            }

            if (candidate is Control control)
            {
                if (LooksLikeStatisticsOutput(control.Text))
                {
                    statisticsText = control.Text;
                    return true;
                }

                foreach (Control childControl in control.Controls)
                {
                    if (TryFindStatisticsText(childControl, depth + 1, visited, out statisticsText))
                    {
                        return true;
                    }
                }
            }

            foreach (var member in GetRelevantMembers(candidate))
            {
                if (TryFindStatisticsText(member.Value, depth + 1, visited, out statisticsText))
                {
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<KeyValuePair<string, object>> GetRelevantMembers(object obj)
        {
            foreach (var field in EnumerateInstanceFields(obj.GetType()))
            {
                if (!IsRelevantMember(field.Name, field.FieldType))
                {
                    continue;
                }

                object value = null;
                try
                {
                    value = field.GetValue(obj);
                }
                catch
                {
                }

                if (value != null)
                {
                    yield return new KeyValuePair<string, object>(field.Name, value);
                }
            }

            var seenProperties = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in EnumerateInstanceProperties(obj.GetType()))
            {
                if (!property.CanRead || property.GetIndexParameters().Length > 0 || !IsRelevantMember(property.Name, property.PropertyType))
                {
                    continue;
                }

                if (!seenProperties.Add(property.Name))
                {
                    continue;
                }

                object value = null;
                try
                {
                    value = property.GetValue(obj);
                }
                catch
                {
                }

                if (value != null)
                {
                    yield return new KeyValuePair<string, object>(property.Name, value);
                }
            }
        }

        private static IEnumerable<FieldInfo> EnumerateInstanceFields(Type type)
        {
            const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly;

            for (var currentType = type; currentType != null; currentType = currentType.BaseType)
            {
                foreach (var field in currentType.GetFields(bindingFlags))
                {
                    yield return field;
                }
            }
        }

        private static IEnumerable<PropertyInfo> EnumerateInstanceProperties(Type type)
        {
            const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly;

            for (var currentType = type; currentType != null; currentType = currentType.BaseType)
            {
                foreach (var property in currentType.GetProperties(bindingFlags))
                {
                    yield return property;
                }
            }
        }

        private static bool IsRelevantMember(string memberName, Type memberType)
        {
            return LooksLikeRelevantMemberName(memberName) || IsRelevantMemberType(memberType);
        }

        private static bool IsRelevantMemberType(Type memberType)
        {
            if (memberType is null)
            {
                return false;
            }

            return typeof(Control).IsAssignableFrom(memberType)
                || typeof(TextWriter).IsAssignableFrom(memberType)
                || typeof(IVsTextView).IsAssignableFrom(memberType)
                || typeof(IVsTextLines).IsAssignableFrom(memberType);
        }

        private static bool LooksLikeRelevantMemberName(string memberName)
        {
            if (string.IsNullOrEmpty(memberName))
            {
                return false;
            }

            memberName = memberName.ToLowerInvariant();
            return memberName.Contains("message")
                || memberName.Contains("text")
                || memberName.Contains("writer")
                || memberName.Contains("buffer")
                || memberName.Contains("document")
                || memberName.Contains("editor")
                || memberName.Contains("view")
                || memberName.Contains("output")
                || memberName.Contains("result")
                || memberName.Contains("pane")
                || memberName.Contains("page")
                || memberName.Contains("sql")
                || memberName.Contains("exec");
        }

        private static string SafeToString(object value)
        {
            try
            {
                return value?.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static void TryFlushWriterLike(object writer)
        {
            if (writer is null)
            {
                return;
            }

            if (writer is TextWriter textWriter)
            {
                try
                {
                    textWriter.Flush();
                }
                catch
                {
                }
            }

            TryInvokeParameterlessMethod(writer, "Flush");
        }

        private static bool TryInvokeParameterlessDelegateField(object target, string fieldName)
        {
            if (target is null || string.IsNullOrWhiteSpace(fieldName))
            {
                return false;
            }

            try
            {
                if (GetNonPublicField(target, fieldName) is Delegate callback)
                {
                    callback.DynamicInvoke();
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool TryInvokeParameterlessMethod(object target, string methodName)
        {
            if (target is null || string.IsNullOrWhiteSpace(methodName))
            {
                return false;
            }

            try
            {
                var method = target.GetType().GetMethod(methodName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, Type.EmptyTypes, null);
                if (method == null)
                {
                    return false;
                }

                method.Invoke(target, null);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetTextFromVsTextView(IVsTextView textView, out string text)
        {
            text = null;

            try
            {
                if (textView == null || textView.GetBuffer(out IVsTextLines textLines) != VSConstants.S_OK)
                {
                    return false;
                }

                return TryGetTextFromVsTextLines(textLines, out text);
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetTextFromVsTextLines(IVsTextLines textLines, out string text)
        {
            text = null;

            try
            {
                if (textLines == null || textLines.GetLastLineIndex(out var lastLine, out var lastIndex) != VSConstants.S_OK)
                {
                    return false;
                }

                if (textLines.GetLineText(0, 0, lastLine, lastIndex, out var fullText) != VSConstants.S_OK)
                {
                    return false;
                }

                text = fullText;
                return !string.IsNullOrWhiteSpace(text);
            }
            catch
            {
                return false;
            }
        }

        public static bool LooksLikeStatisticsOutput(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return text.IndexOf("SQL Server Execution Times:", StringComparison.OrdinalIgnoreCase) >= 0
                || (text.IndexOf("logical reads", StringComparison.OrdinalIgnoreCase) >= 0
                    && text.IndexOf("Scan count", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
        {
            public new bool Equals(object x, object y) => ReferenceEquals(x, y);

            public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
        }
    }
}
