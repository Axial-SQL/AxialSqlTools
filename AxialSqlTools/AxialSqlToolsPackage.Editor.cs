using Aurora;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Input;

namespace AxialSqlTools
{
    // Editor window events, connection coloring, snippets, and query templates.
    public sealed partial class AxialSqlToolsPackage
    {
        private System.Windows.Threading.DispatcherTimer _connectionColorRetryTimer;

        private System.Windows.Threading.DispatcherTimer _activeConnectionMonitorTimer;

        private int _connectionColorRetryCount;

        private string _lastObservedConnectionColorKey;

        private string _lastObservedConnectionWindowKey;

        private Plugin m_plugin = null;

        private CommandRegistry m_commandRegistry = null;

        private CommandBar m_commandBarQueryTemplates = null;

        public Dictionary<string, string> globalSnippets = new Dictionary<string, string>();

        private readonly List<KeypressCommandFilter> _commandFilters = new List<KeypressCommandFilter>();

        private readonly HashSet<IVsTextView> _registeredTextViews = new HashSet<IVsTextView>();

        private bool TryApplyConnectionColorForWindow(EnvDTE.Window window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string windowKind = null;
            try { windowKind = window?.Kind; } catch { }

            if (window == null || windowKind != "Document")
            {
                return false;
            }

            var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfo();
            if (connectionInfo == null || string.IsNullOrWhiteSpace(connectionInfo.ServerName))
            {
                return false;
            }

            GridAccess.ApplyConnectionColor(connectionInfo.ServerName, connectionInfo.Database);
            GridAccess.ColorAllDocumentTabs();
            return true;
        }

        private void StartActiveWindowConnectionMonitor()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_activeConnectionMonitorTimer != null)
            {
                return;
            }

            _activeConnectionMonitorTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(300)
            };

            _activeConnectionMonitorTimer.Tick += (sender, args) =>
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                try
                {
                    RefreshActiveWindowConnectionColorIfChanged();
                }
                catch (Exception ex)
                {
                    _logger?.Error(ex, "Failed to monitor active window connection changes.");
                }
            };

            _activeConnectionMonitorTimer.Start();
        }

        private void RefreshActiveWindowConnectionColorIfChanged(bool force = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var activeWindow = ServiceCache.ExtensibilityModel?.ActiveWindow;
            string windowKind = null;
            try { windowKind = activeWindow?.Kind; } catch { }

            if (activeWindow == null || windowKind != "Document")
            {
                _lastObservedConnectionColorKey = null;
                _lastObservedConnectionWindowKey = null;
                return;
            }

            var connectionInfo = ScriptFactoryAccess.GetCurrentConnectionInfo();
            if (connectionInfo == null || string.IsNullOrWhiteSpace(connectionInfo.ServerName))
            {
                _lastObservedConnectionColorKey = null;
                _lastObservedConnectionWindowKey = null;
                return;
            }

            string connectionKey = $"{connectionInfo.ServerName}|{connectionInfo.Database}";
            string windowKey = GetWindowTrackingKey(activeWindow);

            if (!force &&
                string.Equals(connectionKey, _lastObservedConnectionColorKey, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(windowKey, _lastObservedConnectionWindowKey, StringComparison.Ordinal))
            {
                return;
            }

            GridAccess.ApplyConnectionColor(connectionInfo.ServerName, connectionInfo.Database);
            GridAccess.ColorAllDocumentTabs();

            _lastObservedConnectionColorKey = connectionKey;
            _lastObservedConnectionWindowKey = windowKey;
        }

        private static string GetWindowTrackingKey(EnvDTE.Window window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (window == null)
            {
                return string.Empty;
            }

            try
            {
                if (window.Document != null)
                {
                    return window.Document.FullName ?? window.Caption ?? string.Empty;
                }
            }
            catch
            {
            }

            try
            {
                return window.Caption ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void ScheduleActiveWindowConnectionColorRefresh()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_connectionColorRetryTimer == null)
            {
                _connectionColorRetryTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(200)
                };

                _connectionColorRetryTimer.Tick += (sender, args) =>
                {
                    ThreadHelper.ThrowIfNotOnUIThread();

                    _connectionColorRetryCount++;

                    bool applied = false;
                    try
                    {
                        applied = TryApplyConnectionColorForWindow(ServiceCache.ExtensibilityModel?.ActiveWindow);
                    }
                    catch (Exception ex)
                    {
                        _logger?.Error(ex, "Failed to refresh connection color.");
                    }

                    if (applied || _connectionColorRetryCount >= 6)
                    {
                        _connectionColorRetryTimer.Stop();
                    }
                };
            }

            RefreshActiveWindowConnectionColorIfChanged(force: true);
            _connectionColorRetryCount = 0;
            _connectionColorRetryTimer.Stop();
            _connectionColorRetryTimer.Start();
        }

        private void WindowActivated_Event(EnvDTE.Window GotFocus, EnvDTE.Window LostFocus)
        {

            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                EnsureStatisticsExecutionHookForActiveWindow("window-activated");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to reattach statistics handler during window activation.");
            }

            try
            {
                // Attach editor gestures independently of the snippet setting.
                var DocData = GridAccess.GetProperty(GotFocus?.Object, "DocData");
                if (DocData != null)
                {
                    var txtMgr = (IVsTextManager)GridAccess.GetProperty(DocData, "TextManager");

                    IVsTextView textView;
                    if (txtMgr != null && txtMgr.GetActiveView(0, null, out textView) == VSConstants.S_OK)
                    {
                        AttachStringSelection(textView);
                        // Prevent duplicate filters on the same text view
                        if (SettingsManager.GetUseSnippets() && !_registeredTextViews.Contains(textView))
                        {
                            _registeredTextViews.Add(textView);
                            var CommandFilter = new KeypressCommandFilter(this, textView);
                            CommandFilter.AddToChain();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred attaching editor gestures.");
            }

            // Apply connection-based coloring (document tab + status bar)
            try
            {
                if (GotFocus != null)
                {
                    TryApplyConnectionColorForWindow(GotFocus);
                    GridAccess.ScheduleReapplyAllTabColors();
                    ScheduleActiveWindowConnectionColorRefresh();
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred applying connection color");
            }

        }

        private static void AttachStringSelection(IVsTextView nativeView)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var components = Package.GetGlobalService(typeof(SComponentModel)) as IComponentModel;
            var view = components?.GetService<IVsEditorAdaptersFactoryService>()?.GetWpfTextView(nativeView);
            if (view == null || view.IsClosed || view.Properties.ContainsProperty(typeof(SqlStringContent))) return;

            view.VisualElement.PreviewMouseLeftButtonDown += OnMouseDown;
            view.Closed += OnClosed;
            view.Properties.AddProperty(typeof(SqlStringContent), true);

            void OnMouseDown(object sender, MouseButtonEventArgs e)
            {
                if (e.ClickCount != 2 || Keyboard.Modifiers != ModifierKeys.None || view.IsClosed
                    || !SettingsManager.GetGeneralSettings().selectQuotedStringOnDoubleClick) return;
                try
                {
                    var mouse = e.GetPosition(view.VisualElement);
                    var line = view.TextViewLines.GetTextViewLineContainingYCoordinate(mouse.Y + view.ViewportTop);
                    var point = line?.GetBufferPositionFromXCoordinate(mouse.X + view.ViewportLeft);
                    if (!point.HasValue) return;
                    var snapshot = point.Value.Snapshot;
                    if (!SqlStringContent.TryFind(snapshot.GetText(), point.Value.Position, out int start, out int length)) return;
                    var span = new SnapshotSpan(snapshot, start, length);
                    view.Selection.Mode = TextSelectionMode.Stream;
                    view.Selection.Select(span, false);
                    view.Caret.MoveTo(span.End);
                    e.Handled = true;
                }
                catch (Exception ex)
                {
                    _logger?.Warn("Unable to select SQL string content ({0}).", ex.GetType().Name);
                }
            }

            void OnClosed(object sender, EventArgs e)
            {
                view.VisualElement.PreviewMouseLeftButtonDown -= OnMouseDown;
                view.Closed -= OnClosed;
                view.Properties.RemoveProperty(typeof(SqlStringContent));
            }
        }

        private void WindowClosing_Event(EnvDTE.Window Window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Re-color remaining tabs after a tab closes
            try
            {
                QuerySafety.FatalActionGuard.ForgetDocument(Window?.Document);
                GridAccess.ScheduleReapplyAllTabColors();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred in WindowClosing_Event");
            }
        }

        private void WindowCreated_Event(EnvDTE.Window Window)
        {

            ThreadHelper.ThrowIfNotOnUIThread();

            // subscribe to the execution completed event
            try
            {
                var sqlResultsControl = GridAccess.GetNonPublicField(Window.Object, "m_sqlResultsControl");
                AttachStatisticsExecutionCompletedHandler(sqlResultsControl);
                TryApplyConnectionColorForWindow(Window);
                GridAccess.ScheduleReapplyAllTabColors();
                ScheduleActiveWindowConnectionColorRefresh();

            }
            catch (Exception ex) 
            {
                _logger.Error(ex, "An exception occurred");
            }
        }

        public void LoadGlobalSnippets()
        {

            if (SettingsManager.GetUseSnippets())
            {

                var snippetFolder = SettingsManager.GetSnippetFolder();

                if (Directory.Exists(snippetFolder))
                {
                    var allFiles = Directory.EnumerateFiles(snippetFolder, "*.sql");

                    foreach (var file in allFiles)
                    {

                        FileInfo fi = new FileInfo(file);

                        if (fi.Length < 1024 * 1024)
                        {
                            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fi.Name);

                            globalSnippets.Add(fileNameWithoutExtension.ToUpper(), System.IO.File.ReadAllText(fi.FullName));
                        }

                    }


                }

            }



        }

        public void RefreshTemplatesList()
        {
            if (m_commandBarQueryTemplates is null)
            {
                return;
            }

            //Delete existing autogenerated controls
            for (int idx = m_commandBarQueryTemplates.Controls.Count; idx >= 3; idx--)
            {
                CommandBarControl control = m_commandBarQueryTemplates.Controls[idx];
                if (control == null) continue;

                control.Delete();
            }

            Dictionary<string, string> fileNamesCache = new Dictionary<string, string>();

            string Folder = SettingsManager.GetTemplatesFolder();
            if (!Directory.Exists(Folder))
            {
                _logger.Warn("The configured templates folder is unavailable: {0}", Folder);
                return;
            }
            int i = 2;
            CreateCommands(ref i, ref fileNamesCache, Folder, m_commandRegistry, m_commandBarQueryTemplates);

            UpdateRenamedTemplatesControls(m_commandBarQueryTemplates, fileNamesCache);
            
        }

        private void UpdateRenamedTemplatesControls(CommandBar commandBarFolder, Dictionary<string, string> fileNamesCache)
        {
            foreach (CommandBarControl control in commandBarFolder.Controls)
            {
                string keyToFind = control.Caption;
                if (fileNamesCache.TryGetValue(keyToFind, out string value))
                {
                    control.Caption = value;
                    control.Tag = keyToFind;
                }
                else if (fileNamesCache.TryGetValue(control.Tag, out string value2)) //This is the case when the file was renamed
                {
                    control.Caption = value2;
                }
            }
        }

        private void CreateCommands(ref int i, ref Dictionary<string, string> fileNamesCache, string Folder,
                CommandRegistry m_commandRegistry, CommandBar commandBarFolder)
        {

            var dirs = Directory.GetDirectories(Folder);

            foreach (var dirStr in dirs)
            {
                var di = new FileInfo(dirStr);

                string controlName = "Folder_" + i;

                fileNamesCache.Add(controlName, di.Name);
                i = i + 1;

                CommandBar commandBarFolderNext = m_plugin.AddCommandBarMenu(controlName, MsoBarPosition.msoBarMenuBar, commandBarFolder);

                CreateCommands(ref i, ref fileNamesCache, Path.Combine(Folder, dirStr), m_commandRegistry, commandBarFolderNext);

            }

            var files = Directory.GetFiles(Folder);

            foreach (var file in files)
            {
                var fi = new FileInfo(file);

                string controlName = "Template_" + i;
                fileNamesCache.Add(controlName, fi.Name);

                var nc = new CommandProcessor(m_plugin, controlName, controlName, "");
                nc.package = this;
                nc.FullFileName = fi.FullName;
                m_commandRegistry.RegisterCommand(true, nc, true, commandBarFolder);

                i = i + 1;

            }

            UpdateRenamedTemplatesControls(commandBarFolder, fileNamesCache);

        }
    }
}
