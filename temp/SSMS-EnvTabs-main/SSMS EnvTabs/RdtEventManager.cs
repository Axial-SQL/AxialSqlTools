using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager : IVsRunningDocTableEvents3, IVsSelectionEvents, IDisposable
    {
        internal static bool SuppressVerboseLogs;
        private const int RenameRetryCount = 20; // Increased to handle slow connection changes (up to 5s)
        private const int RenameRetryDelayMs = 250;
        private const int ConnectionPollIntervalMs = 1000;

        private readonly AsyncPackage package;
        private readonly IVsRunningDocumentTable rdt;
        private readonly IVsUIShellOpenDocument shellOpenDoc;
        private readonly IVsMonitorSelection monitorSelection;
        private readonly ColorByRegexConfigWriter colorWriter;

        private uint rdtEventsCookie;
        private uint selectionEventsCookie;

        private readonly Dictionary<uint, int> renameRetryCounts = new Dictionary<uint, int>();
        private readonly Dictionary<uint, (string Server, string Database)> lastConnectionByCookie = new Dictionary<uint, (string Server, string Database)>();
        private readonly Dictionary<uint, string> lastCaptionByCookie = new Dictionary<uint, string>();
        private FileSystemWatcher configWatcher;
        private FileSystemWatcher groupColorWatcher;
        private CancellationTokenSource debounceCts;
        private readonly object debounceLock = new object();
        private CancellationTokenSource groupColorDebounceCts;
        private readonly object groupColorDebounceLock = new object();
        private string groupColorWatcherDir;
        private string lastColorConfigPath;
        private Timer connectionPollTimer;

        private TabGroupConfig cachedConfig;
        private DateTime cachedConfigLastWriteUtc;
        private List<TabRuleMatcher.CompiledRule> cachedRules;
        private List<TabRuleMatcher.CompiledManualRule> cachedManualRules;

        private Dictionary<string, string> lastAliasSnapshot = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private DateTime suppressColorUpdatesUntilUtc;


        
        internal sealed class OpenDocumentInfo
        {
            public uint Cookie { get; set; }
            public IVsWindowFrame Frame { get; set; }
            public string Caption { get; set; }
            public string Moniker { get; set; }
            public string Server { get; set; }
            public string Database { get; set; }
        }

        private RdtEventManager(AsyncPackage package, IVsRunningDocumentTable rdt, IVsUIShellOpenDocument shellOpenDoc, IVsMonitorSelection monitorSelection)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            this.rdt = rdt ?? throw new ArgumentNullException(nameof(rdt));
            this.shellOpenDoc = shellOpenDoc;
            this.monitorSelection = monitorSelection;
            this.colorWriter = new ColorByRegexConfigWriter();
        }

        public static async Task<RdtEventManager> CreateAndStartAsync(AsyncPackage package, CancellationToken cancellationToken)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var rdt = await package.GetServiceAsync(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            var shellOpenDoc = await package.GetServiceAsync(typeof(SVsUIShellOpenDocument)) as IVsUIShellOpenDocument;
            var monitorSelection = await package.GetServiceAsync(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;

            if (rdt == null)
            {
                EnvTabsLog.Error("RdtEventManager: SVsRunningDocumentTable missing; EnvTabs disabled.");
                return null;
            }

            var mgr = new RdtEventManager(package, rdt, shellOpenDoc, monitorSelection);
            mgr.Start();
            return mgr;
        }

        private void Start()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            AutoConfigurationService.DialogClosed += OnAutoConfigDialogClosed;
            colorWriter.ConfigPathResolved += OnColorConfigPathResolved;

            try
            {
                rdt.AdviseRunningDocTableEvents(this, out rdtEventsCookie);
                EnvTabsLog.Info($"RdtEventManager: Successfully subscribed to RDT events. Cookie={rdtEventsCookie}");
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"RdtEventManager: failed to subscribe RDT events: {ex.Message}");
            }

            try
            {
                if (monitorSelection != null)
                {
                    monitorSelection.AdviseSelectionEvents(this, out selectionEventsCookie);
                    EnvTabsLog.Info($"RdtEventManager: Successfully subscribed to Selection events. Cookie={selectionEventsCookie}");
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"RdtEventManager: failed to subscribe Selection events: {ex.Message}");
            }

            // Poll for connection changes on active/open documents (SSMS often does not raise RDT events on connection switches)
            connectionPollTimer = new Timer(ConnectionPollTick, null, ConnectionPollIntervalMs, ConnectionPollIntervalMs);
            EnvTabsLog.Verbose("RdtEventManager.cs::Start - Connection polling enabled.");
            
            try
            {
                string configPath = TabGroupConfigLoader.GetUserConfigPath();
                string configDir = Path.GetDirectoryName(configPath);
                if (Directory.Exists(configDir))
                {
                    configWatcher = new FileSystemWatcher(configDir, "TabGroupConfig.json");
                    configWatcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.Size | NotifyFilters.FileName;
                    configWatcher.Changed += OnConfigChanged;
                    configWatcher.Created += OnConfigChanged;
                    configWatcher.Renamed += OnConfigRenamed;
                    configWatcher.EnableRaisingEvents = true;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"RdtEventManager: failed to start config watcher: {ex.Message}");
            }
        }

        public void Dispose()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            AutoConfigurationService.DialogClosed -= OnAutoConfigDialogClosed;
            colorWriter.ConfigPathResolved -= OnColorConfigPathResolved;

            try
            {
                connectionPollTimer?.Dispose();
                connectionPollTimer = null;
                configWatcher?.Dispose();
                groupColorWatcher?.Dispose();
                debounceCts?.Cancel();
                debounceCts?.Dispose();
                groupColorDebounceCts?.Cancel();
                groupColorDebounceCts?.Dispose();
            }
            catch (Exception ex)
            {
                EnvTabsLog.Verbose($"RdtEventManager.Dispose - Cleanup failed: {ex.Message}");
            }

            try
            {
                if (rdtEventsCookie != 0)
                {
                    rdt.UnadviseRunningDocTableEvents(rdtEventsCookie);
                    rdtEventsCookie = 0;
                }
            }
            catch
            {
                // best-effort
            }
        }

        private bool HandlePotentialChange(uint docCookie, IVsWindowFrame frame, string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            var config = LoadConfigOrNull();
            if (config == null)
            {
                return true;
            }

            var rules = cachedRules ?? TabRuleMatcher.CompileRules(config);
            var manualRules = cachedManualRules ?? TabRuleMatcher.CompileManualRules(config);

            bool needsRetry = false;
            int renamedCount = 0;
            bool autoConfigTriggered = false;

            string frameMoniker = null;
            if (frame != null)
            {
                TryGetMonikerFromFrame(frame, out frameMoniker);
                if (!string.IsNullOrWhiteSpace(frameMoniker) && !frameMoniker.EndsWith(".sql", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            if (frame != null && config.Settings?.EnableAutoRename != false)
            {
                string caption = TryReadFrameCaption(frame);
                string moniker = frameMoniker;

                if (IsExecutingCaption(caption))
                {
                    return true;
                }
                
                // If the reason is AttributeChange (connection change), we should attempt rename even if it doesn't look eligible anymore (e.g. if it was already renamed)
                // because the user is changing the connection to something new.
                // Or if it reverted to "SQLQuery1.sql".
                bool forceCheck = reason != null && (
                    reason.StartsWith("AttributeChange", StringComparison.OrdinalIgnoreCase) ||
                    reason.IndexOf("AttributeChange", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    reason.IndexOf("AfterSave", StringComparison.OrdinalIgnoreCase) >= 0
                );

                if (TryGetConnectionInfo(frame, out string server, out string database))
                {
                    var manualMatch = TabRuleMatcher.MatchManual(manualRules, moniker);
                    string matchedGroup = manualMatch?.GroupName ?? TabRuleMatcher.MatchGroup(rules, server, database);
                    bool hasMatchingRule = manualMatch != null || TabRuleMatcher.MatchRule(rules, server, database) != null;

                    bool captionMissingGroup = !string.IsNullOrWhiteSpace(matchedGroup)
                        && (string.IsNullOrWhiteSpace(caption) || caption.IndexOf(matchedGroup, StringComparison.OrdinalIgnoreCase) < 0);

                    bool shouldRename = forceCheck || IsRenameEligible(moniker, caption) || captionMissingGroup;

                    if (shouldRename)
                    {
                        try
                        {
                            var ctx = new TabRenameContext
                            {
                                Cookie = docCookie,
                                Frame = frame,
                                Server = server,
                                ServerAlias = (config.ServerAliases != null && !string.IsNullOrEmpty(server) && config.ServerAliases.TryGetValue(server, out string _sa)) ? _sa : server,
                                Database = database,
                                FrameCaption = caption,
                                Moniker = moniker
                            };
                            string style = config.Settings?.NewQueryRenameStyle;
                            string savedStyle = config.Settings?.SavedFileRenameStyle;
                            bool removeDotSql = config.Settings?.EnableRemoveDotSql != false;
                            renamedCount = TabRenamer.ApplyRenamesOrThrow(new[] { ctx }, rules, manualRules, style, savedStyle, removeDotSql);

                            if (renamedCount > 0)
                            {
                                EnvTabsLog.Info($"Renamed tab. Reason={reason}, Cookie={docCookie}, Server='{server}', DB='{database}', Count={renamedCount}");
                            }
                        }
                        catch (Exception ex)
                        {
                            EnvTabsLog.Info($"Rename failed ({reason}) cookie={docCookie}: {ex.Message}");
                        }
                    }

                    // Auto-Configure logic if no rule matched
                    if (!hasMatchingRule && !string.IsNullOrWhiteSpace(config.Settings?.AutoConfigure))
                    {
                        if (!string.IsNullOrWhiteSpace(server))
                        {
                            EnvTabsLog.Info($"AutoConfigure: No matching rule. Reason={reason}, Cookie={docCookie}, Server='{server}', DB='{database}'");
                            autoConfigTriggered = true;
                            // Dispatch to UI thread later to avoid blocking RDT event
                            _ = package.JoinableTaskFactory.RunAsync(async () =>
                            {
                                await package.JoinableTaskFactory.SwitchToMainThreadAsync();
                                AutoConfigurationService.ProposeNewRule(config, server, database);
                            });
                        }
                    }
                }
                else
                {
                    // Needs retry if connection info not found yet
                    needsRetry = true;
                }
            }
            else
            {
                if (frame == null)
                {
                    return true;
                }
                else if (config.Settings?.EnableAutoRename == false)
                {
                    // Auto-rename disabled: still attempt connection info and auto-config.
                    TryGetMonikerFromFrame(frame, out string moniker);
                    if (TryGetConnectionInfo(frame, out string server, out string database))
                    {
                        string matchedGroup = TabRuleMatcher.MatchGroup(rules, server, database);
                        var manualMatch = TabRuleMatcher.MatchManual(manualRules, moniker);
                        bool hasMatchingRule = manualMatch != null || TabRuleMatcher.MatchRule(rules, server, database) != null;

                        if (!hasMatchingRule && !string.IsNullOrWhiteSpace(config.Settings?.AutoConfigure) && !string.IsNullOrWhiteSpace(server))
                        {
                            EnvTabsLog.Info($"AutoConfigure: No matching rule (auto-rename disabled). Reason={reason}, Cookie={docCookie}, Server='{server}', DB='{database}'");
                            autoConfigTriggered = true;
                            _ = package.JoinableTaskFactory.RunAsync(async () =>
                            {
                                await package.JoinableTaskFactory.SwitchToMainThreadAsync();
                                AutoConfigurationService.ProposeNewRule(config, server, database);
                            });
                        }
                    }
                }
            }

            bool isConnectionEvent = reason != null && (
                reason.IndexOf("DocumentWindowShow", StringComparison.OrdinalIgnoreCase) >= 0 ||
                reason.IndexOf("FirstDocumentLock", StringComparison.OrdinalIgnoreCase) >= 0 ||
                reason.IndexOf("AttributeChange", StringComparison.OrdinalIgnoreCase) >= 0 ||
                reason.IndexOf("AttributeChangeEx", StringComparison.OrdinalIgnoreCase) >= 0 ||
                reason.IndexOf("ActiveFrameChanged", StringComparison.OrdinalIgnoreCase) >= 0 ||
                reason.IndexOf("ConnectionPoll", StringComparison.OrdinalIgnoreCase) >= 0 ||
                reason.IndexOf("CaptionPoll", StringComparison.OrdinalIgnoreCase) >= 0
            );

            bool shouldUpdateColor = !needsRetry && (renamedCount > 0 || isConnectionEvent || autoConfigTriggered);
                        
            if (config.Settings?.EnableAutoColor == true && shouldUpdateColor)
            {
                try
                {
                    if (!IsColorUpdateSuppressed())
                    {
                        var docs = GetOpenDocumentsSnapshot();
                        colorWriter.UpdateFromSnapshot(docs, rules, manualRules);
                    }
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"ColorByRegex update failed ({reason}): {ex.Message}");
                }
            }

            return !needsRetry;
        }







        private void ScheduleRenameRetry(uint docCookie, string reason)
        {
            if (docCookie == 0) return;

            if (renameRetryCounts.ContainsKey(docCookie))
            {
                return;
            }

            renameRetryCounts[docCookie] = 0;

            _ = package.JoinableTaskFactory.RunAsync(async () =>
            {
                for (int i = 0; i < RenameRetryCount; i++)
                {
                    await Task.Delay(RenameRetryDelayMs).ConfigureAwait(true);
                    await package.JoinableTaskFactory.SwitchToMainThreadAsync();

                    if (!renameRetryCounts.ContainsKey(docCookie))
                    {
                        return;
                    }

                    renameRetryCounts[docCookie] = i + 1;
                    EnvTabsLog.Verbose($"RenameRetry: Reason={reason}, Cookie={docCookie}, Attempt={i + 1}");

                    string moniker = TryGetMonikerFromCookie(docCookie);
                    if (!IsSqlDocumentMoniker(moniker))
                    {
                        renameRetryCounts.Remove(docCookie);
                        return;
                    }

                    IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                    if (frame == null)
                    {
                        continue;
                    }

                    string attemptReason = $"{reason}:Retry#{i + 1}";
                    bool done = HandlePotentialChange(docCookie, frame, attemptReason);
                    if (done)
                    {
                        renameRetryCounts.Remove(docCookie);
                        return;
                    }
                }

                renameRetryCounts.Remove(docCookie);
            });
        }

        private static bool IsExecutingCaption(string caption)
        {
            if (string.IsNullOrWhiteSpace(caption)) return false;
            return caption.StartsWith("Executing...", StringComparison.OrdinalIgnoreCase);
        }

        private void ConnectionPollTick(object state)
        {
            _ = package.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    await package.JoinableTaskFactory.SwitchToMainThreadAsync();
                    PollConnectionChanges();
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"RdtEventManager.cs::ConnectionPollTick - Error: {ex.Message}");
                }
            });
        }

        private void PollConnectionChanges()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var config = LoadConfigOrNull();
            if (config == null)
            {
                return;
            }

            if (config.Settings?.EnableConnectionPolling != true)
            {
                return;
            }

            SuppressVerboseLogs = true;
            try
            {

            var docs = GetOpenDocumentsSnapshot();
            var seen = new HashSet<uint>();

            foreach (var doc in docs)
            {
                if (doc == null || doc.Frame == null)
                {
                    continue;
                }

                seen.Add(doc.Cookie);

                string newServer = doc.Server ?? string.Empty;
                string newDatabase = doc.Database ?? string.Empty;

                if (lastConnectionByCookie.TryGetValue(doc.Cookie, out var previous))
                {
                    bool serverChanged = !string.Equals(previous.Server ?? string.Empty, newServer, StringComparison.OrdinalIgnoreCase);
                    bool databaseChanged = !string.Equals(previous.Database ?? string.Empty, newDatabase, StringComparison.OrdinalIgnoreCase);

                    if (serverChanged || databaseChanged)
                    {
                        EnvTabsLog.Verbose($"RdtEventManager.cs::PollConnectionChanges - Detected connection change. Cookie={doc.Cookie}, '{previous.Server}.{previous.Database}' -> '{newServer}.{newDatabase}'");
                        lastConnectionByCookie[doc.Cookie] = (newServer, newDatabase);
                        HandlePotentialChange(doc.Cookie, doc.Frame, "ConnectionPoll");
                    }
                }
                else
                {
                    // First observation for this cookie (no log; only log on change)
                    lastConnectionByCookie[doc.Cookie] = (newServer, newDatabase);

                    if (config.Settings?.EnableAutoRename == false && config.Settings?.EnableAutoColor == true)
                    {
                        HandlePotentialChange(doc.Cookie, doc.Frame, "ConnectionPoll:Initial");
                    }
                }

                // Caption change detection (execution resets to default)
                string caption = doc.Caption ?? string.Empty;
                bool hadPreviousCaption = lastCaptionByCookie.TryGetValue(doc.Cookie, out string previousCaption);
                bool captionChanged = !hadPreviousCaption || !string.Equals(previousCaption ?? string.Empty, caption, StringComparison.OrdinalIgnoreCase);
                if (captionChanged)
                {
                    lastCaptionByCookie[doc.Cookie] = caption;

                    if (!IsExecutingCaption(caption))
                    {
                        EnvTabsLog.Verbose($"CaptionPoll: Detected caption change. Cookie={doc.Cookie}, Caption='{caption}'");
                        bool done = HandlePotentialChange(doc.Cookie, doc.Frame, "CaptionPoll");
                        if (!done)
                        {
                            ScheduleRenameRetry(doc.Cookie, "CaptionPoll");
                        }
                    }
                }
            }

            // Clean up stale entries
            var staleCookies = lastConnectionByCookie.Keys.Where(c => !seen.Contains(c)).ToList();
            foreach (var cookie in staleCookies)
            {
                lastConnectionByCookie.Remove(cookie);
                lastCaptionByCookie.Remove(cookie);
            }
            }
            finally
            {
                SuppressVerboseLogs = false;
            }
        }

        // --- IVsSelectionEvents ---
        public int OnSelectionChanged(IVsHierarchy pHierOld, uint itemidOld, IVsMultiItemSelect pMISOld, ISelectionContainer pSCOld, IVsHierarchy pHierNew, uint itemidNew, IVsMultiItemSelect pMISNew, ISelectionContainer pSCNew)
        {
            return VSConstants.S_OK;
        }

        public int OnElementValueChanged(uint elementid, object varValueOld, object varValueNew)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // SEID_WindowFrame = 2
            if (elementid == (uint)VSConstants.VSSELELEMID.SEID_WindowFrame)
            {
                // When active window changes, we check if we need to update info.
                // This might catch cases where connection changed via a dialog and we return to the window.
                var frame = varValueNew as IVsWindowFrame;
                if (frame != null)
                {
                    _ = package.JoinableTaskFactory.RunAsync(async () =>
                    {
                        // Slight delay to allow SSMS to update its internal state
                        await Task.Delay(500);
                        await package.JoinableTaskFactory.SwitchToMainThreadAsync();
                        CheckActiveFrame(frame);
                    });
                }
            }
            return VSConstants.S_OK;
        }

        public int OnCmdUIContextChanged(uint dwCmdUICookie, int fActive)
        {
            return VSConstants.S_OK;
        }

        private void CheckActiveFrame(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try 
            {
                var docs = GetOpenDocumentsSnapshot();
                var match = docs.FirstOrDefault(d => d.Frame == frame);
                if (match != null)
                {
                    // Force check even if we think nothing changed
                    HandlePotentialChange(match.Cookie, frame, "ElementValueChanged");
                }
            }
            catch(Exception ex)
            {
                EnvTabsLog.Info($"CheckActiveFrame Error: {ex.Message}");
            }
        }

    }
}
