using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System.Diagnostics.CodeAnalysis;
using System;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text.RegularExpressions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager
    {
        private void OnAutoConfigDialogClosed(AutoConfigurationService.DialogClosedInfo info)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string reason = info == null
                ? "DialogClosed"
                : $"Result={info.Result}, Server={info.Server}, Database={info.Database}, ChangesApplied={info.ChangesApplied}";

            LogColorSnapshot(reason);

            if (info != null && info.Result == System.Windows.Forms.DialogResult.Cancel)
            {
                // The auto-configure dialog was cancelled. The rule was already saved with
                // GroupName=null/ColorIndex=null to suppress future prompts, so color clears on
                // its own when the config reloads. However, the tab rename must be explicitly
                // reverted here because ReloadAndApplyConfig skips tabs that have no matching group.
                RevertRenamedTabsForServer(info.Server);
                UpdateColorOnly("DialogClosed:Cancel", force: true);
                suppressColorUpdatesUntilUtc = DateTime.UtcNow.AddSeconds(2);
            }
            else if (info != null && !info.ChangesApplied)
            {
                UpdateColorOnly("DialogClosed", force: true);
                // Avoid immediately overwriting regex with a partial snapshot after dialog close.
                suppressColorUpdatesUntilUtc = DateTime.UtcNow.AddSeconds(2);
            }
        }

        private void RevertRenamedTabsForServer(string server)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrWhiteSpace(server))
            {
                return;
            }

            var docs = GetOpenDocumentsSnapshot();
            foreach (var doc in docs)
            {
                if (!string.Equals(doc.Server, server, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                TabRenamer.RevertCookieCaption(doc.Frame, doc.Cookie, doc.Caption);
            }
        }

        private bool IsColorUpdateSuppressed()
        {
            return DateTime.UtcNow < suppressColorUpdatesUntilUtc;
        }

        private void OnColorConfigPathResolved(string configPath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrWhiteSpace(configPath))
            {
                return;
            }

            lastColorConfigPath = configPath;

            string dir = Path.GetDirectoryName(configPath);
            if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
            {
                return;
            }

            if (string.Equals(dir, groupColorWatcherDir, StringComparison.OrdinalIgnoreCase) && groupColorWatcher != null)
            {
                return;
            }

            try
            {
                groupColorWatcher?.Dispose();
                groupColorWatcher = new FileSystemWatcher(dir, "customized-groupid-color-*.json");
                groupColorWatcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.Size | NotifyFilters.FileName;
                groupColorWatcher.Changed += OnGroupColorMapChanged;
                groupColorWatcher.Created += OnGroupColorMapChanged;
                groupColorWatcher.Renamed += OnGroupColorMapChanged;
                groupColorWatcher.EnableRaisingEvents = true;
                groupColorWatcherDir = dir;

                EnvTabsLog.Verbose($"RdtEventManager.Config.cs::OnColorConfigPathResolved - Watching '{dir}' for group color overrides.");
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Group color watcher setup failed: {ex.Message}");
            }
        }

        private void LogColorSnapshot(string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var config = LoadConfigOrNull();
            var rules = cachedRules ?? new System.Collections.Generic.List<TabRuleMatcher.CompiledRule>();
            var docs = GetOpenDocumentsSnapshot();

            EnvTabsLog.Info($"Snapshot ({reason}) - Rules={rules.Count}, Tabs={docs.Count}, AutoColor={(config?.Settings?.EnableAutoColor == true)}");

            foreach (var rule in rules)
            {
                EnvTabsLog.Info($"Rule: Name='{rule.GroupName}', Server='{rule.Server}', Database='{rule.Database}', Priority={rule.Priority}, ColorIndex={rule.ColorIndex}");
            }

            foreach (var doc in docs)
            {
                string fileName = null;
                try { fileName = System.IO.Path.GetFileName(doc.Moniker); }
                catch (Exception ex)
                {
                    EnvTabsLog.Verbose($"LogColorSnapshot - File name parse failed: {ex.Message}");
                }
                string group = TabRuleMatcher.MatchGroup(rules, doc.Server, doc.Database);
                EnvTabsLog.Info($"Tab: Cookie={doc.Cookie}, Server='{doc.Server}', Database='{doc.Database}', Group='{group}', File='{fileName}', Moniker='{doc.Moniker}'");
            }

            string block = colorWriter.BuildGeneratedBlockPreview(docs, rules);
            if (!string.IsNullOrWhiteSpace(block))
            {
                EnvTabsLog.Info("Regex Preview:\n" + block);
            }
        }
        [SuppressMessage("Usage", "VSTHRD010", Justification = "FileSystemWatcher callbacks are off-thread; OnConfigChanged marshals as needed.")]
        private void OnConfigRenamed(object sender, RenamedEventArgs e)
        {
            OnConfigChanged(sender, e);
        }

        [SuppressMessage("Usage", "VSTHRD010", Justification = "FileSystemWatcher callbacks are off-thread; this method marshals as needed.")]
        private void OnGroupColorMapChanged(object sender, FileSystemEventArgs e)
        {
            try
            {
                CancellationToken token;
                lock (groupColorDebounceLock)
                {
                    groupColorDebounceCts?.Cancel();
                    groupColorDebounceCts = new CancellationTokenSource();
                    token = groupColorDebounceCts.Token;
                }

                _ = package.JoinableTaskFactory.RunAsync(async () =>
                {
                    try
                    {
                        EnvTabsLog.Info($"Group color override event: {e.ChangeType} - {e.FullPath}");
                        await Task.Delay(500, token);
                        await package.JoinableTaskFactory.SwitchToMainThreadAsync(token);
                        ApplyGroupColorOverridesToConfig();
                        UpdateColorOnly("GroupColorMapChanged", force: true);
                    }
                    catch (OperationCanceledException)
                    {
                        // Ignored
                    }
                    catch (Exception ex)
                    {
                        EnvTabsLog.Info($"Group color override reload failed: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"OnGroupColorMapChanged error: {ex.Message}");
            }
        }

        private void ApplyGroupColorOverridesToConfig()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrWhiteSpace(lastColorConfigPath))
            {
                return;
            }

            var config = LoadConfigOrNull();
            if (config?.ConnectionGroups == null || config.ConnectionGroups.Count == 0)
            {
                return;
            }

            var rules = cachedRules ?? TabRuleMatcher.CompileRules(config);
            var docs = GetOpenDocumentsSnapshot();

            var overrideByBaseRegex = colorWriter.TryGetOverrideColorsByBaseRegex(lastColorConfigPath);
            if (overrideByBaseRegex == null || overrideByBaseRegex.Count == 0)
            {
                return;
            }

            var groupToBaseRegex = colorWriter.BuildGroupNameToBaseRegex(docs, rules);
            if (groupToBaseRegex == null || groupToBaseRegex.Count == 0)
            {
                return;
            }

            bool updated = false;
            bool enableColorWarning = config.Settings?.EnableColorWarning != false;

            foreach (var rule in config.ConnectionGroups)
            {
                if (rule == null || string.IsNullOrWhiteSpace(rule.GroupName))
                {
                    continue;
                }

                if (groupToBaseRegex.TryGetValue(rule.GroupName, out string baseRegex)
                    && overrideByBaseRegex.TryGetValue(baseRegex, out int newColor)
                    && rule.ColorIndex != newColor)
                {
                    if (enableColorWarning && IsColorUsedByOtherRule(config, rule, newColor))
                    {
                        var choice = MessageBox.Show(
                            "This color is already assigned to another connection group. Open the config to resolve?",
                            "SSMS EnvTabs",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning);

                        if (choice == DialogResult.Yes)
                        {
                            OpenConfigInEditor();
                        }
                    }

                    EnvTabsLog.Info($"Updating colorIndex for group '{rule.GroupName}' from {rule.ColorIndex} to {newColor}");
                    rule.ColorIndex = newColor;
                    updated = true;
                }
            }

            if (updated)
            {
                SaveConfig(config);
                EnvTabsLog.Info("Group color overrides saved. Awaiting config watcher reload.");
            }
        }

        private static bool IsColorUsedByOtherRule(TabGroupConfig config, TabGroupRule currentRule, int colorIndex)
        {
            if (config?.ConnectionGroups == null) return false;

            foreach (var rule in config.ConnectionGroups)
            {
                if (!ReferenceEquals(rule, currentRule) && rule?.ColorIndex == colorIndex)
                {
                    return true;
                }
            }

            return false;
        }

        private static void OpenConfigInEditor()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string path = TabGroupConfigLoader.GetUserConfigPath();
            VsShellUtilities.OpenDocument(ServiceProvider.GlobalProvider, path);
        }

        private static void SaveConfig(TabGroupConfig config)
        {
            TabGroupConfigLoader.SaveConfig(config);
        }

        internal bool EditRuleForActiveConnection()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var config = LoadConfigOrNull();
            if (config?.ConnectionGroups == null || config.ConnectionGroups.Count == 0)
            {
                return false;
            }

            if (!TryGetActiveWindowFrame(out IVsWindowFrame activeFrame))
            {
                return false;
            }

            if (!TryGetMonikerFromFrame(activeFrame, out string moniker) || !IsSqlDocumentMoniker(moniker))
            {
                return false;
            }

            if (!TryGetConnectionInfo(activeFrame, out string server, out string database) || string.IsNullOrWhiteSpace(server))
            {
                return false;
            }

            var matchedRule = FindMatchingConnectionRule(config, server, database);
            if (matchedRule == null)
            {
                EnvTabsLog.Info($"EditRuleForActiveConnection: No matching connection rule for Server='{server}', Database='{database}'.");
                return false;
            }

            bool enableAutoRename = config.Settings?.EnableAutoRename != false;
            bool enableColorWarning = config.Settings?.EnableColorWarning != false;

            if (config.ServerAliases == null)
            {
                config.ServerAliases = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            config.ServerAliases.TryGetValue(server, out string existingAlias);

            var usedColorIndexes = new System.Collections.Generic.HashSet<int>(
                config.ConnectionGroups
                    .Where(r => !ReferenceEquals(r, matchedRule) && r.ColorIndex.HasValue)
                    .Select(r => r.ColorIndex.Value));

            var dialogOptions = new NewRuleDialog.NewRuleDialogOptions
            {
                Server = server,
                Database = database,
                SuggestedName = matchedRule.GroupName,
                SuggestedGroupNameStyle = config.Settings?.SuggestedGroupNameStyle,
                SuggestedColorIndex = matchedRule.ColorIndex,
                ExistingAlias = enableAutoRename ? existingAlias : null,
                HideDatabaseRow = string.IsNullOrWhiteSpace(database),
                HideAliasStep = true,
                HideGroupNameRow = !enableAutoRename,
                IsEditMode = true,
                UsedColorIndexes = usedColorIndexes
            };

            using (var dlg = new NewRuleDialog(dialogOptions))
            {
                var result = dlg.ShowDialog();
                int? resultingColorIndex = matchedRule.ColorIndex;

                if (result == DialogResult.OK || result == DialogResult.Yes)
                {
                    // Use dlg.RuleName directly; null means user cleared the field → save null (no rename).
                    string updatedName = dlg.RuleName;
                    int? updatedColor = dlg.SelectedColorIndex;
                    resultingColorIndex = updatedColor;

                    bool configChanged = false;

                    if (enableAutoRename && !string.Equals(updatedName, matchedRule.GroupName, StringComparison.Ordinal))
                    {
                        matchedRule.GroupName = updatedName;
                        configChanged = true;
                    }

                    if (updatedColor != matchedRule.ColorIndex)
                    {
                        matchedRule.ColorIndex = updatedColor;
                        configChanged = true;
                    }

                    if (configChanged)
                    {
                        SaveConfig(config);
                        ReloadAndApplyConfig();
                    }

                    if (result == DialogResult.Yes)
                    {
                        OpenConfigInEditor();
                    }
                }

                if (enableColorWarning && (result == DialogResult.OK || result == DialogResult.Cancel))
                {
                    if (resultingColorIndex.HasValue && IsColorUsedByOtherRule(config, matchedRule, resultingColorIndex.Value))
                    {
                        var choice = MessageBox.Show(
                            "This color is already assigned to another connection group. Open the config to resolve?",
                            "SSMS EnvTabs",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning);

                        if (choice == DialogResult.Yes)
                        {
                            OpenConfigInEditor();
                        }
                    }
                }
            }

            return true;
        }

        private bool TryGetActiveWindowFrame(out IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            frame = null;

            if (monitorSelection == null)
            {
                return false;
            }

            try
            {
                int hr = monitorSelection.GetCurrentElementValue((uint)VSConstants.VSSELELEMID.SEID_WindowFrame, out object value);
                if (ErrorHandler.Succeeded(hr))
                {
                    frame = value as IVsWindowFrame;
                    return frame != null;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"TryGetActiveWindowFrame failed: {ex.Message}");
            }

            return false;
        }

        private static TabGroupRule FindMatchingConnectionRule(TabGroupConfig config, string server, string database)
        {
            if (config?.ConnectionGroups == null)
            {
                return null;
            }

            // Include rules with a null/empty GroupName (e.g. cancelled auto-config rules) so the
            // Edit command can reopen and restore them rather than saying "no group found".
            return config.ConnectionGroups
                .Where(rule => rule != null)
                .Where(rule => RuleMatchesConnection(rule, server, database))
                .OrderBy(rule => rule.Priority)
                .ThenBy(rule => rule.GroupName, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
        }

        private static bool RuleMatchesConnection(TabGroupRule rule, string server, string database)
        {
            bool hasServer = !string.IsNullOrWhiteSpace(rule.Server);
            bool hasDatabase = !string.IsNullOrWhiteSpace(rule.Database);
            if (!hasServer && !hasDatabase)
            {
                return false;
            }

            if (hasServer && !MatchesRulePattern(rule.Server, server))
            {
                return false;
            }

            if (hasDatabase && !MatchesRulePattern(rule.Database, database))
            {
                return false;
            }

            return true;
        }

        private static bool MatchesRulePattern(string pattern, string value)
        {
            if (string.IsNullOrWhiteSpace(pattern))
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            if (pattern.IndexOf('%') >= 0)
            {
                string regex = "^" + Regex.Escape(pattern).Replace("%", ".*") + "$";
                return Regex.IsMatch(value, regex, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }

            return string.Equals(pattern, value, StringComparison.OrdinalIgnoreCase);
        }

        [SuppressMessage("Usage", "VSTHRD010", Justification = "FileSystemWatcher callbacks are off-thread; this method marshals as needed.")]
        private void OnConfigChanged(object sender, FileSystemEventArgs e)
        {
            try
            {
                // EnvTabsLog.Info($"Config file system event: {e.ChangeType} - {e.FullPath}"); 

                // Debounce logic
                CancellationToken token;
                lock (debounceLock)
                {
                    debounceCts?.Cancel();
                    debounceCts = new CancellationTokenSource();
                    token = debounceCts.Token;
                }

                _ = package.JoinableTaskFactory.RunAsync(async () =>
                {
                    try
                    {
                        // Log inside async task to allow switching to UI thread if needed (though we just write to file off-thread primarily)
                        // EnvTabsLog.Info handles CheckAccess internally now.
                        EnvTabsLog.Info($"Config file system event: {e.ChangeType} - {e.FullPath}");

                        await Task.Delay(500, token);
                        await package.JoinableTaskFactory.SwitchToMainThreadAsync(token);
                        ReloadAndApplyConfig();
                    }
                    catch (OperationCanceledException)
                    {
                        // Ignored
                    }
                    catch (Exception ex)
                    {
                        EnvTabsLog.Info($"Config reload task failed: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"OnConfigChanged error: {ex.Message}");
            }
        }

        private void ReloadAndApplyConfig()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            EnvTabsLog.Info("Configuration changed, reloading...");

            var previousAliasSnapshot = CloneAliasSnapshot(lastAliasSnapshot);

            // Invalidate cache
            cachedConfig = null;
            // Clear suppressed "do not ask again" connections so that if user removed a rule, we effectively reset the "ignore" state
            // and will prompt again if they connect to it.
            AutoConfigurationService.ClearSuppressed();

            var config = LoadConfigOrNull();
            if (TryApplyAliasGroupNameUpdates(previousAliasSnapshot, config))
            {
                cachedRules = TabRuleMatcher.CompileRules(config);
                cachedManualRules = TabRuleMatcher.CompileManualRules(config);
            }

            lastAliasSnapshot = CloneAliasSnapshot(config?.ServerAliases);

            var rules = cachedRules; // LoadConfigOrNull updates cachedRules
            var manualRules = cachedManualRules;

            if (config == null) return;

            var docs = GetOpenDocumentsSnapshot();

            // Rename
            int renamedCount = 0;
            if (config.Settings?.EnableAutoRename != false)
            {
                try
                {
                    var renameCandidates = docs
                        .Where(doc => !string.IsNullOrWhiteSpace(doc?.Moniker)) // Filter empty?
                        .Select(doc => new TabRenameContext
                        {
                            Cookie = doc.Cookie,
                            Frame = doc.Frame,
                            Server = doc.Server,
                            ServerAlias = (config.ServerAliases != null && !string.IsNullOrEmpty(doc.Server) && config.ServerAliases.TryGetValue(doc.Server, out string _sa)) ? _sa : doc.Server,
                            Database = doc.Database,
                            FrameCaption = doc.Caption,
                            Moniker = doc.Moniker
                        })
                        .ToList();

                    if (renameCandidates.Count > 0)
                    {
                        renamedCount = TabRenamer.ApplyRenamesOrThrow(
                            renameCandidates,
                            rules,
                            manualRules,
                            config.Settings?.NewQueryRenameStyle,
                            config.Settings?.SavedFileRenameStyle,
                            config.Settings?.EnableRemoveDotSql != false);
                    }
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"Reload rename failed: {ex.Message}");
                }
            }

            // Color
            if (config.Settings?.EnableAutoColor == true)
            {
                try
                {
                    if (!IsColorUpdateSuppressed())
                    {
                        colorWriter.UpdateFromSnapshot(docs, rules, manualRules);
                    }
                    else
                    {
                        EnvTabsLog.Info("Color update suppressed (ConfigReload)");
                    }
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"Reload color update failed: {ex.Message}");
                }
            }
        }

        private bool TryApplyAliasGroupNameUpdates(TabGroupConfig previousConfig, TabGroupConfig currentConfig)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            return TryApplyAliasGroupNameUpdates(previousConfig?.ServerAliases, currentConfig);
        }

        private bool TryApplyAliasGroupNameUpdates(System.Collections.Generic.Dictionary<string, string> previousAliases, TabGroupConfig currentConfig)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (previousAliases == null || previousAliases.Count == 0 || currentConfig?.ServerAliases == null)
            {
                EnvTabsLog.Info("Alias update check skipped: missing previous or current serverAlias map.");
                return false;
            }

            if (currentConfig.ConnectionGroups == null || currentConfig.ConnectionGroups.Count == 0)
            {
                EnvTabsLog.Info("Alias update check skipped: no connectionGroups found.");
                return false;
            }

            bool updated = false;
            int aliasChanges = 0;
            int aliasMatchedRules = 0;
            int aliasUpdatedRules = 0;

            string suggestedStyle = currentConfig.Settings?.SuggestedGroupNameStyle;

            foreach (var kvp in currentConfig.ServerAliases)
            {
                string server = kvp.Key;
                string newAlias = kvp.Value;

                if (string.IsNullOrWhiteSpace(server) || string.IsNullOrWhiteSpace(newAlias))
                {
                    continue;
                }

                if (!previousAliases.TryGetValue(server, out string oldAlias))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(oldAlias))
                {
                    continue;
                }

                if (string.Equals(oldAlias, newAlias, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                aliasChanges++;

                foreach (var rule in currentConfig.ConnectionGroups)
                {
                    if (rule == null)
                    {
                        continue;
                    }

                    if (!string.Equals(rule.Server, server, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(rule.Database) || string.IsNullOrWhiteSpace(rule.GroupName))
                    {
                        continue;
                    }

                    string database = rule.Database;
                    string expectedGroup = BuildSuggestedGroupNameFromStyle(suggestedStyle, server, oldAlias, database);
                    if (!string.Equals(rule.GroupName, expectedGroup, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    aliasMatchedRules++;

                    string updatedGroup = BuildSuggestedGroupNameFromStyle(suggestedStyle, server, newAlias, database);
                    EnvTabsLog.Info($"Updating group name for server '{server}' from '{rule.GroupName}' to '{updatedGroup}' due to alias change.");
                    rule.GroupName = updatedGroup;
                    aliasUpdatedRules++;
                    updated = true;
                }
            }

            if (aliasChanges == 0)
            {
                EnvTabsLog.Info("Alias update check: no alias changes detected.");
            }
            else
            {
                EnvTabsLog.Info($"Alias update check: aliasChanges={aliasChanges}, matchedRules={aliasMatchedRules}, updatedRules={aliasUpdatedRules}.");
            }

            if (updated)
            {
                SaveConfig(currentConfig);
                EnvTabsLog.Info("Alias-driven group name updates saved. Awaiting config watcher reload.");
            }

            return updated;
        }

        private static string BuildSuggestedGroupNameFromStyle(string style, string serverName, string serverAlias, string databaseValue)
        {
            string serverPart = serverName ?? string.Empty;
            string aliasPart = !string.IsNullOrEmpty(serverAlias) ? serverAlias : serverPart;
            string dbPart = databaseValue ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(style))
            {
                return style
                    .Replace("[server]", serverPart)
                    .Replace("[serverAlias]", aliasPart)
                    .Replace("[db]", dbPart);
            }

            if (string.IsNullOrWhiteSpace(dbPart))
            {
                return serverPart;
            }

            return $"{serverPart} {dbPart}";
        }

        private static System.Collections.Generic.Dictionary<string, string> CloneAliasSnapshot(System.Collections.Generic.Dictionary<string, string> source)
        {
            if (source == null || source.Count == 0)
            {
                return new System.Collections.Generic.Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
            }

            return new System.Collections.Generic.Dictionary<string, string>(source, System.StringComparer.OrdinalIgnoreCase);
        }

        private TabGroupConfig LoadConfigOrNull()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                string path = TabGroupConfigLoader.GetUserConfigPath();
                if (!File.Exists(path))
                {
                    cachedConfig = null;
                    cachedRules = null;
                    cachedManualRules = null;
                    cachedConfigLastWriteUtc = default;
                    return null;
                }

                DateTime lastWriteUtc = File.GetLastWriteTimeUtc(path);
                if (cachedConfig != null && lastWriteUtc == cachedConfigLastWriteUtc)
                {
                    if ((cachedRules == null || cachedManualRules == null) && cachedConfig != null)
                    {
                        cachedRules = TabRuleMatcher.CompileRules(cachedConfig);
                        cachedManualRules = TabRuleMatcher.CompileManualRules(cachedConfig);
                        EnvTabsLog.Info($"RdtEventManager.Config.cs::LoadConfigOrNull - Rebuilt cached rules. Rules={cachedRules?.Count ?? 0}, ManualRules={cachedManualRules?.Count ?? 0}");
                    }
                    return cachedConfig;
                }

                var loaded = TabGroupConfigLoader.LoadOrNull();

                if (loaded != null && loaded.Settings != null)
                {
                    EnvTabsLog.Enabled = loaded.Settings.EnableLogging;
                    EnvTabsLog.VerboseEnabled = loaded.Settings.EnableVerboseLogging;
                }

                cachedConfig = loaded;
                cachedConfigLastWriteUtc = lastWriteUtc;
                cachedRules = TabRuleMatcher.CompileRules(loaded);
                cachedManualRules = TabRuleMatcher.CompileManualRules(loaded);
                if (lastAliasSnapshot == null || lastAliasSnapshot.Count == 0)
                {
                    lastAliasSnapshot = CloneAliasSnapshot(loaded?.ServerAliases);
                }
                EnvTabsLog.Info($"RdtEventManager.Config.cs::LoadConfigOrNull - Loaded. Rules={cachedRules?.Count ?? 0}, ManualRules={cachedManualRules?.Count ?? 0}, AutoColor={loaded?.Settings?.EnableAutoColor}, AutoRename={loaded?.Settings?.EnableAutoRename}, Polling={loaded?.Settings?.EnableConnectionPolling}");
                return loaded;
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"RdtEventManager: config reload failed: {ex.Message}");
                cachedConfig = null;
                cachedRules = null;
                cachedConfigLastWriteUtc = default;
                return null;
            }
        }

        private void UpdateColorOnly(string reason, bool force = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!force && IsColorUpdateSuppressed())
            {
                EnvTabsLog.Info($"Color update suppressed ({reason})");
                return;
            }

            var config = LoadConfigOrNull();
            if (config?.Settings?.EnableAutoColor != true)
            {
                return;
            }

            var rules = cachedRules ?? TabRuleMatcher.CompileRules(config);
            var manualRules = cachedManualRules ?? TabRuleMatcher.CompileManualRules(config);

            if (rules.Count == 0 && manualRules.Count == 0)
            {
                return;
            }

            try
            {
                var docs = GetOpenDocumentsSnapshot();
                colorWriter.UpdateFromSnapshot(docs, rules, manualRules);
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"ColorByRegex update failed ({reason}): {ex.Message}");
            }
        }
    }
}
