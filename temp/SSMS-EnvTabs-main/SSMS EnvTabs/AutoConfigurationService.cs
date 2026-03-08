using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Windows.Forms;
using Microsoft.VisualStudio.Shell;

namespace SSMS_EnvTabs
{
    internal static class AutoConfigurationService
    {
        private static readonly HashSet<string> suppressedConnections = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        internal static event Action<DialogClosedInfo> DialogClosed;

        internal sealed class DialogClosedInfo
        {
            public DialogResult Result { get; set; }
            public string Server { get; set; }
            public string Database { get; set; }
            public bool ChangesApplied { get; set; }
        }

        private struct AddRuleContext
        {
            public TabGroupConfig Config { get; set; }
            public string Server { get; set; }
            public string Database { get; set; }
            public bool UseDb { get; set; }
            public string GroupName { get; set; }
            public int? ColorIndex { get; set; }
        }

        public static void ClearSuppressed()
        {
            suppressedConnections.Clear();
        }

        public static void ProposeNewRule(TabGroupConfig config, string server, string database)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrWhiteSpace(server)) return;
            if (config == null || config.Settings == null) return;

            string mode = config.Settings.AutoConfigure;
            if (string.IsNullOrWhiteSpace(mode)) return;

            // Normalize mode
            bool useDb = mode.IndexOf("db", System.StringComparison.OrdinalIgnoreCase) >= 0;

            string connectionKey = useDb ? $"{server}::{database}" : server;
            
            if (suppressedConnections.Contains(connectionKey)) 
            {
                EnvTabsLog.Info($"AutoConfig: Suppressed key {connectionKey}");
                return;
            }

            bool enableAutoRename = config.Settings.EnableAutoRename != false;
            bool enableAliasPrompt = config.Settings.EnableServerAliasPrompt != false;
            bool enableColorWarning = config.Settings.EnableColorWarning != false;

            // Prepare suggested values
            int nextColor = FindNextColorValues(config);
            string suggestedName;
            
            // Check for Server Alias (only used when auto-rename is enabled)
            if (config.ServerAliases == null)
            {
                config.ServerAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
            
            string existingAlias = null;
            if (enableAutoRename && config.ServerAliases.TryGetValue(server, out string alias))
            {
                existingAlias = alias;
            }

            // If alias prompt is enabled, treat "alias == server" as no alias so the prompt can show.
            if (enableAliasPrompt && !string.IsNullOrWhiteSpace(existingAlias)
                && string.Equals(existingAlias, server, StringComparison.OrdinalIgnoreCase))
            {
                existingAlias = null;
            }

            string style = config.Settings.SuggestedGroupNameStyle;
            if (string.IsNullOrWhiteSpace(style))
            {
                // Fallback to legacy logic
                if (useDb && !string.IsNullOrWhiteSpace(database))
                {
                    string prefix = !string.IsNullOrWhiteSpace(existingAlias) ? existingAlias : server;
                    suggestedName = $"{prefix} {database}";
                }
                else
                {
                    suggestedName = !string.IsNullOrWhiteSpace(existingAlias) ? existingAlias : server;
                }
            }
            else
            {
                // Use configured style
                string rawServerPart = server ?? "";
                string aliasPart = !string.IsNullOrWhiteSpace(existingAlias) ? existingAlias : rawServerPart;
                string dbPart = (useDb && !string.IsNullOrWhiteSpace(database)) ? database : "";
                
                suggestedName = style
                    .Replace("[server]", rawServerPart)
                    .Replace("[serverAlias]", aliasPart)
                    .Replace("[db]", dbPart);
            }
            
            EnvTabsLog.Info($"AutoConfig: Proposing new rule for {connectionKey}. Color={nextColor}, Name={suggestedName}, Prompt={config.Settings.EnableConfigurePrompt}");

            // Suppress immediately so we don't re-enter if RDT events fire while the dialog is open.
            suppressedConnections.Add(connectionKey);

            if (!config.Settings.EnableConfigurePrompt)
            {
                // Silent mode: save the rule immediately with the suggested defaults.
                var silentContext = new AddRuleContext
                {
                    Config = config,
                    Server = server,
                    Database = database,
                    UseDb = useDb,
                    GroupName = suggestedName,
                    ColorIndex = nextColor
                };
                AddRuleAndSave(silentContext);
                return;
            }

            // Prompt mode: show the dialog BEFORE saving any rule so that tabs are
            // not renamed or colored until the user explicitly clicks Save.
            try
            {
                EnvTabsLog.Info("AutoConfig: Opening prompt dialog...");

                var usedColorIndexes = new HashSet<int>(
                    config.ConnectionGroups
                        .Where(r => r != null && r.ColorIndex.HasValue)
                        .Select(r => r.ColorIndex.Value));

                var dialogOptions = new NewRuleDialog.NewRuleDialogOptions
                {
                    Server = server,
                    Database = !useDb ? null : database,
                    SuggestedName = suggestedName,
                    SuggestedGroupNameStyle = config.Settings?.SuggestedGroupNameStyle,
                    SuggestedColorIndex = nextColor,
                    ExistingAlias = enableAutoRename ? existingAlias : null,
                    HideDatabaseRow = !useDb,
                    HideAliasStep = !enableAutoRename || !enableAliasPrompt,
                    HideGroupNameRow = !enableAutoRename,
                    UsedColorIndexes = usedColorIndexes
                };

                using (var dlg = new NewRuleDialog(dialogOptions))
                {
                    // When the user clicks Next on the alias step, persist the alias right away
                    // so it is available even if the user later cancels the rule step.
                    dlg.AliasConfirmed += (serverAlias) =>
                    {
                        if (enableAutoRename && !string.IsNullOrWhiteSpace(serverAlias))
                        {
                            config.ServerAliases[server] = serverAlias;
                            SaveConfig(config);
                            EnvTabsLog.Info($"AutoConfig: Server alias saved on Next: '{server}' -> '{serverAlias}'");
                        }
                    };

                    var result = dlg.ShowDialog();
                    EnvTabsLog.Info($"AutoConfig: Dialog result = {result}");
                    bool changesApplied = false;
                    int? resultingColorIndex = null;
                    TabGroupRule newRule = null;

                    if (result == DialogResult.OK || result == DialogResult.Yes)
                    {
                        string updatedName = dlg.RuleName;
                        int? updatedColor = dlg.SelectedColorIndex;
                        resultingColorIndex = updatedColor;

                        // Alias may already be saved from AliasConfirmed, but update in case the
                        // user went straight to the rule step (HideAliasStep = true).
                        if (enableAutoRename && enableAliasPrompt && !string.IsNullOrWhiteSpace(dlg.ServerAlias)
                            && !string.Equals(dlg.ServerAlias, existingAlias, StringComparison.Ordinal))
                        {
                            config.ServerAliases[server] = dlg.ServerAlias;
                        }

                        var context = new AddRuleContext
                        {
                            Config = config,
                            Server = server,
                            Database = database,
                            UseDb = useDb,
                            GroupName = updatedName,
                            ColorIndex = updatedColor
                        };
                        newRule = AddRuleAndSave(context);
                        changesApplied = true;

                        if (result == DialogResult.Yes)
                        {
                            OpenConfigInEditor();
                        }
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        // Save a silent null rule to suppress future auto-configure prompts
                        // for this connection without renaming or coloring the tab.
                        var silentContext = new AddRuleContext
                        {
                            Config = config,
                            Server = server,
                            Database = database,
                            UseDb = useDb,
                            GroupName = null,
                            ColorIndex = null
                        };
                        newRule = AddRuleAndSave(silentContext);
                        changesApplied = true;
                    }

                    if (enableColorWarning && result == DialogResult.OK && newRule != null)
                    {
                        if (resultingColorIndex.HasValue && IsColorUsedByOtherRule(config, newRule, resultingColorIndex.Value))
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

                    DialogClosed?.Invoke(new DialogClosedInfo
                    {
                        Result = result,
                        Server = server,
                        Database = database,
                        ChangesApplied = changesApplied
                    });
                }
            }
            catch (System.Exception ex)
            {
                EnvTabsLog.Error($"AutoConfig: Error showing prompt: {ex}");
            }
        }

        private static TabGroupRule AddRuleAndSave(AddRuleContext ctx)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var config = ctx.Config;

            // Check if we should remove the default example rule
            var exampleRule = config.ConnectionGroups.FirstOrDefault(g => g.GroupName == "Example: Exact Match");
            if (exampleRule != null)
            {
                config.ConnectionGroups.Remove(exampleRule);
            }

            // Check if we should remove the default example alias
            if (config.ServerAliases != null && config.ServerAliases.ContainsKey("MY-APP-SERVER"))
            {
                config.ServerAliases.Remove("MY-APP-SERVER");
            }
            
            // Renumber priorities: existing rules move to 20, 30, 40...
            var sorted = config.ConnectionGroups.OrderBy(x => x.Priority).ToList();
            int currentBase = 20;
            foreach (var rule in sorted)
            {
                rule.Priority = currentBase;
                currentBase += 10;
            }

            // Create new rule at 10
            var newRule = new TabGroupRule
            {
                GroupName = ctx.GroupName,
                Server = ctx.Server,
                Database = ctx.UseDb ? (ctx.Database ?? "%") : "%",
                Priority = 10,
                ColorIndex = ctx.ColorIndex
            };

            config.ConnectionGroups.Add(newRule);
            
            // Save
            SaveConfig(config);
            return newRule;
        }

        private static bool IsColorUsedByOtherRule(TabGroupConfig config, TabGroupRule currentRule, int colorIndex)
        {
            if (config?.ConnectionGroups == null) return false;

            foreach (var rule in config.ConnectionGroups)
            {
                if (!ReferenceEquals(rule, currentRule) && rule.ColorIndex == colorIndex)
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

        private static int FindNextColorValues(TabGroupConfig config)
        {
            // Requirement 1: "First available" - find a color not currently used.
            // Requirement 2: Strict "Fill the gap" logic without jumping.
            
            var used = new HashSet<int>(config.ConnectionGroups.Where(x => x.ColorIndex.HasValue).Select(x => x.ColorIndex.Value));
            
            // Find first hole in 0-15
            for (int i = 0; i <= 15; i++)
            {
                if (!used.Contains(i)) return i;
            }

            // If all used, return 0
            return 0;
        }

        private static void SaveConfig(TabGroupConfig config)
        {
            try
            {
                string path = TabGroupConfigLoader.GetUserConfigPath();
                var serializer = new DataContractJsonSerializer(typeof(TabGroupConfig), new DataContractJsonSerializerSettings
                {
                    UseSimpleDictionaryFormat = true
                });

                using (var stream = new MemoryStream())
                {
                    using (var writer = JsonReaderWriterFactory.CreateJsonWriter(stream, System.Text.Encoding.UTF8, true, true, "  "))
                    {
                        serializer.WriteObject(writer, config);
                        writer.Flush();
                    }
                    
                    // JsonReaderWriterFactory/DataContractJsonSerializer escapes forward slashes as \/.
                    // We decode to string, replace them for readability, and write back.
                    string json = System.Text.Encoding.UTF8.GetString(stream.ToArray());
                    json = json.Replace("\\/", "/");
                    
                    File.WriteAllText(path, json, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                }
            }
            catch (System.Exception ex)
            {
                EnvTabsLog.Info($"AutoConfig save failed: {ex.Message}");
            }
        }
    }
}
