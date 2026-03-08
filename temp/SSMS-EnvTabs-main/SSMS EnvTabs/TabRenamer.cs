using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace SSMS_EnvTabs
{
    internal class TabRenameContext
    {
        public uint Cookie { get; set; }
        public IVsWindowFrame Frame { get; set; }
        public string Server { get; set; }
        public string ServerAlias { get; set; }
        public string Database { get; set; }
        public string FrameCaption { get; set; }
        public string Moniker { get; set; }
    }

    internal static class TabRenamer
    {
        private static readonly Dictionary<uint, (string GroupName, int Index)> CookieToAssignment =
            new Dictionary<uint, (string GroupName, int Index)>();
        // Stores the SSMS-appended suffix (e.g. " - MYSERVER") learned from the original caption before any rename.
        private static readonly Dictionary<uint, string> CookieToSsmsSuffix =
            new Dictionary<uint, string>();
        // Stores the caption of each tab captured before we first rename it, so it can be restored on cancel/rule-removal.
        private static readonly Dictionary<uint, string> OriginalCaptionByCookie =
            new Dictionary<uint, string>();
        // Stores the pure user-given name (suffix and .sql stripped) captured the first time we rename a
        // non-default temp-file tab, so subsequent poll cycles use the stable original name as the
        // filename token instead of re-deriving it from an already-renamed caption (which would cause
        // the group suffix to be appended repeatedly).
        private static readonly Dictionary<uint, string> CookieToOriginalPureName =
            new Dictionary<uint, string>();

        // Matches SSMS's built-in sequential caption for new unsaved queries, e.g. "SQLQuery1".
        // After .sql is stripped this is the only form we need to match.
        private static readonly Regex SsmsDefaultQueryCaptionRegex =
            new Regex(@"^SQLQuery\d+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static void ForgetCookie(uint cookie)
        {
            if (cookie == 0) return;
            CookieToAssignment.Remove(cookie);
            CookieToSsmsSuffix.Remove(cookie);
            OriginalCaptionByCookie.Remove(cookie);
            CookieToOriginalPureName.Remove(cookie);
        }

        /// <summary>
        /// Reverts the caption of a tab to the caption captured before this extension first renamed it.
        /// Call this when a rule is removed or the dialog is cancelled so the tab goes back to its original name.
        /// </summary>
        public static void RevertCookieCaption(IVsWindowFrame frame, uint cookie, string currentCaption)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!CookieToAssignment.ContainsKey(cookie))
            {
                return; // Tab was never renamed by us; nothing to revert.
            }

            CookieToAssignment.Remove(cookie);
            CookieToSsmsSuffix.Remove(cookie);
            OriginalCaptionByCookie.TryGetValue(cookie, out string originalCaption);
            OriginalCaptionByCookie.Remove(cookie);
            CookieToOriginalPureName.Remove(cookie);

            if (frame == null || string.IsNullOrWhiteSpace(originalCaption))
            {
                return;
            }

            if (string.Equals(currentCaption, originalCaption, StringComparison.OrdinalIgnoreCase))
            {
                return; // Already showing the original name.
            }

            TrySetTabCaption(frame, originalCaption, out string propUsed);
            EnvTabsLog.Info($"Reverted tab caption ({propUsed}): cookie={cookie}, '{currentCaption}' -> '{originalCaption}'");
        }

        public static int ApplyRenamesOrThrow(IEnumerable<TabRenameContext> tabs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules, IReadOnlyList<TabRuleMatcher.CompiledManualRule> manualRules, string renameStyle = null, string savedFileRenameStyle = null, bool enableRemoveDotSql = true)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (rules == null) throw new ArgumentNullException(nameof(rules));
            
            // Default style if null
            if (string.IsNullOrWhiteSpace(renameStyle))
            {
                renameStyle = "[groupName][#]";
            }

            // Default saved file style if null (fallback to legacy behavior)
            string effectiveSavedStyle = savedFileRenameStyle;
            if (string.IsNullOrWhiteSpace(effectiveSavedStyle))
            {
                effectiveSavedStyle = "[filename] [groupName]";
            }

            var nextIndexByGroup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var assignment in CookieToAssignment.Values)
            {
                if (!nextIndexByGroup.TryGetValue(assignment.GroupName, out int next))
                {
                    next = 1;
                }

                nextIndexByGroup[assignment.GroupName] = Math.Max(next, assignment.Index + 1);
            }

            int renamed = 0;

            foreach (var tab in tabs)
            {
                if (tab.Frame == null) continue;

                var manualMatch = TabRuleMatcher.MatchManual(manualRules, tab.Moniker);
                string group = manualMatch?.GroupName ?? TabRuleMatcher.MatchGroup(rules, tab.Server, tab.Database);

                if (string.IsNullOrWhiteSpace(group))
                {
                    // No rename group, but still strip ".sql" from temp file captions when the
                    // connection matches a known (possibly silent/null-group) rule and the setting is on.
                    if (enableRemoveDotSql && RdtEventManager.IsTempFile(tab.Moniker))
                    {
                        bool hasKnownRule = manualMatch != null
                            || TabRuleMatcher.MatchRule(rules, tab.Server, tab.Database) != null;

                        if (hasKnownRule)
                        {
                            string stripped = StripSqlExtension(tab.FrameCaption, true);
                            if (!string.IsNullOrWhiteSpace(stripped)
                                && !string.Equals(tab.FrameCaption, stripped, StringComparison.OrdinalIgnoreCase))
                            {
                                int stripHr = TrySetTabCaption(tab.Frame, stripped, out string propUsed);
                                if (ErrorHandler.Succeeded(stripHr))
                                {
                                    renamed++;
                                    EnvTabsLog.Info($"Stripped .sql ({propUsed}): cookie={tab.Cookie}, '{tab.FrameCaption}' -> '{stripped}'");
                                }
                            }
                        }
                    }
                    continue;
                }

                if (!CookieToAssignment.TryGetValue(tab.Cookie, out var assignment) || !string.Equals(assignment.GroupName, group, StringComparison.OrdinalIgnoreCase))
                {
                    // Capture the pre-rename caption the very first time we rename this tab,
                    // so we can restore it if the rule is later cancelled or removed.
                    if (!CookieToAssignment.ContainsKey(tab.Cookie) && !OriginalCaptionByCookie.ContainsKey(tab.Cookie))
                    {
                        OriginalCaptionByCookie[tab.Cookie] = tab.FrameCaption;
                        CookieToSsmsSuffix[tab.Cookie] = DetectSsmsSuffix(tab.FrameCaption, tab.Moniker);
                    }

                    if (!nextIndexByGroup.TryGetValue(group, out int next))
                    {
                        next = 1;
                    }

                    assignment = (group, next);
                    CookieToAssignment[tab.Cookie] = assignment;
                    nextIndexByGroup[group] = next + 1;
                }

                string newCaption;
                if (manualMatch != null)
                {
                    // Manual match implies overwriting caption with GroupName (User Request)
                    newCaption = assignment.GroupName;
                }
                else if (RdtEventManager.IsTempFile(tab.Moniker))
                {
                    CookieToSsmsSuffix.TryGetValue(tab.Cookie, out string ssmsSuffix);
                    string pureName = GetPureName(tab.FrameCaption, ssmsSuffix, enableRemoveDotSql);

                    string generatedDefault = renameStyle
                        .Replace("[groupName]", assignment.GroupName)
                        .Replace("[#]", assignment.Index.ToString());

                    // When SSMS caption options (server/DB suffix) are disabled, the caption is just
                    // the raw temp filename (e.g. "5tekxtn0") instead of "SQLQuery1". Treat that as
                    // a default name too, so we don't mistake it for a user-customised caption.
                    string monikerBase = System.IO.Path.GetFileNameWithoutExtension(tab.Moniker ?? "");
                    bool isDefault = string.IsNullOrWhiteSpace(pureName)
                        || SsmsDefaultQueryCaptionRegex.IsMatch(pureName)
                        || string.Equals(pureName, generatedDefault, StringComparison.OrdinalIgnoreCase)
                        || (!string.IsNullOrEmpty(monikerBase) && string.Equals(pureName, monikerBase, StringComparison.OrdinalIgnoreCase));

                    if (!isDefault)
                    {
                        // Use the pure name captured the very first time we processed this tab so that
                        // subsequent poll cycles don't re-derive it from an already-renamed caption
                        // (which would cause the group suffix to be appended repeatedly, e.g.
                        // "test ILS" -> "test ILS ILS" -> ...).
                        //
                        // If the stored pure name exists, check whether the current caption still matches
                        // what we'd produce from it. If not, the user has manually renamed the tab again,
                        // so update the stored name to reflect their new choice.
                        if (CookieToOriginalPureName.TryGetValue(tab.Cookie, out string storedPureName))
                        {
                            string expectedCaption = BuildSavedStyleCaption(effectiveSavedStyle, storedPureName, assignment.GroupName, tab.Server, tab.ServerAlias, tab.Database);
                            if (!string.Equals(pureName, storedPureName, StringComparison.OrdinalIgnoreCase)
                                && !string.Equals(pureName, expectedCaption, StringComparison.OrdinalIgnoreCase))
                            {
                                // User has set a new caption — treat that as the new base name.
                                CookieToOriginalPureName[tab.Cookie] = pureName;
                            }
                        }
                        else
                        {
                            CookieToOriginalPureName[tab.Cookie] = pureName;
                        }
                        pureName = CookieToOriginalPureName[tab.Cookie];
                    }

                    newCaption = isDefault
                        ? generatedDefault
                        : BuildSavedStyleCaption(effectiveSavedStyle, pureName, assignment.GroupName, tab.Server, tab.ServerAlias, tab.Database);

                    EnvTabsLog.Info($"[Rename] cookie={tab.Cookie} pureName='{pureName}' isDefault={isDefault} -> '{newCaption}'");
                }
                else
                {
                    // Case 2: Saved File -> Use Saved File Configured Style
                    // Replace [filename], [groupName], [server], [serverAlias], [db]
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(tab.Moniker);
                    newCaption = BuildSavedStyleCaption(effectiveSavedStyle, fileName, assignment.GroupName, tab.Server, tab.ServerAlias, tab.Database);
                }
                
                // Skip if the visible name (SSMS suffix stripped) already matches what we'd set.
                // Note: do NOT strip .sql here — we want to still rename if the raw caption still contains ".sql"
                // but newCaption does not (i.e. when enableRemoveDotSql=true and the user renamed to "foo.sql").
                CookieToSsmsSuffix.TryGetValue(tab.Cookie, out string tabSuffix);
                string currentPure = GetPureName(tab.FrameCaption, tabSuffix ?? "", enableRemoveDotSql: false);
                if (!string.IsNullOrEmpty(tab.FrameCaption) && string.Equals(currentPure, newCaption, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                int hr = TrySetTabCaption(tab.Frame, newCaption, out string propertyNameUsed);
                if (ErrorHandler.Succeeded(hr))
                {
                    renamed++;
                    EnvTabsLog.Info($"Renamed ({propertyNameUsed}): cookie={tab.Cookie}, '{tab.FrameCaption}' -> '{newCaption}'");
                }
                else
                {
                    EnvTabsLog.Info($"Rename failed (hr=0x{hr:X8}): cookie={tab.Cookie}, caption='{tab.FrameCaption}', target='{newCaption}'");
                }
            }

            return renamed;
        }

        /// <summary>
        /// Learns the SSMS-appended suffix from a raw caption by comparing it to the known base
        /// (SQLQuery# for temp files, filename-without-extension for saved files).
        /// Returns e.g. " - MYSERVER" or "" if it can't be determined.
        /// </summary>
        private static string DetectSsmsSuffix(string rawCaption, string moniker)
        {
            string caption = StripSqlExtension(StripDirtyIndicators(rawCaption), true);
            if (string.IsNullOrEmpty(caption)) return "";

            if (RdtEventManager.IsTempFile(moniker))
            {
                var m = Regex.Match(caption, @"^(SQLQuery\d+)(.*)", RegexOptions.IgnoreCase);
                if (m.Success)
                    return m.Groups[2].Value;
            }
            else
            {
                string fileBase = System.IO.Path.GetFileNameWithoutExtension(moniker ?? "");
                if (!string.IsNullOrEmpty(fileBase) && caption.StartsWith(fileBase, StringComparison.OrdinalIgnoreCase))
                    return caption.Substring(fileBase.Length);
            }

            return "";
        }

        /// <summary>
        /// Strips .sql and the stored SSMS suffix from a live caption to get the user-visible pure name.
        /// </summary>
        private static string GetPureName(string rawCaption, string ssmsSuffix, bool enableRemoveDotSql)
        {
            string caption = StripDirtyIndicators(rawCaption);
            if (string.IsNullOrEmpty(caption)) return "";

            // Strip the SSMS suffix first so that a .sql embedded before it (e.g. "test.sql - SERVER")
            // ends up at the tail of the string before we try to remove it.
            if (!string.IsNullOrEmpty(ssmsSuffix))
            {
                while (caption.EndsWith(ssmsSuffix, StringComparison.OrdinalIgnoreCase))
                    caption = caption.Substring(0, caption.Length - ssmsSuffix.Length).TrimEnd();
            }

            caption = StripSqlExtension(caption, enableRemoveDotSql);

            return caption;
        }

        /// <summary>
        /// Removes SSMS dirty-state indicators from a raw caption.
        /// SSMS uses two styles: a trailing '*' and a trailing ' ⬤' (U+2B24).
        /// Multiple instances can accumulate if we rename without stripping them first.
        /// </summary>
        private static string StripDirtyIndicators(string name)
        {
            if (string.IsNullOrEmpty(name)) return name ?? "";
            string s = name.Trim();
            bool changed;
            do
            {
                changed = false;
                // Strip trailing asterisk dirty indicator
                while (s.EndsWith("*"))
                {
                    s = s.TrimEnd('*').TrimEnd();
                    changed = true;
                }
                // Strip trailing ⬤ (U+2B24) dirty indicator, with or without leading space
                while (s.EndsWith("\u2B24"))
                {
                    s = s.Substring(0, s.Length - 1).TrimEnd();
                    changed = true;
                }
            } while (changed);
            return s;
        }

        private static string StripSqlExtension(string name, bool enabled)
        {
            if (!enabled || string.IsNullOrEmpty(name))
            {
                return name;
            }

            // Anchor to end-of-string so that a .sql embedded in the middle of an already-renamed
            // caption (e.g. "test.sql ILS") is not incorrectly stripped, which would cause the
            // rename loop to corrupt the stored pure name on every poll cycle.
            return Regex.Replace(name, @"\.sql$", "", RegexOptions.IgnoreCase);
        }

        private static string BuildSavedStyleCaption(string savedStyle, string filenameToken, string groupName, string server, string serverAlias, string database)
        {
            return savedStyle
                .Replace("[filename]", filenameToken ?? "")
                .Replace("[groupName]", groupName ?? "")
                .Replace("[server]", server ?? "")
                .Replace("[serverAlias]", serverAlias ?? server ?? "")
                .Replace("[db]", database ?? "");
        }


        private static int TrySetTabCaption(IVsWindowFrame frame, string newCaption, out string propertyNameUsed)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            propertyNameUsed = null;

            int hr = frame.SetProperty((int)__VSFPROPID.VSFPROPID_OwnerCaption, newCaption);
            if (ErrorHandler.Succeeded(hr))
            {
                propertyNameUsed = nameof(__VSFPROPID.VSFPROPID_OwnerCaption);
                return hr;
            }

            int hr2 = frame.SetProperty((int)__VSFPROPID.VSFPROPID_Caption, newCaption);
            if (ErrorHandler.Succeeded(hr2))
            {
                propertyNameUsed = nameof(__VSFPROPID.VSFPROPID_Caption);
                return hr2;
            }

            int hr3 = frame.SetProperty((int)__VSFPROPID.VSFPROPID_EditorCaption, newCaption);
            if (ErrorHandler.Succeeded(hr3))
            {
                propertyNameUsed = nameof(__VSFPROPID.VSFPROPID_EditorCaption);
                return hr3;
            }

            return hr;
        }
    }
}
