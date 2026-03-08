using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace SSMS_EnvTabs
{
    internal sealed class ColorByRegexConfigWriter
    {
        private const string BeginMarker = "// SSMS EnvTabs: BEGIN generated";
        private const string EndMarker = "// SSMS EnvTabs: END generated";
        private const int ResolveRetryMax = 12;
        private const int ResolveRetryDelayMs = 500;
        private const int TempScanBackoffMs = 3000;
        private const double CreationSkewSeconds = 60.0;
        private const double CreationMaxWindowSeconds = 60.0;

        private string resolvedConfigPath;
        private DateTime? firstDocSeenUtc;
        private bool resolveRetryScheduled;
        private int resolveRetryCount;
        private DateTime? lastTempScanUtc;
        private readonly object tempScanLock = new object();
        private List<RdtEventManager.OpenDocumentInfo> lastDocsSnapshot;
        private IReadOnlyList<TabRuleMatcher.CompiledRule> lastRulesSnapshot;
        private IReadOnlyList<TabRuleMatcher.CompiledManualRule> lastManualRulesSnapshot;

        public event Action<string> ConfigPathResolved;

        public void UpdateFromSnapshot(IEnumerable<RdtEventManager.OpenDocumentInfo> docs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules, IReadOnlyList<TabRuleMatcher.CompiledManualRule> manualRules)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (docs == null)
            {
                return;
            }

            lastDocsSnapshot = docs as List<RdtEventManager.OpenDocumentInfo> ?? docs.ToList();
            lastRulesSnapshot = rules;
            lastManualRulesSnapshot = manualRules;

            if (!firstDocSeenUtc.HasValue && lastDocsSnapshot.Count > 0)
            {
                firstDocSeenUtc = DateTime.UtcNow;
            }

            var safeRules = rules ?? Array.Empty<TabRuleMatcher.CompiledRule>();
            bool hasManualRules = manualRules != null && manualRules.Count > 0;
            EnvTabsLog.Verbose($"ColorByRegexConfigWriter.cs::UpdateFromSnapshot - Rules={safeRules.Count}, ManualRules={(manualRules?.Count ?? 0)}, Docs={lastDocsSnapshot.Count}");
            if (safeRules.Count == 0 && !hasManualRules)
            {
                EnvTabsLog.Verbose("ColorByRegexConfigWriter.cs::UpdateFromSnapshot - No rules to write. Skipping.");
                return;
            }

            string configPath = ResolveConfigPath(lastDocsSnapshot.Select(d => d?.Moniker));
            if (string.IsNullOrWhiteSpace(configPath))
            {
                EnvTabsLog.Verbose("ColorByRegexConfigWriter.cs::UpdateFromSnapshot - Config path not resolved. Skipping.");
                ScheduleResolveRetryIfNeeded();
                StartTempScanIfNeeded();
                return;
            }
            resolveRetryCount = 0;
            resolveRetryScheduled = false;
            EnvTabsLog.Verbose($"ColorByRegexConfigWriter.cs::UpdateFromSnapshot - ConfigPath='{configPath}'");

            try
            {
                ConfigPathResolved?.Invoke(configPath);
            }
            catch
            {
                // Best-effort
            }

            var groupToPaths = safeRules.Count > 0 ? BuildGroupPathMap(lastDocsSnapshot, safeRules) : null;
            EnvTabsLog.Verbose($"ColorByRegexConfigWriter.cs::UpdateFromSnapshot - GroupToPaths={(groupToPaths?.Count ?? 0)}");
            var rulesSnapshot = safeRules;
            var manualRulesSnapshot = manualRules;
            var groupToPathsSnapshot = groupToPaths;
            var configPathSnapshot = configPath;

            _ = System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    string newContent = BuildConfigContent(configPathSnapshot, groupToPathsSnapshot, rulesSnapshot, manualRulesSnapshot);
                    WriteIfChanged(configPathSnapshot, newContent);
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::UpdateFromSnapshot - Background write failed: {ex.Message}");
                }
            });
        }

        private struct RegexEntry
        {
            public string Pattern;
            public int Priority;
        }

        private string BuildConfigContent(string existingPath, Dictionary<string, SortedSet<string>> groupToPaths, IReadOnlyList<TabRuleMatcher.CompiledRule> rules, IReadOnlyList<TabRuleMatcher.CompiledManualRule> manualRules)
        {
            var entries = new List<RegexEntry>();

            // Preserve content outside markers.
            string existing = File.Exists(existingPath) ? File.ReadAllText(existingPath) : string.Empty;
            var groupIdColorMap = TryLoadGroupIdColorMap(existingPath);
            var overrideColorByBaseRegex = BuildOverrideColorByBaseRegex(existing, groupIdColorMap);

            // Manual rules
            if (manualRules != null)
            {
                foreach (var m in manualRules)
                {
                    string original = m.OriginalPattern;
                    string baseRegex = StripSaltComment(original);
                    string line = original;
                    if (overrideColorByBaseRegex != null && overrideColorByBaseRegex.TryGetValue(baseRegex, out int overrideColor))
                    {
                        try
                        {
                            line = ApplyColorIndex(baseRegex, overrideColor);
                        }
                        catch
                        {
                            line = original;
                        }
                    }
                    else if (m.ColorIndex.HasValue)
                    {
                        try
                        {
                            line = ApplyColorIndex(baseRegex, m.ColorIndex.Value);
                        }
                        catch
                        {
                            line = original;
                        }
                    }
                    entries.Add(new RegexEntry { Pattern = line, Priority = m.Priority });
                }
            }

            // Generated rules
            if (groupToPaths != null)
            {
                foreach (var rule in rules)
                {
                    // Skip null-GroupName (silent) rules and null-ColorIndex rules.
                    if (string.IsNullOrWhiteSpace(rule.GroupName) || !rule.ColorIndex.HasValue) continue;
                    if (groupToPaths.TryGetValue(rule.GroupName, out var paths) && paths.Count > 0)
                    {
                        string baseRegex = BuildBaseRegex(paths);
                        string line = BuildResolvedRegexLine(baseRegex, rule, overrideColorByBaseRegex, groupIdColorMap);
                        entries.Add(new RegexEntry { Pattern = line, Priority = rule.Priority });
                    }
                }
            }

            var sortedEntries = entries.OrderBy(e => e.Priority).ToList();

            var generatedLines = new List<string>();

            generatedLines.Add(BeginMarker);
            foreach(var e in sortedEntries)
            {
                generatedLines.Add(e.Pattern);
            }
            generatedLines.Add(EndMarker);

            return ReplaceOrAppendBlock(existing, generatedLines);
        }        

        internal Dictionary<string, int> TryGetOverrideColorsByBaseRegex(string configPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(configPath) || !File.Exists(configPath))
                {
                    return null;
                }

                string existing = File.ReadAllText(configPath);
                var groupIdColorMap = TryLoadGroupIdColorMap(configPath);
                return BuildOverrideColorByBaseRegex(existing, groupIdColorMap);
            }
            catch
            {
                return null;
            }
        }

        internal Dictionary<string, string> BuildGroupNameToBaseRegex(IEnumerable<RdtEventManager.OpenDocumentInfo> docs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            if (docs == null || rules == null || rules.Count == 0)
            {
                return null;
            }

            var groupToPaths = BuildGroupPathMap(docs, rules);
            if (groupToPaths == null || groupToPaths.Count == 0)
            {
                return null;
            }

            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var rule in rules)
            {
                groupToPaths.TryGetValue(rule.GroupName, out var paths);
                string baseRegex = BuildBaseRegex(paths);
                if (!string.IsNullOrWhiteSpace(baseRegex))
                {
                    result[rule.GroupName] = StripSaltComment(baseRegex);
                }
            }

            return result.Count > 0 ? result : null;
        }

        internal string BuildGeneratedBlockPreview(IEnumerable<RdtEventManager.OpenDocumentInfo> docs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            if (docs == null || rules == null || rules.Count == 0)
            {
                return string.Empty;
            }

            var groupToPaths = BuildGroupPathMap(docs, rules);
            var lines = BuildGeneratedBlock(groupToPaths, rules);
            return string.Join("\n", lines);
        }

        private Dictionary<string, SortedSet<string>> BuildGroupPathMap(IEnumerable<RdtEventManager.OpenDocumentInfo> docs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            var map = new Dictionary<string, SortedSet<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var rule in rules)
            {
                // Skip null-GroupName (silent) rules — they do not generate color entries.
                if (string.IsNullOrWhiteSpace(rule.GroupName)) continue;
                if (!map.ContainsKey(rule.GroupName))
                {
                    map[rule.GroupName] = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
                }
            }

            foreach (var doc in docs)
            {
                if (doc == null)
                {
                    continue;
                }

                string group = TabRuleMatcher.MatchGroup(rules, doc.Server, doc.Database);
                if (string.IsNullOrWhiteSpace(group))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(doc.Moniker) || !Path.IsPathRooted(doc.Moniker))
                {
                    continue;
                }

                if (!map.TryGetValue(group, out var set))
                {
                    set = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
                    map[group] = set;
                }

                // Use only the filename so colors follow file moves.
                try 
                {
                    string fileName = Path.GetFileName(doc.Moniker);
                    if (!string.IsNullOrEmpty(fileName))
                    {
                        set.Add(fileName);
                    }
                }
                catch
                {
                    // If path is invalid, skip
                }
            }

            return map;
        }

        private static List<string> BuildGeneratedBlock(Dictionary<string, SortedSet<string>> groupToPaths, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            var lines = new List<string> { BeginMarker };

            foreach (var rule in rules.OrderBy(r => r.Priority).ThenBy(r => r.GroupName, StringComparer.OrdinalIgnoreCase))
            {
                // Skip null-GroupName (silent) rules and null-ColorIndex rules — they produce no color entry.
                if (string.IsNullOrWhiteSpace(rule.GroupName) || !rule.ColorIndex.HasValue) continue;
                groupToPaths.TryGetValue(rule.GroupName, out var paths);
                string baseRegex = BuildBaseRegex(paths);
                lines.Add(BuildResolvedRegexLine(baseRegex, rule, null));
            }

            lines.Add(EndMarker);
            return lines;
        }

        private static string BuildResolvedRegexLine(string baseRegex, TabRuleMatcher.CompiledRule rule, Dictionary<string, int> overrideColorByBaseRegex, Dictionary<int, int> groupIdColorMap = null)
        {
            if (string.IsNullOrWhiteSpace(baseRegex))
            {
                return "(?!)";
            }

            string normalizedBase = StripSaltComment(baseRegex);

            // Determine the effective target color: prefer what the config says; fall back to
            // the SSMS JSON override only when the two agree (or the config has no value).
            int? targetColor = rule.ColorIndex;

            if (overrideColorByBaseRegex != null && overrideColorByBaseRegex.TryGetValue(normalizedBase, out int overrideColor))
            {
                if (!targetColor.HasValue || overrideColor == targetColor.Value)
                {
                    EnvTabsLog.Verbose($"BuildResolvedRegexLine: Applying SSMS override colorIndex={overrideColor} for group='{rule.GroupName}'");
                    targetColor = overrideColor;
                }
                else
                {
                    // The SSMS JSON entry is stale: it was recorded for a previous visit to that
                    // color. The JSON accumulates all past entries, so after cycling 0→1→2→0 the
                    // base regex for color 0 can collide with an old entry pointing to color 1.
                    // Keep the config value as-is and fall through to pick a collision-free salt.
                    EnvTabsLog.Info($"BuildResolvedRegexLine: Ignoring stale SSMS override colorIndex={overrideColor} (config colorIndex={targetColor.Value}) for group='{rule.GroupName}'");
                }
            }

            if (!targetColor.HasValue)
            {
                return normalizedBase;
            }

            // Build the set of hashes that SSMS already has mapped to a DIFFERENT color.
            // Solve() will skip any candidate whose hash is in this set so that SSMS can
            // never silently override our chosen color by matching a stale JSON entry.
            System.Collections.Generic.HashSet<int> forbiddenHashes = null;
            if (groupIdColorMap != null && groupIdColorMap.Count > 0)
            {
                foreach (var kvp in groupIdColorMap)
                {
                    if (kvp.Value != targetColor.Value)
                    {
                        if (forbiddenHashes == null)
                        {
                            forbiddenHashes = new System.Collections.Generic.HashSet<int>();
                        }

                        forbiddenHashes.Add(kvp.Key);
                    }
                }

                if (forbiddenHashes != null)
                {
                    EnvTabsLog.Verbose($"BuildResolvedRegexLine: {forbiddenHashes.Count} forbidden hash(es) for group='{rule.GroupName}' targetColor={targetColor.Value}");
                }
            }

            return ApplyColorIndex(normalizedBase, targetColor.Value, forbiddenHashes);
        }

        private static string BuildBaseRegex(SortedSet<string> paths)
        {
            if (paths == null || paths.Count == 0)
            {
                return "(?!)";
            }

            var pathList = paths.ToList();
            bool allSql = pathList.All(path => path.EndsWith(".sql", StringComparison.OrdinalIgnoreCase));
            IEnumerable<string> escaped;
            string suffix = string.Empty;
            if (allSql)
            {
                escaped = pathList.Select(path => Regex.Escape(Path.GetFileNameWithoutExtension(path)));
                suffix = "\\.sql";
            }
            else
            {
                escaped = pathList.Select(Regex.Escape);
            }

            // Match path ending with the filename.
            return $"(?:^|[\\\\/])(?:{string.Join("|", escaped)}){suffix}$";
        }

        private static string ApplyColorIndex(string baseRegex, int targetColorIndex, System.Collections.Generic.ICollection<int> forbiddenHashes = null)
        {
            if (string.IsNullOrWhiteSpace(baseRegex))
            {
                return baseRegex;
            }

            if (targetColorIndex < 0 || targetColorIndex > 15)
            {
                return baseRegex;
            }

            int currentHash = TabGroupColorSolver.GetSsmsStableHashCode(baseRegex);
            int currentColor = Math.Abs(currentHash) % 16;
            // Only use the unsalted base regex if it already produces the right color AND its
            // hash isn't recorded in the SSMS JSON with a different color (which would make SSMS
            // override it back to the wrong color).
            if (currentColor == targetColorIndex && (forbiddenHashes == null || !forbiddenHashes.Contains(currentHash)))
            {
                return baseRegex;
            }

            string newSalt = TabGroupColorSolver.Solve(baseRegex, targetColorIndex, forbiddenHashes);
            if (!string.IsNullOrWhiteSpace(newSalt))
            {
                return baseRegex + $"(?#salt:{newSalt})";
            }

            return baseRegex;
        }

        private static string StripSaltComment(string regexLine)
        {
            if (string.IsNullOrWhiteSpace(regexLine))
            {
                return regexLine;
            }

            string trimmed = regexLine.Trim();
            return Regex.Replace(trimmed, "\\(\\?#salt:[^)]*\\)", string.Empty);
        }

        private static Dictionary<int, int> TryLoadGroupIdColorMap(string configPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(configPath))
                {
                    return null;
                }

                string dir = Path.GetDirectoryName(configPath);
                if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
                {
                    return null;
                }

                var candidates = Directory.GetFiles(dir, "customized-groupid-color-*.json");
                if (candidates == null || candidates.Length == 0)
                {
                    return null;
                }

                string latest = candidates
                    .Select(path => new { Path = path, LastWriteUtc = File.GetLastWriteTimeUtc(path) })
                    .OrderByDescending(x => x.LastWriteUtc)
                    .Select(x => x.Path)
                    .FirstOrDefault();

                if (string.IsNullOrWhiteSpace(latest) || !File.Exists(latest))
                {
                    return null;
                }

                using (var stream = File.OpenRead(latest))
                {
                    var serializer = new DataContractJsonSerializer(typeof(GroupIdColorMapFile), new DataContractJsonSerializerSettings
                    {
                        UseSimpleDictionaryFormat = true
                    });

                    var data = serializer.ReadObject(stream) as GroupIdColorMapFile;
                    if (data?.ColorMap == null || data.ColorMap.Count == 0)
                    {
                        return null;
                    }

                    var result = new Dictionary<int, int>();
                    foreach (var kvp in data.ColorMap)
                    {
                        if (kvp.Value == null)
                        {
                            continue;
                        }

                        int groupId = kvp.Value.GroupId;
                        int colorIndex = kvp.Value.ColorIndex;
                        if (colorIndex < 0 || colorIndex > 15)
                        {
                            continue;
                        }

                        result[groupId] = colorIndex;
                    }

                    return result.Count > 0 ? result : null;
                }
            }
            catch
            {
                return null;
            }
        }

        private static Dictionary<string, int> BuildOverrideColorByBaseRegex(string existingContent, Dictionary<int, int> groupIdColorMap)
        {
            if (groupIdColorMap == null || groupIdColorMap.Count == 0)
            {
                return null;
            }

            EnvTabsLog.Verbose($"BuildOverrideColorByBaseRegex: SSMS JSON has {groupIdColorMap.Count} entries: [{string.Join(", ", groupIdColorMap.Select(kvp => $"hash={kvp.Key}→colorIndex={kvp.Value}"))}]");

            var lines = SplitLines(existingContent);
            int begin = lines.FindIndex(line => string.Equals(line?.TrimEnd(), BeginMarker, StringComparison.Ordinal));
            int end = lines.FindIndex(line => string.Equals(line?.TrimEnd(), EndMarker, StringComparison.Ordinal));

            if (begin < 0 || end < begin)
            {
                return null;
            }

            var overrides = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = begin + 1; i < end; i++)
            {
                string line = lines[i]?.Trim();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                if (line.StartsWith("//", StringComparison.Ordinal))
                {
                    continue;
                }

                int groupId = TabGroupColorSolver.GetSsmsStableHashCode(line);
                if (groupIdColorMap.TryGetValue(groupId, out int desiredColor))
                {
                    string baseRegex = StripSaltComment(line);
                    if (!string.IsNullOrWhiteSpace(baseRegex))
                    {
                        EnvTabsLog.Verbose($"BuildOverrideColorByBaseRegex: Matched line hash={groupId}→colorIndex={desiredColor} (line='{line}')");
                        overrides[baseRegex] = desiredColor;
                    }
                }
                else
                {
                    EnvTabsLog.Verbose($"BuildOverrideColorByBaseRegex: No SSMS JSON entry for line hash={groupId} (line='{line}')");
                }
            }

            return overrides.Count > 0 ? overrides : null;
        }

        [DataContract]
        private sealed class GroupIdColorMapFile
        {
            [DataMember(Name = "Version")]
            public int Version { get; set; }

            [DataMember(Name = "ColorMap")]
            public Dictionary<string, GroupIdColorEntry> ColorMap { get; set; }
        }

        [DataContract]
        private sealed class GroupIdColorEntry
        {
            [DataMember(Name = "GroupId")]
            public int GroupId { get; set; }

            [DataMember(Name = "ColorIndex")]
            public int ColorIndex { get; set; }
        }

        private static string ReplaceOrAppendBlock(string existing, List<string> generatedLines)
        {
            var lines = SplitLines(existing);
            int begin = lines.FindIndex(line => string.Equals(line?.TrimEnd(), BeginMarker, StringComparison.Ordinal));
            int end = lines.FindIndex(line => string.Equals(line?.TrimEnd(), EndMarker, StringComparison.Ordinal));

            var result = new List<string>();
            if (begin >= 0 && end >= begin)
            {
                result.AddRange(lines.Take(begin));
                result.AddRange(generatedLines);
                result.AddRange(lines.Skip(end + 1));
            }
            else
            {
                result.AddRange(lines);
                if (result.Count > 0 && !string.IsNullOrWhiteSpace(result[result.Count - 1]))
                {
                    result.Add(string.Empty);
                }

                result.AddRange(generatedLines);
            }

            return string.Join(Environment.NewLine, result);
        }

        private static List<string> SplitLines(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new List<string>();
            }

            return text
                .Replace("\r\n", "\n")
                .Replace("\r", "\n")
                .Split('\n')
                .ToList();
        }

        private static void WriteIfChanged(string path, string content)
        {
            string existing = File.Exists(path) ? File.ReadAllText(path) : string.Empty;
            if (string.Equals(NormalizeNewlines(existing), NormalizeNewlines(content ?? string.Empty), StringComparison.Ordinal))
            {
                EnvTabsLog.Verbose($"ColorByRegexConfigWriter.cs::WriteIfChanged - No changes detected for '{path}'.");
                return;
            }

            const int maxAttempts = 6;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    string dir = Path.GetDirectoryName(path);
                    if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    string tmp = path + ".tmp";
                    EnvTabsLog.Verbose($"ColorByRegexConfigWriter.cs::WriteIfChanged - Writing file '{path}' (attempt {attempt}/{maxAttempts}).");
                    File.WriteAllText(tmp, content ?? string.Empty, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

                    if (File.Exists(path))
                    {
                        File.Replace(tmp, path, null, ignoreMetadataErrors: true);
                    }
                    else
                    {
                        File.Move(tmp, path);
                    }

                    return;
                }
                catch (IOException) when (attempt < maxAttempts)
                {
                    Thread.Sleep(150 * attempt);
                }
                catch (UnauthorizedAccessException) when (attempt < maxAttempts)
                {
                    Thread.Sleep(150 * attempt);
                }
            }
        }

        private static string NormalizeNewlines(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            return text.Replace("\r\n", "\n").Replace("\r", "\n");
        }

        private void ScheduleResolveRetryIfNeeded()
        {
            if (resolveRetryScheduled)
            {
                return;
            }

            if (resolveRetryCount >= ResolveRetryMax)
            {
                return;
            }

            resolveRetryScheduled = true;
            int attempt = ++resolveRetryCount;
            EnvTabsLog.Verbose($"ColorByRegexConfigWriter.cs::UpdateFromSnapshot - Resolve retry scheduled ({attempt}/{ResolveRetryMax}) in {ResolveRetryDelayMs}ms.");

            _ = System.Threading.Tasks.Task.Run(async () =>
            {
                try
                {
                    await System.Threading.Tasks.Task.Delay(ResolveRetryDelayMs).ConfigureAwait(false);
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                    resolveRetryScheduled = false;

                    if (lastDocsSnapshot != null)
                    {
                        UpdateFromSnapshot(lastDocsSnapshot, lastRulesSnapshot, lastManualRulesSnapshot);
                    }
                }
                catch
                {
                    resolveRetryScheduled = false;
                }
            });
        }

        private string ResolveConfigPath(IEnumerable<string> monikers)
        {
            if (!string.IsNullOrWhiteSpace(resolvedConfigPath) && File.Exists(resolvedConfigPath))
            {
                return resolvedConfigPath;
            }

            // Try from open documents first.
            foreach (var moniker in monikers ?? Enumerable.Empty<string>())
            {
                string candidate = TryGetConfigPathFromMoniker(moniker);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    resolvedConfigPath = candidate;
                    return candidate;
                }
            }

            return null;
        }

        private static string TryGetConfigPathFromMoniker(string moniker)
        {
            if (string.IsNullOrWhiteSpace(moniker) || !Path.IsPathRooted(moniker))
            {
                return null;
            }

            try
            {
                string dir = Path.GetDirectoryName(moniker);
                if (string.IsNullOrWhiteSpace(dir))
                {
                    return null;
                }

                var current = new DirectoryInfo(dir);

                string guidRoot = TryGetTempGuidRoot(current);
                if (!string.IsNullOrWhiteSpace(guidRoot))
                {
                    EnvTabsLog.Verbose($"ColorByRegexConfigWriter.cs::TryGetConfigPathFromMoniker - Temp GUID root='{guidRoot}'");
                    string candidate = Path.Combine(guidRoot, "ColorByRegexConfig.txt");
                    if (File.Exists(candidate))
                    {
                        if (HasSiblingVFolder(candidate))
                        {
                            return candidate;
                        }

                        EnvTabsLog.Verbose($"ColorByRegexConfigWriter.cs::TryGetConfigPathFromMoniker - Skipping temp candidate (missing v# folder): '{candidate}'");
                    }
                }

                var walker = current;
                for (int i = 0; i < 4 && walker != null; i++)
                {
                    string candidate = Path.Combine(walker.FullName, "ColorByRegexConfig.txt");
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }

                    walker = walker.Parent;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private static string TryScanTempForConfig(DateTime? referenceUtc)
        {
            try
            {
                if (!referenceUtc.HasValue)
                {
                    EnvTabsLog.Verbose("ColorByRegexConfigWriter.cs::TryScanTempForConfig - No firstDocSeenUtc; skipping temp scan.");
                    return null;
                }

                string temp = Path.GetTempPath();
                if (string.IsNullOrWhiteSpace(temp) || !Directory.Exists(temp))
                {
                    return null;
                }

                // SSMS places the file in a GUID-named Temp subdirectory.

                var dirs = Directory.GetDirectories(temp)
                    .Where(d => IsGuidDirectoryName(Path.GetFileName(d)))
                    .Select(d => new { Dir = d, CreatedUtc = Directory.GetCreationTimeUtc(d) })
                    .OrderBy(d => d.CreatedUtc)
                    .ToList();

                foreach (var entry in dirs)
                {
                    try
                    {
                        string candidate = Path.Combine(entry.Dir, "ColorByRegexConfig.txt");
                        bool fileExists = File.Exists(candidate);

                        DateTime dirCreateUtc = entry.CreatedUtc;
                        double deltaSeconds = (dirCreateUtc - referenceUtc.Value).TotalSeconds;

                        // Skip folders outside the creation window.
                        if (deltaSeconds < -CreationSkewSeconds)
                        {
                            continue;
                        }

                        if (deltaSeconds > CreationMaxWindowSeconds)
                        {
                            continue;
                        }

                        EnvTabsLog.Verbose($"ColorByRegexConfigWriter.cs::TryScanTempForConfig - FirstDocSeenUtc={referenceUtc:O}, RegexFolderCreateUtc={dirCreateUtc:O}, DeltaSeconds={deltaSeconds:0.###}, FileExists={fileExists}, Folder='{entry.Dir}'");
                        if (fileExists)
                        {
                            if (HasSiblingVFolder(candidate))
                            {
                                return candidate;
                            }

                            EnvTabsLog.Verbose($"ColorByRegexConfigWriter.cs::TryScanTempForConfig - Skipping candidate (missing v# folder): '{candidate}'");
                        }
                    }
                    catch
                    {
                        // Ignore access errors
                    }
                }
                EnvTabsLog.Verbose($"ColorByRegexConfigWriter.cs::TryScanTempForConfig - No candidate found. FirstDocSeenUtc={referenceUtc:O}");
                return null;
            }
            catch
            {
                return null;
            }
        }

        private void StartTempScanIfNeeded()
        {
            if (!TryBeginTempScan())
            {
                return;
            }

            _ = System.Threading.Tasks.Task.Run(async () =>
            {
                try
                {
                    string fallback = TryScanTempForConfig(firstDocSeenUtc);
                    if (string.IsNullOrWhiteSpace(fallback))
                    {
                        return;
                    }

                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    resolvedConfigPath = fallback;
                    if (lastDocsSnapshot != null)
                    {
                        UpdateFromSnapshot(lastDocsSnapshot, lastRulesSnapshot, lastManualRulesSnapshot);
                    }
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::StartTempScanIfNeeded - Scan failed: {ex.Message}");
                }
            });
        }

        private bool TryBeginTempScan()
        {
            lock (tempScanLock)
            {
                var now = DateTime.UtcNow;
                if (lastTempScanUtc.HasValue && (now - lastTempScanUtc.Value).TotalMilliseconds < TempScanBackoffMs)
                {
                    return false;
                }

                lastTempScanUtc = now;
                return true;
            }
        }

        private static string TryGetTempGuidRoot(DirectoryInfo start)
        {
            if (start == null)
            {
                return null;
            }

            string tempRoot = Path.GetTempPath();
            var current = start;
            while (current != null)
            {
                if (IsGuidDirectoryName(current.Name))
                {
                    if (string.IsNullOrWhiteSpace(tempRoot))
                    {
                        return current.FullName;
                    }

                    if (current.FullName.StartsWith(tempRoot, StringComparison.OrdinalIgnoreCase))
                    {
                        return current.FullName;
                    }
                }

                current = current.Parent;
            }

            return null;
        }

        private static bool IsGuidDirectoryName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            return Guid.TryParse(name, out _);
        }

        private static bool HasSiblingVFolder(string configPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(configPath))
                {
                    return false;
                }

                string dir = Path.GetDirectoryName(configPath);
                if (string.IsNullOrWhiteSpace(dir))
                {
                    return false;
                }

                foreach (var subDir in Directory.GetDirectories(dir))
                {
                    string name = Path.GetFileName(subDir);
                    if (IsVersionFolderName(name))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsVersionFolderName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            if (name.Length < 2 || name[0] != 'v')
            {
                return false;
            }

            for (int i = 1; i < name.Length; i++)
            {
                if (!char.IsDigit(name[i]))
                {
                    return false;
                }
            }

            return true;
        }
    }
}
