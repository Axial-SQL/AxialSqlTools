using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    internal static class UpdateChecker
    {
        private const string ProductName = "AxialSqlTools";
        private const string DisplayName = "Axial SQL Tools";
        private const string ReleasesApiUrl = "https://api.github.com/repos/Axial-SQL/AxialSqlTools/releases/latest";
        private const string ReleasePageUrl = "https://github.com/Axial-SQL/AxialSqlTools/releases/latest";
        private const string StagedVsixFilePattern = "AxialSqlTools-*.vsix";
        private const string StagedZipFilePattern = "AxialSqlTools-*.zip";
        private const string ExpectedVsixName = "AxialSqlTools.vsix";
        private const string Sha256DigestPrefix = "sha256:";
        private static readonly TimeSpan DeferredUpdateStageWaitTimeout = TimeSpan.FromSeconds(10);
#if DEBUG
        private const string ForceUpdateAvailableEnvironmentVariable = "AXIALSQLTOOLS_FORCE_UPDATE_AVAILABLE";
#endif

        private static int pendingStartupCheck;
        private static readonly object diagnosticsLock = new object();
        private static readonly object updateStateLock = new object();

        private static string lastUpdateResult = "No update check has run yet.";
        private static string stagedVsixPath;
        private static GitHubRelease stagedRelease;
        private static bool pendingUpdateOnClose;
        private static bool stageDownloadInProgress;
        private static bool stageDownloadFailed;
        private static Task stageDownloadTask = Task.CompletedTask;
        private static Task cleanupDownloadedVsixFilesTask = Task.CompletedTask;
        private static UpdateInfoBar activeInfoBar;

        internal static event Action LastUpdateResultChanged;

        internal static string LastUpdateResult
        {
            get
            {
                lock (diagnosticsLock)
                {
                    return lastUpdateResult;
                }
            }
        }

        internal static void Log(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            try
            {
                AxialSqlToolsPackage._logger?.Info(message);
            }
            catch
            {
            }
        }

        public static void ScheduleCheck(AsyncPackage package, bool enableUpdateChecks)
        {
            if (package == null)
            {
                return;
            }

            Task cleanupTask = ScheduleCleanupDownloadedVsixFiles();

            if (!enableUpdateChecks)
            {
                Log("Update check skipped: disabled by settings.");
                SetLastUpdateResult("Update check skipped because it is disabled in settings.");
                return;
            }

            Interlocked.Exchange(ref pendingStartupCheck, 1);
            Log("Update check scheduled shortly after package initialization.");
            SetLastUpdateResult("Startup update check scheduled.");

            _ = package.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    var token = package.DisposalToken;
                    await cleanupTask;
                    await Task.Delay(TimeSpan.FromMilliseconds(900), token);

                    if (Interlocked.CompareExchange(ref pendingStartupCheck, 0, 1) != 1)
                    {
                        return;
                    }

                    Log("Running startup update check after initialization delay.");
                    SetLastUpdateResult("Running startup update check.");
                    await CheckForUpdatesAsync(package, token, showUpToDate: false);
                }
                catch (OperationCanceledException)
                {
                }
                catch (Exception ex)
                {
                    Log($"Startup update check failed: {ex.Message}");
                    SetLastUpdateResult($"Startup update check failed: {ex.Message}");
                }
            });
        }

        public static void CheckNow(AsyncPackage package, bool ignoreSettings = true)
        {
            if (package == null)
            {
                return;
            }

            if (!ignoreSettings && !SettingsManager.GetEnableUpdateChecks())
            {
                Log("Manual update check skipped: disabled by settings.");
                SetLastUpdateResult("Manual update check skipped because it is disabled in settings.");
                return;
            }

            _ = package.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    Log("Manual update check started.");
                    SetLastUpdateResult("Manual update check started.");
                    await CheckForUpdatesAsync(package, package.DisposalToken, showUpToDate: true);
                }
                catch (OperationCanceledException)
                {
                }
                catch (Exception ex)
                {
                    Log($"Manual update check failed: {ex.Message}");
                    SetLastUpdateResult($"Manual update check failed: {ex.Message}");
                }
            });
        }

        internal static void LaunchDeferredUpdateOnClose()
        {
            Task downloadTask;
            lock (updateStateLock)
            {
                if (!pendingUpdateOnClose)
                {
                    return;
                }

                downloadTask = stageDownloadInProgress ? stageDownloadTask : null;
            }

            if (downloadTask != null && !downloadTask.IsCompleted)
            {
                Log($"Deferred update on close: waiting up to {DeferredUpdateStageWaitTimeout.TotalSeconds:0} seconds for staging to finish.");
                try
                {
                    if (!downloadTask.Wait(DeferredUpdateStageWaitTimeout))
                    {
                        lock (updateStateLock)
                        {
                            pendingUpdateOnClose = false;
                        }

                        Log("Deferred update on close: staging did not finish before timeout. Skipping.");
                        SetLastUpdateResult("Deferred update skipped because the VSIX download did not finish before SSMS closed.");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Log($"Deferred update on close: staging wait failed: {ex.Message}");
                }
            }

            string vsixPath;
            GitHubRelease release;
            lock (updateStateLock)
            {
                pendingUpdateOnClose = false;
                vsixPath = stagedVsixPath;
                release = stagedRelease;
            }

            if (string.IsNullOrWhiteSpace(vsixPath) || !File.Exists(vsixPath))
            {
                Log("Deferred update on close: staged VSIX not found. Skipping.");
                SetLastUpdateResult("Deferred update skipped because the staged VSIX was not ready.");
                return;
            }

            Log("Deferred update on close: launching VSIXInstaller.");
            if (!LaunchVsixInstaller(vsixPath))
            {
                SetLastUpdateResult("Could not launch VSIXInstaller automatically; opened release page instead.");
                OpenUrl(release?.HtmlUrl ?? ReleasePageUrl);
            }
        }

        private static async Task CheckForUpdatesAsync(AsyncPackage package, CancellationToken token, bool showUpToDate)
        {
            var currentVersion = GetCurrentVersion();
            if (currentVersion == null)
            {
                Log("Update check failed: current version unavailable.");
                SetLastUpdateResult("Update check failed because current version could not be determined.");
                return;
            }

            var release = await GetLatestReleaseAsync(token);
            if (release == null || release.Draft)
            {
                Log("Update check failed: release info unavailable.");
                SetLastUpdateResult("Update check failed because latest release info was unavailable.");
                return;
            }

            var latestVersion = ParseVersion(release.TagName);
            if (latestVersion == null)
            {
                Log("Update check failed: latest version parse failed.");
                SetLastUpdateResult("Update check failed because latest release version could not be parsed.");
                return;
            }

            bool forceUpdateAvailable = false;
#if DEBUG
            forceUpdateAvailable = IsForceUpdateAvailableForDebug();
            if (forceUpdateAvailable)
            {
                Log($"Update check: debug force enabled by {ForceUpdateAvailableEnvironmentVariable}.");
            }
#endif

            if (latestVersion <= currentVersion && !forceUpdateAvailable)
            {
                Log($"Update check: already on latest ({currentVersion}).");
                SetLastUpdateResult($"Up to date ({FormatVersion(currentVersion)}). Latest release is {FormatVersion(latestVersion)}.");
                if (showUpToDate)
                {
                    await package.JoinableTaskFactory.SwitchToMainThreadAsync(token);
                    ShowUpToDatePrompt(package, currentVersion);
                }

                return;
            }

            await package.JoinableTaskFactory.SwitchToMainThreadAsync(token);

            if (forceUpdateAvailable)
            {
                Log($"Update prompt forced for testing: latest {latestVersion}, current {currentVersion}.");
                SetLastUpdateResult($"Debug update test forced. Showing latest release {FormatVersion(latestVersion)} while current version is {FormatVersion(currentVersion)}.");
            }
            else
            {
                Log($"Update available: {latestVersion} (current {currentVersion}).");
                SetLastUpdateResult($"Update available: {FormatVersion(currentVersion)} -> {FormatVersion(latestVersion)}.");
            }

            ShowUpdatePrompt(package, release, latestVersion);
        }

#if DEBUG
        private static bool IsForceUpdateAvailableForDebug()
        {
            string value = Environment.GetEnvironmentVariable(ForceUpdateAvailableEnvironmentVariable);
            return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
        }
#endif

        private static async Task<GitHubRelease> GetLatestReleaseAsync(CancellationToken token)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.UserAgent.ParseAdd(ProductName);
                    client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");

                    using (var response = await client.GetAsync(ReleasesApiUrl, token))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            Log($"Update check HTTP {response.StatusCode}");
                            return null;
                        }

                        string json = await response.Content.ReadAsStringAsync();
                        return JsonConvert.DeserializeObject<GitHubRelease>(json);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Update check fetch failed: {ex.Message}");
                return null;
            }
        }

        internal static Version GetCurrentVersion()
        {
            try
            {
                var assemblyVersion = Assembly.GetExecutingAssembly().GetName().Version;
                if (assemblyVersion != null)
                {
                    Log($"Update check current version (assembly): {assemblyVersion}");
                    return assemblyVersion;
                }
            }
            catch (Exception ex)
            {
                Log($"Update check version parse failed: {ex.Message}");
            }

            Log("Update check current version not found in assembly. Defaulting to 0.0.0.");
            return new Version(0, 0, 0, 0);
        }

        private static Version ParseVersion(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            string value = text.Trim();
            if (value.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                value = value.Substring(1);
            }

            if (Version.TryParse(value, out Version parsed))
            {
                return parsed;
            }

            return null;
        }

        private static void ShowUpdatePrompt(AsyncPackage package, GitHubRelease release, Version latestVersion)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                if (activeInfoBar != null)
                {
                    activeInfoBar.Dismiss();
                    activeInfoBar = null;
                }

                activeInfoBar = new UpdateInfoBar(
                    package,
                    FormatVersion(latestVersion),
                    release?.HtmlUrl,
                    action => HandleInfoBarAction(action));

                if (activeInfoBar.TryShow())
                {
                    Log("Update prompt shown via InfoBar.");
                    StageDownloadInBackground(release);
                    return;
                }

                Log("InfoBar unavailable; opening release page as fallback.");
                activeInfoBar = null;
                SetLastUpdateResult("Update is available, but the InfoBar was unavailable. Opened release page.");
                OpenUrl(release?.HtmlUrl ?? ReleasePageUrl);
            }
            catch (Exception ex)
            {
                Log($"Update prompt failed: {ex.Message}");
            }
        }

        private static void HandleInfoBarAction(string action)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!string.Equals(action, UpdateInfoBar.ActionUpdateOnClose, StringComparison.Ordinal))
            {
                return;
            }

            activeInfoBar = null;
            string vsixPath;
            bool inProgress;
            bool failed;
            GitHubRelease release;
            lock (updateStateLock)
            {
                pendingUpdateOnClose = true;
                vsixPath = stagedVsixPath;
                inProgress = stageDownloadInProgress;
                failed = stageDownloadFailed;
                release = stagedRelease;
            }

            Log("Deferred update on close enabled.");

            if (!string.IsNullOrWhiteSpace(vsixPath) && File.Exists(vsixPath))
            {
                SetLastUpdateResult("Update will install when SSMS closes.");
                return;
            }

            if (inProgress)
            {
                SetLastUpdateResult("Update will install when SSMS closes after the download finishes.");
                return;
            }

            if (failed)
            {
                SetLastUpdateResult("Update download failed; opened release page.");
                OpenUrl(release?.HtmlUrl ?? ReleasePageUrl);
                return;
            }

            SetLastUpdateResult("Update package is not ready; opened release page.");
            OpenUrl(release?.HtmlUrl ?? ReleasePageUrl);
        }

        private static void ShowUpToDatePrompt(AsyncPackage package, Version currentVersion)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                VsShellUtilities.ShowMessageBox(
                    package,
                    $"{DisplayName} is up to date ({FormatVersion(currentVersion)}).",
                    "Information",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
            catch (Exception ex)
            {
                Log($"Up-to-date prompt failed: {ex.Message}");
            }
        }

        private static void StageDownloadInBackground(GitHubRelease release)
        {
            if (release == null)
            {
                return;
            }

            lock (updateStateLock)
            {
                stagedRelease = release;
                stageDownloadFailed = false;
            }

            GitHubAsset asset = GetInstallAsset(release);
            if (asset == null || string.IsNullOrWhiteSpace(asset.DownloadUrl))
            {
                Log("Stage download skipped: no ZIP or VSIX asset found.");
                MarkStageDownloadFailed("Update download failed: no ZIP or VSIX release asset was found.");
                return;
            }

            Task cleanupTask = GetCleanupDownloadedVsixFilesTask();
            var completion = new TaskCompletionSource<bool>();
            lock (updateStateLock)
            {
                stageDownloadInProgress = true;
                stageDownloadTask = completion.Task;
            }

            _ = Task.Run(async () =>
            {
                string downloadedPath = null;
                string extractedVsixPath = null;
                bool staged = false;
                bool verified = false;

                try
                {
                    SetLastUpdateResult("Downloading update package in background.");

                    await cleanupTask;

                    bool hasDigest;
                    string expectedSha256 = GetAssetSha256(asset, out hasDigest);
                    if (hasDigest && string.IsNullOrWhiteSpace(expectedSha256))
                    {
                        Log("Stage download aborted: GitHub release digest is invalid.");
                        MarkStageDownloadFailed("Update download failed: invalid GitHub release digest.");
                        return;
                    }

                    if (!hasDigest)
                    {
                        Log("Stage download: GitHub release digest was not provided; download will not be checksum verified.");
                    }

                    bool isZip = asset.Name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase);
                    string downloadExtension = isZip ? ".zip" : ".vsix";
                    downloadedPath = Path.Combine(Path.GetTempPath(), $"AxialSqlTools-{Guid.NewGuid():N}{downloadExtension}");
                    await DownloadFileAsync(asset.DownloadUrl, downloadedPath, CancellationToken.None);

                    if (!string.IsNullOrWhiteSpace(expectedSha256))
                    {
                        string actualSha256 = ComputeSha256Hex(downloadedPath);
                        if (!string.Equals(actualSha256, expectedSha256, StringComparison.OrdinalIgnoreCase))
                        {
                            Log($"Stage download aborted: checksum mismatch. Expected={expectedSha256}, Actual={actualSha256}");
                            MarkStageDownloadFailed("Update download failed: checksum verification failed.");
                            return;
                        }

                        verified = true;
                    }

                    extractedVsixPath = isZip
                        ? ExtractVsixFromZip(downloadedPath)
                        : downloadedPath;

                    downloadedPath = isZip ? downloadedPath : null;

                    ReplaceStagedVsixPath(extractedVsixPath);
                    staged = true;
                    Log($"Update package staged at: {extractedVsixPath}");

                    bool installPending;
                    lock (updateStateLock)
                    {
                        installPending = pendingUpdateOnClose;
                    }

                    string statusSuffix = verified
                        ? "downloaded and verified"
                        : "downloaded without checksum verification";

                    if (installPending)
                    {
                        SetLastUpdateResult($"Update package {statusSuffix}. It will install when SSMS closes.");
                    }
                    else
                    {
                        SetLastUpdateResult($"Update package {statusSuffix}. Ready to install on close.");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Stage download failed: {ex.Message}");
                    MarkStageDownloadFailed($"Update download failed: {ex.Message}");
                }
                finally
                {
                    lock (updateStateLock)
                    {
                        stageDownloadInProgress = false;
                    }

                    if (!staged && !string.IsNullOrWhiteSpace(extractedVsixPath))
                    {
                        TryDeleteFile(extractedVsixPath);
                    }

                    if (!string.IsNullOrWhiteSpace(downloadedPath))
                    {
                        TryDeleteFile(downloadedPath);
                    }

                    completion.TrySetResult(true);
                }
            });
        }

        private static void MarkStageDownloadFailed(string message)
        {
            bool shouldOpenReleasePage;
            GitHubRelease release;
            lock (updateStateLock)
            {
                stageDownloadFailed = true;
                shouldOpenReleasePage = pendingUpdateOnClose;
                release = stagedRelease;
            }

            SetLastUpdateResult(message);

            if (shouldOpenReleasePage)
            {
                OpenUrl(release?.HtmlUrl ?? ReleasePageUrl);
            }
        }

        private static GitHubAsset GetInstallAsset(GitHubRelease release)
        {
            if (release?.Assets == null)
            {
                return null;
            }

            GitHubAsset firstZip = null;
            GitHubAsset firstVsix = null;

            foreach (var asset in release.Assets)
            {
                if (string.IsNullOrWhiteSpace(asset?.Name))
                {
                    continue;
                }

                if (asset.Name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    if (asset.Name.IndexOf(ProductName, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return asset;
                    }

                    firstZip = firstZip ?? asset;
                }
                else if (asset.Name.EndsWith(".vsix", StringComparison.OrdinalIgnoreCase))
                {
                    firstVsix = firstVsix ?? asset;
                }
            }

            return firstZip ?? firstVsix;
        }

        private static string ExtractVsixFromZip(string zipPath)
        {
            string tempVsixPath = Path.Combine(Path.GetTempPath(), $"AxialSqlTools-{Guid.NewGuid():N}.vsix");

            using (ZipArchive archive = ZipFile.OpenRead(zipPath))
            {
                ZipArchiveEntry selectedEntry = null;
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    string fileName = Path.GetFileName(entry.FullName);
                    if (string.Equals(fileName, ExpectedVsixName, StringComparison.OrdinalIgnoreCase))
                    {
                        selectedEntry = entry;
                        break;
                    }

                    if (selectedEntry == null && fileName.EndsWith(".vsix", StringComparison.OrdinalIgnoreCase))
                    {
                        selectedEntry = entry;
                    }
                }

                if (selectedEntry == null)
                {
                    throw new InvalidOperationException("The release ZIP did not contain a VSIX file.");
                }

                using (Stream source = selectedEntry.Open())
                using (Stream target = File.Create(tempVsixPath))
                {
                    source.CopyTo(target);
                }
            }

            return tempVsixPath;
        }

        private static string GetAssetSha256(GitHubAsset asset, out bool hasDigest)
        {
            string digest = asset?.Digest;
            if (string.IsNullOrWhiteSpace(digest))
            {
                hasDigest = false;
                return null;
            }

            hasDigest = true;
            string trimmed = digest.Trim();
            if (!trimmed.StartsWith(Sha256DigestPrefix, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            string value = trimmed.Substring(Sha256DigestPrefix.Length).Trim();
            if (Regex.IsMatch(value, "^[A-Fa-f0-9]{64}$"))
            {
                return value.ToLowerInvariant();
            }

            return null;
        }

        private static Task ScheduleCleanupDownloadedVsixFiles()
        {
            Task cleanupTask = Task.Run((Action)CleanupDownloadedVsixFiles);
            lock (updateStateLock)
            {
                cleanupDownloadedVsixFilesTask = cleanupTask;
            }

            return cleanupTask;
        }

        private static Task GetCleanupDownloadedVsixFilesTask()
        {
            lock (updateStateLock)
            {
                return cleanupDownloadedVsixFilesTask;
            }
        }

        private static void CleanupDownloadedVsixFiles()
        {
            string activePath;
            lock (updateStateLock)
            {
                if (stageDownloadInProgress)
                {
                    Log("Cleanup staged update files skipped because a download is in progress.");
                    return;
                }

                activePath = stagedVsixPath;
                stagedVsixPath = null;
                pendingUpdateOnClose = false;
                stageDownloadFailed = false;
                stagedRelease = null;
            }

            try
            {
                string tempDirectory = Path.GetTempPath();
                foreach (string filePath in Directory.GetFiles(tempDirectory, StagedVsixFilePattern))
                {
                    TryDeleteFile(filePath);
                }

                foreach (string filePath in Directory.GetFiles(tempDirectory, StagedZipFilePattern))
                {
                    TryDeleteFile(filePath);
                }

                if (!string.IsNullOrWhiteSpace(activePath) && File.Exists(activePath))
                {
                    TryDeleteFile(activePath);
                }
            }
            catch (Exception ex)
            {
                Log($"Cleanup staged update files failed: {ex.Message}");
            }
        }

        private static void ReplaceStagedVsixPath(string tempPath)
        {
            if (string.IsNullOrWhiteSpace(tempPath))
            {
                return;
            }

            lock (updateStateLock)
            {
                string previousPath = stagedVsixPath;
                stagedVsixPath = tempPath;

                if (!string.IsNullOrWhiteSpace(previousPath) &&
                    !string.Equals(previousPath, tempPath, StringComparison.OrdinalIgnoreCase))
                {
                    TryDeleteFile(previousPath);
                }
            }
        }

        private static void TryDeleteFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    Log($"Deleted staged update file: {filePath}");
                }
            }
            catch (Exception ex)
            {
                Log($"Delete staged update file failed for '{filePath}': {ex.Message}");
            }
        }

        private static string ComputeSha256Hex(string filePath)
        {
            using (var stream = File.OpenRead(filePath))
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(stream);
                var builder = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash)
                {
                    builder.Append(b.ToString("x2"));
                }

                return builder.ToString();
            }
        }

        private static async Task DownloadFileAsync(string url, string path, CancellationToken token)
        {
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.ParseAdd(ProductName);
                using (var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, token))
                {
                    response.EnsureSuccessStatusCode();
                    using (Stream source = await response.Content.ReadAsStreamAsync())
                    using (Stream target = File.Create(path))
                    {
                        await source.CopyToAsync(target);
                    }
                }
            }
        }

        private static bool LaunchVsixInstaller(string vsixPath)
        {
            try
            {
                string installerPath = GetVsixInstallerPath();
                if (string.IsNullOrWhiteSpace(installerPath) || !File.Exists(installerPath))
                {
                    Log("VSIXInstaller.exe not found; falling back to browser.");
                    return false;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = installerPath,
                    Arguments = $"\"{vsixPath}\"",
                    UseShellExecute = true
                });

                Log($"Launched VSIXInstaller: {installerPath} \"{vsixPath}\"");
                SetLastUpdateResult("VSIXInstaller launched.");
                return true;
            }
            catch (Exception ex)
            {
                Log($"Launch VSIXInstaller failed: {ex.Message}");
                return false;
            }
        }

        private static string GetVsixInstallerPath()
        {
            try
            {
                string exePath = Process.GetCurrentProcess()?.MainModule?.FileName;
                if (!string.IsNullOrWhiteSpace(exePath))
                {
                    string exeDir = Path.GetDirectoryName(exePath);
                    string candidate = Path.Combine(exeDir ?? string.Empty, "VSIXInstaller.exe");
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Resolve VSIXInstaller from process failed: {ex.Message}");
            }

            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string candidate = Path.Combine(baseDir ?? string.Empty, "VSIXInstaller.exe");
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            catch (Exception ex)
            {
                Log($"Resolve VSIXInstaller from AppDomain failed: {ex.Message}");
            }

            return null;
        }

        internal static string FormatVersion(Version version)
        {
            if (version == null)
            {
                return "0.0.0";
            }

            if (version.Revision > 0)
            {
                return version.ToString(4);
            }

            if (version.Build > 0)
            {
                return version.ToString(3);
            }

            return version.ToString(2);
        }

        private static void OpenUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Log($"Open URL failed: {ex.Message}");
            }
        }

        private static void SetLastUpdateResult(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            string stamped = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            lock (diagnosticsLock)
            {
                lastUpdateResult = stamped;
            }

            try
            {
                LastUpdateResultChanged?.Invoke();
            }
            catch
            {
            }
        }

        private sealed class GitHubRelease
        {
            [JsonProperty("tag_name")]
            public string TagName { get; set; }

            [JsonProperty("html_url")]
            public string HtmlUrl { get; set; }

            [JsonProperty("draft")]
            public bool Draft { get; set; }

            [JsonProperty("prerelease")]
            public bool Prerelease { get; set; }

            [JsonProperty("assets")]
            public List<GitHubAsset> Assets { get; set; }
        }

        private sealed class GitHubAsset
        {
            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("digest")]
            public string Digest { get; set; }

            [JsonProperty("browser_download_url")]
            public string DownloadUrl { get; set; }
        }
    }
}