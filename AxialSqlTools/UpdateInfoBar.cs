using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics;

namespace AxialSqlTools
{
    internal sealed class UpdateInfoBar : IVsInfoBarUIEvents
    {
        internal const string ActionUpdateOnClose = "update_on_close";
        private const string ActionReleaseNotes = "release_notes";

        private readonly AsyncPackage package;
        private readonly string latestVersion;
        private readonly string releaseUrl;
        private readonly Action<string> onActionRequested;

        private IVsInfoBarUIElement currentElement;
        private uint adviseCookie;

        internal UpdateInfoBar(AsyncPackage package, string latestVersion, string releaseUrl, Action<string> onActionRequested)
        {
            this.package = package;
            this.latestVersion = latestVersion;
            this.releaseUrl = releaseUrl;
            this.onActionRequested = onActionRequested;
        }

        internal bool TryShow()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var shell = package.GetService<SVsShell, IVsShell>();
                if (shell == null)
                {
                    UpdateChecker.Log("UpdateInfoBar: IVsShell service unavailable.");
                    return false;
                }

                if (ErrorHandler.Failed(shell.GetProperty((int)__VSSPROPID7.VSSPROPID_MainWindowInfoBarHost, out object hostObj))
                    || !(hostObj is IVsInfoBarHost host))
                {
                    UpdateChecker.Log("UpdateInfoBar: MainWindowInfoBarHost not available.");
                    return false;
                }

                var factory = package.GetService<SVsInfoBarUIFactory, IVsInfoBarUIFactory>();
                if (factory == null)
                {
                    UpdateChecker.Log("UpdateInfoBar: IVsInfoBarUIFactory service unavailable.");
                    return false;
                }

                var model = new InfoBarModel(
                    new IVsInfoBarTextSpan[]
                    {
                        new InfoBarTextSpan($"Axial SQL Tools v{latestVersion} update available.  ")
                    },
                    new IVsInfoBarActionItem[]
                    {
                        new InfoBarButton("Update on Close", ActionUpdateOnClose),
                        new InfoBarHyperlink("Release Notes", ActionReleaseNotes)
                    },
                    KnownMonikers.StatusInformation,
                    isCloseButtonVisible: true);

                currentElement = factory.CreateInfoBar(model);
                if (currentElement == null)
                {
                    UpdateChecker.Log("UpdateInfoBar: CreateInfoBar returned null.");
                    return false;
                }

                currentElement.Advise(this, out adviseCookie);
                host.AddInfoBar(currentElement);
                UpdateChecker.Log("UpdateInfoBar: Shown successfully.");
                return true;
            }
            catch (Exception ex)
            {
                UpdateChecker.Log($"UpdateInfoBar: Failed to show: {ex.Message}");
                return false;
            }
        }

        internal void Dismiss()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (currentElement != null)
                {
                    currentElement.Unadvise(adviseCookie);
                    currentElement.Close();
                    currentElement = null;
                }
            }
            catch (Exception ex)
            {
                UpdateChecker.Log($"UpdateInfoBar: Dismiss failed: {ex.Message}");
                currentElement = null;
            }
        }

        public void OnActionItemClicked(IVsInfoBarUIElement infoBarUIElement, IVsInfoBarActionItem actionItem)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string context = actionItem.ActionContext as string;
            switch (context)
            {
                case ActionUpdateOnClose:
                    UpdateChecker.Log("UpdateInfoBar: User clicked Update on Close.");
                    Dismiss();
                    onActionRequested?.Invoke(ActionUpdateOnClose);
                    break;

                case ActionReleaseNotes:
                    UpdateChecker.Log("UpdateInfoBar: User clicked Release Notes.");
                    OpenReleaseNotes();
                    break;
            }
        }

        public void OnClosed(IVsInfoBarUIElement infoBarUIElement)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (currentElement != null)
                {
                    currentElement.Unadvise(adviseCookie);
                    currentElement = null;
                }
            }
            catch
            {
                currentElement = null;
            }
        }

        private void OpenReleaseNotes()
        {
            if (string.IsNullOrWhiteSpace(releaseUrl))
            {
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = releaseUrl,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                UpdateChecker.Log($"UpdateInfoBar: Failed to open release notes: {ex.Message}");
            }
        }
    }
}