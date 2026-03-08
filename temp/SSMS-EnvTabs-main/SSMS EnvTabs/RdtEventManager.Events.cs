using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager
    {
        // --- RDT events ---
        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            ThreadHelper.ThrowIfNotOnUIThread(); // Ensure UI thread

            // Try to hook into this event since OnAfterDocumentWindowShow is unreliable in SSMS
            string moniker = TryGetMonikerFromCookie(docCookie);

            if (!IsSqlDocumentMoniker(moniker))
            {
                return VSConstants.S_OK;
            }

            IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
            if (frame != null)
            {
                bool done = HandlePotentialChange(docCookie, frame, reason: "FirstDocumentLock");
                if (!done)
                {
                    ScheduleRenameRetry(docCookie, "FirstDocumentLock");
                }
            }

            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            if (pFrame != null)
            {
                if (TryGetMonikerFromFrame(pFrame, out string moniker) && IsSqlDocumentMoniker(moniker))
                {
                    bool done = HandlePotentialChange(docCookie, pFrame, reason: "DocumentWindowShow");
                    if (!done)
                    {
                        ScheduleRenameRetry(docCookie, "DocumentWindowShow");
                    }
                }
            }

            return VSConstants.S_OK;
        }

        public int OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld, string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            string moniker = pszMkDocumentNew;
            if (string.IsNullOrWhiteSpace(moniker))
            {
                moniker = pszMkDocumentOld;
            }
            
            if (string.IsNullOrWhiteSpace(moniker))
            {
                // Fallback: If moniker args are empty (common for attribute-only changes like Dirty/Reload),
                // fetch it from the RDT using the cookie.
                moniker = TryGetMonikerFromCookie(docCookie);
            }

            if (IsSqlDocumentMoniker(moniker))
            {
                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame != null)
                {
                    bool done = HandlePotentialChange(docCookie, frame, reason: "AttributeChangeEx");

                    // Force a retry if it's an AttributeChange, often the DocView is not yet updated with new connection info
                    if (!done)
                    {
                        ScheduleRenameRetry(docCookie, "AttributeChangeEx");
                    }
                }
            }

            return VSConstants.S_OK;
        }

        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (dwReadLocksRemaining == 0 && dwEditLocksRemaining == 0)
            {
                TabRenamer.ForgetCookie(docCookie);
                renameRetryCounts.Remove(docCookie);
                lastConnectionByCookie.Remove(docCookie);
                lastCaptionByCookie.Remove(docCookie);

                UpdateColorOnly("LastDocumentUnlock");
            }

            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterSave(uint docCookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string moniker = TryGetMonikerFromCookie(docCookie);
            if (IsSqlDocumentMoniker(moniker))
            {
                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame != null)
                {
                    bool done = HandlePotentialChange(docCookie, frame, reason: "AfterSave");
                    if (!done)
                    {
                        ScheduleRenameRetry(docCookie, "AfterSave");
                    }
                }
            }
            UpdateColorOnly("AfterSave");
            return VSConstants.S_OK;
        }

        public int OnBeforeSave(uint docCookie) => VSConstants.S_OK;

        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string moniker = TryGetMonikerFromCookie(docCookie);
            if (IsSqlDocumentMoniker(moniker))
            {
                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame != null)
                {
                    bool done = HandlePotentialChange(docCookie, frame, reason: "AttributeChange");

                    // Force a retry if it's an AttributeChange, because the caption is likely stale.
                    if (!done)
                    {
                        ScheduleRenameRetry(docCookie, "AttributeChange");
                    }
                    else
                    {
                        ScheduleRenameRetry(docCookie, "AttributeChange:Force");
                    }
                }
            }

            return VSConstants.S_OK;
        }

        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame) => VSConstants.S_OK;
    }
}
