using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace AxialSqlTools.TabColoring
{
    internal sealed class TabColoringRdtManager : IVsRunningDocTableEvents3, IDisposable
    {
        private readonly IVsRunningDocumentTable _rdt;
        private uint _rdtEventsCookie;

        private TabColoringRdtManager(IVsRunningDocumentTable rdt)
        {
            _rdt = rdt ?? throw new ArgumentNullException(nameof(rdt));
        }

        public static async Task<TabColoringRdtManager> CreateAndStartAsync(AsyncPackage package, CancellationToken cancellationToken)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var rdt = await package.GetServiceAsync(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            if (rdt == null)
            {
                return null;
            }

            var manager = new TabColoringRdtManager(rdt);
            manager.Start();
            return manager;
        }

        private void Start()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _rdt.AdviseRunningDocTableEvents(this, out _rdtEventsCookie);
        }

        public void Dispose()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_rdtEventsCookie != 0)
            {
                _rdt.UnadviseRunningDocTableEvents(_rdtEventsCookie);
                _rdtEventsCookie = 0;
            }
        }

        public int OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld, string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew)
        {
            // Infrastructure hook only.
            // Next implementation step: detect active SQL query window + active connection,
            // execute query rules against that connection and map first match to SSMS color index.
            _ = TabColoringConfigLoader.LoadOrNull();
            return VSConstants.S_OK;
        }

        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs) => VSConstants.S_OK;
        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame) => VSConstants.S_OK;
        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining) => VSConstants.S_OK;
        public int OnAfterSave(uint docCookie) => VSConstants.S_OK;
        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame) => VSConstants.S_OK;
        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining) => VSConstants.S_OK;
        public int OnBeforeSave(uint docCookie) => VSConstants.S_OK;
    }
}
