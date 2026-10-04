using Microsoft.VisualStudio.Shell;
using System.Runtime.InteropServices;

namespace AxialSqlTools
{
    [Guid("ecec194e-699a-4e17-9a13-b49aa44d44a0")]
    public sealed class DistributedAgDashboardWindow : ToolWindowPane
    {
        public DistributedAgDashboardWindow() : base(null)
        {
            Caption = "AG Health Dashboard";
            Content = new DistributedAgDashboardControl();
        }

        public void Initialize(ScriptFactoryAccess.ConnectionInfo connection, string availabilityGroupName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ((DistributedAgDashboardControl)Content).Initialize(connection, availabilityGroupName);
            Caption = "AG Health Dashboard | " + (availabilityGroupName ?? connection.ServerName);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) (Content as DistributedAgDashboardControl)?.Dispose();
            base.Dispose(disposing);
        }
    }
}
