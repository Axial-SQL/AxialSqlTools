using Microsoft.VisualStudio.Shell;
using System.Runtime.InteropServices;

namespace AxialSqlTools
{
    [Guid("3f0450e0-5ef4-4d6b-bf0f-cfdd95d7c003")]
    public class QuickSearchWindow : ToolWindowPane
    {
        public QuickSearchWindow() : base(null)
        {
            this.Caption = "Quick Search";
            this.Content = new QuickSearchWindowControl();
        }
    }
}
