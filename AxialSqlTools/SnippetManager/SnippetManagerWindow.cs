using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace AxialSqlTools
{
    [Guid("a7b2c3d4-e5f6-4789-abcd-ef0123456789")]
    public class SnippetManagerWindow : ToolWindowPane
    {
        public SnippetManagerWindow() : base(null)
        {
            this.Caption = "Snippet Manager";
            this.Content = new SnippetManagerWindowControl();
        }
    }
}
