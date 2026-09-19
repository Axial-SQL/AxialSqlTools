using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;
using System.ComponentModel.Composition;

namespace AxialSqlTools.EditorSelection
{
    [Export(typeof(IMouseProcessorProvider))]
    [Name("AxialSqlTools.QuotedStringSelection")]
    [ContentType("SQL")]
    [ContentType("SQL Server Tools")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    internal sealed class QuotedStringMouseProcessorProvider : IMouseProcessorProvider
    {
        public IMouseProcessor GetAssociatedProcessor(IWpfTextView wpfTextView)
        {
            // Register even while disabled so changing the setting takes effect in open views.
            return wpfTextView.Properties.GetOrCreateSingletonProperty(
                () => new QuotedStringMouseProcessor(wpfTextView));
        }
    }
}
