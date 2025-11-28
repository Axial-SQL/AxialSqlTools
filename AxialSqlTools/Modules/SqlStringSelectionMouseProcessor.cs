using System;
using System.ComponentModel.Composition;
using System.Windows;
using System.Windows.Input;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;
using Microsoft.VisualStudio.Utilities;

namespace AxialSqlTools
{
    [Export(typeof(IMouseProcessorProvider))]
    [Name("sql-string-double-click-selector")]
    [ContentType("text")]
    [ContentType("code")]
    [Order(Before = PredefinedMouseProcessorNames.WordSelection)]
    [TextViewRole(PredefinedTextViewRoles.Editable)]
    public class SqlStringSelectionMouseProcessorProvider : IMouseProcessorProvider
    {
        public IMouseProcessor GetAssociatedProcessor(IWpfTextView wpfTextView)
        {
            return new SqlStringSelectionMouseProcessor(wpfTextView);
        }
    }

    public class SqlStringSelectionMouseProcessor : MouseProcessorBase
    {
        private readonly IWpfTextView view;

        public SqlStringSelectionMouseProcessor(IWpfTextView view)
        {
            this.view = view ?? throw new ArgumentNullException(nameof(view));
        }

        public override void PreprocessMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left)
            {
                return;
            }

            if (e.ClickCount != 2)
            {
                return;
            }

            if (Keyboard.Modifiers != ModifierKeys.None)
            {
                return;
            }

            if (SelectQuotedString(e))
            {
                e.Handled = true;
            }
        }

        private bool SelectQuotedString(MouseButtonEventArgs e)
        {
            SnapshotPoint? snapshotPoint = GetSnapshotAtCursor(e);

            if (snapshotPoint.HasValue && TryGetQuotedStringSpan(snapshotPoint.Value, out SnapshotSpan span))
            {
                view.Selection.Select(span, isReversed: false);
                view.Caret.MoveTo(span.End);
                return true;
            }

            return false;
        }

        private SnapshotPoint? GetSnapshotAtCursor(MouseButtonEventArgs e)
        {
            if (view.TextViewLines == null || view.TextViewLines.IsEmpty)
            {
                return null;
            }

            Point cursorPosition = GetPositionInViewport(e);
            ITextViewLine textViewLine = view.TextViewLines.GetTextViewLineContainingYCoordinate(cursorPosition.Y);

            if (textViewLine != null)
            {
                return textViewLine.GetBufferPositionFromXCoordinate(cursorPosition.X, true);
            }

            return null;
        }

        private Point GetPositionInViewport(MouseButtonEventArgs e)
        {
            Point relativePosition = e.GetPosition(view.VisualElement);
            return new Point(relativePosition.X + view.ViewportLeft, relativePosition.Y + view.ViewportTop);
        }

        private bool TryGetQuotedStringSpan(SnapshotPoint point, out SnapshotSpan span)
        {
            string text = point.Snapshot.GetText();
            int position = point.Position;

            bool insideString = false;
            int stringStart = -1;

            for (int index = 0; index < text.Length; index++)
            {
                char current = text[index];

                if (current == '\'')
                {
                    if (insideString)
                    {
                        if (index + 1 < text.Length && text[index + 1] == '\'')
                        {
                            index++;
                            continue;
                        }

                        int stringEnd = index;
                        if (position >= stringStart && position <= stringEnd)
                        {
                            span = new SnapshotSpan(point.Snapshot, stringStart, stringEnd - stringStart + 1);
                            return true;
                        }

                        insideString = false;
                    }
                    else
                    {
                        stringStart = index;
                        if (index > 0 && (text[index - 1] == 'N' || text[index - 1] == 'n'))
                        {
                            stringStart--;
                        }

                        insideString = true;
                    }
                }
            }

            if (insideString && position >= stringStart)
            {
                span = new SnapshotSpan(point.Snapshot, stringStart, text.Length - stringStart);
                return true;
            }

            span = default;
            return false;
        }
    }
}
