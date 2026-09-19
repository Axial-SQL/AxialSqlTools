using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using NLog;
using System;
using System.Windows.Input;

namespace AxialSqlTools.EditorSelection
{
    internal sealed class QuotedStringMouseProcessor : MouseProcessorBase
    {
        // Bound synchronous lexing on the editor thread; larger scripts keep SSMS selection.
        private const int MaximumSnapshotLength = 1024 * 1024;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly IWpfTextView _view;
        private ITextSnapshot _cachedSnapshot;
        private SqlStringSelectionResolver _resolver;

        internal QuotedStringMouseProcessor(IWpfTextView view)
        {
            _view = view;
            _view.Closed += OnViewClosed;
            Logger.Debug("Quoted-string selection attached to editor content type {0}.",
                view.TextBuffer.ContentType.TypeName);
        }

        public override void PreprocessMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            if (e.Handled || e.ChangedButton != MouseButton.Left || e.ClickCount != 2
                || Keyboard.Modifiers != ModifierKeys.None || _view.IsClosed || _view.InLayout
                || !SettingsManager.GetSelectQuotedStringOnDoubleClick())
                return;

            try
            {
                var position = e.GetPosition(_view.VisualElement);
                position.X += _view.ViewportLeft;
                position.Y += _view.ViewportTop;
                var line = _view.TextViewLines.GetTextViewLineContainingYCoordinate(position.Y);
                if (line == null || position.X < line.TextLeft || position.X >= line.TextRight)
                    return;

                var point = line.GetBufferPositionFromXCoordinate(position.X, true);
                if (!point.HasValue) return;
                var snapshot = point.Value.Snapshot;
                if (!ReferenceEquals(snapshot, _view.TextSnapshot)) return;

                if (!ReferenceEquals(snapshot, _cachedSnapshot))
                {
                    _cachedSnapshot = snapshot;
                    _resolver = null;
                    if (snapshot.Length > MaximumSnapshotLength) return;
                    _resolver = SqlStringSelectionResolver.Create(snapshot.GetText());
                }

                if (_resolver == null || !_resolver.TryGetContentSpan(point.Value.Position, out var content))
                    return;

                _view.Selection.Select(new SnapshotSpan(snapshot, content.Start, content.Length), false);
                _view.Caret.MoveTo(_view.Selection.ActivePoint);
                e.Handled = true;
            }
            catch (Exception ex)
            {
                // Never log query text and never suppress normal selection on a failed attempt.
                Logger.Warn("Quoted-string selection failed ({0}). Normal editor selection will be used.",
                    ex.GetType().Name);
            }
        }

        private void OnViewClosed(object sender, EventArgs e)
        {
            _view.Closed -= OnViewClosed;
            _cachedSnapshot = null;
            _resolver = null;
        }
    }
}
