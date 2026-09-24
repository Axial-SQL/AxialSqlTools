using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using System;
using System.Windows;
using System.Windows.Media;
using System.Xml;

namespace AxialSqlTools
{
    public static class SqlEditorSupport
    {
        // AvalonEdit.Text is not a dependency property. This one-way binding bridge keeps
        // the read-only history preview in sync when the selected record changes/clears.
        public static readonly DependencyProperty PreviewTextProperty = DependencyProperty.RegisterAttached(
            "PreviewText", typeof(string), typeof(SqlEditorSupport), new PropertyMetadata(string.Empty, PreviewTextChanged));

        public static string GetPreviewText(DependencyObject target) => (string)target.GetValue(PreviewTextProperty);
        public static void SetPreviewText(DependencyObject target, string value) => target.SetValue(PreviewTextProperty, value);

        private static void PreviewTextChanged(DependencyObject target, DependencyPropertyChangedEventArgs args)
        {
            if (!(target is TextEditor editor)) return;
            string text = args.NewValue as string ?? string.Empty;
            if (editor.Text == text) return;
            editor.Text = text;
            editor.Document.UndoStack.ClearAll();
            editor.ScrollToHome();
        }

        internal static void ApplyTheme(TextEditor editor, FrameworkElement scope)
        {
            if (editor == null) return;
            var background = VsThemeBrushResolver.GetBrushColor(editor.Background, SystemColors.WindowColor);
            var foreground = VsThemeBrushResolver.GetBrushColor(editor.Foreground, SystemColors.WindowTextColor);
            bool isLightTheme = VsThemeBrushResolver.GetRelativeLuminance(background) > 0.6;
            // Load a fresh definition so theme changes never mutate a shared definition or
            // progressively wash out its colors. Existing text, selection and undo stay intact.
            using (var stream = typeof(SqlEditorSupport).Assembly.GetManifestResourceStream("AxialSqlTools.QuickSearch.sql.xshd"))
            {
                if (stream == null) throw new InvalidOperationException("The embedded SQL highlighting definition is missing.");
                using (var reader = XmlReader.Create(stream))
                {
                    var highlighting = HighlightingLoader.Load(reader, HighlightingManager.Instance);
                    foreach (var color in highlighting.NamedHighlightingColors)
                    {
                        var original = color.Foreground?.GetColor(null) ?? foreground;
                        // Keep SSMS's red strings and magenta functions on light backgrounds.
                        // A contrast correction there would change their standard hues.
                        color.Foreground = new SimpleHighlightingBrush(SystemParameters.HighContrast ? foreground
                            : isLightTheme ? original
                            : VsThemeBrushResolver.EnsureTextContrast(original, background, foreground));
                    }
                    editor.SyntaxHighlighting = highlighting;
                }
            }

            var selection = scope.TryFindResource("AxialThemeGridSelectionBrush") as Brush ?? SystemColors.HighlightBrush;
            editor.TextArea.SelectionBrush = selection;
            editor.TextArea.SelectionForeground = scope.TryFindResource("AxialThemeGridSelectionTextBrush") as Brush ?? SystemColors.HighlightTextBrush;
            editor.TextArea.SelectionBorder = new Pen(selection, 1);
            editor.TextArea.SelectionCornerRadius = 0;
            var lineNumbers = editor.SyntaxHighlighting.GetNamedColor("Operator").Foreground;
            editor.LineNumbersForeground = new SolidColorBrush(lineNumbers.GetColor(null) ?? foreground);
            editor.Options.EnableHyperlinks = false;
            editor.Options.EnableEmailHyperlinks = false;
        }
    }
}
