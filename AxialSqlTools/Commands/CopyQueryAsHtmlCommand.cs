using System;
using System.ComponentModel.Design;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using static AxialSqlTools.AxialSqlToolsPackage;

namespace AxialSqlTools
{
    internal sealed class CopyQueryAsHtmlCommand
    {
        public const int CommandId = 4160;
        public static readonly Guid CommandSet = new Guid("45457e02-6dec-4a4d-ab22-c9ee126d23c5");

        private readonly AsyncPackage _package;

        private CopyQueryAsHtmlCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandId = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandId);
            commandService.AddCommand(menuItem);
        }

        public static CopyQueryAsHtmlCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new CopyQueryAsHtmlCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _ = ThreadHelper.JoinableTaskFactory.RunAsync(async delegate
            {
                await ExecuteAsync();
            });
        }

        private async Task ExecuteAsync()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (dte?.ActiveDocument == null)
            {
                return;
            }

            var selection = dte.ActiveDocument.Selection as TextSelection;
            if (selection == null)
            {
                return;
            }

            var originalLine = selection.ActivePoint.Line;
            var originalOffset = selection.ActivePoint.LineCharOffset;
            var hadSelection = !selection.IsEmpty;

            if (!hadSelection)
            {
                selection.SelectAll();
            }

            if (selection.IsEmpty)
            {
                return;
            }

            // DTE's TextSelection.Copy() does not always populate the clipboard with the
            // formatted payload (RTF/HTML) that OWA expects. Mimic the user pressing
            // Ctrl+C by invoking the Edit.Copy command, which reliably fills the
            // clipboard with all supported formats for the current editor.
            dte.ExecuteCommand("Edit.Copy");

            // Give the clipboard a brief moment to populate before we read back
            // the formats needed for HTML export.
            await Task.Delay(50);

            var textContent = selection.Text;
            if (!hadSelection)
            {
                selection.MoveToLineAndOffset(originalLine, originalOffset);
            }

            var dataObject = Clipboard.GetDataObject();
            var clipboardHtml = dataObject?.GetData(DataFormats.Html) as string;
            var rtfContent = dataObject?.GetData(DataFormats.Rtf) as string;

            var htmlFragment = !string.IsNullOrWhiteSpace(clipboardHtml)
                ? NormalizeClipboardHtml(clipboardHtml)
                : !string.IsNullOrWhiteSpace(rtfContent)
                    ? ConvertRtfToHtmlClipboardFragment(rtfContent)
                    : BuildClipboardHtml(WrapPlainText(textContent));

            var newDataObject = new DataObject();
            newDataObject.SetData(DataFormats.UnicodeText, textContent);
            newDataObject.SetData(DataFormats.Text, textContent);

            if (!string.IsNullOrWhiteSpace(rtfContent))
            {
                newDataObject.SetData(DataFormats.Rtf, rtfContent);
            }

            if (!string.IsNullOrWhiteSpace(htmlFragment))
            {
                newDataObject.SetData(DataFormats.Html, htmlFragment);
            }

            Clipboard.SetDataObject(newDataObject, true);
        }

        private static string ConvertRtfToHtmlClipboardFragment(string rtfContent)
        {
            using (var richTextBox = new RichTextBox())
            {
                richTextBox.Rtf = rtfContent;

                var sb = new StringBuilder();
                sb.Append("<pre style=\"white-space:pre-wrap;\">");

                var text = richTextBox.Text;
                var index = 0;
                while (index < text.Length)
                {
                    richTextBox.Select(index, 1);

                    var currentColor = richTextBox.SelectionColor;
                    var currentFont = richTextBox.SelectionFont;

                    var runStart = index;
                    index++;

                    while (index < text.Length)
                    {
                        richTextBox.Select(index, 1);
                        var nextColor = richTextBox.SelectionColor;
                        var nextFont = richTextBox.SelectionFont;
                        if (!ColorsEqual(currentColor, nextColor) || !FontsEqual(currentFont, nextFont))
                        {
                            break;
                        }
                        index++;
                    }

                    var runLength = index - runStart;
                    var segment = text.Substring(runStart, runLength);
                    sb.Append(WrapSegment(segment, currentColor, currentFont));
                }

                sb.Append("</pre>");
                return BuildClipboardHtml(sb.ToString());
            }
        }

        private static string WrapPlainText(string textContent)
        {
            var encodedText = System.Web.HttpUtility.HtmlEncode(textContent ?? string.Empty);
            return $"<pre style=\"white-space:pre-wrap;\">{encodedText}</pre>";
        }

        private static string WrapSegment(string text, System.Drawing.Color color, System.Drawing.Font font)
        {
            var styleBuilder = new StringBuilder();

            if (font != null)
            {
                styleBuilder.Append($"font-family: '{font.FontFamily.Name}', monospace;");
                styleBuilder.Append($"font-size: {font.SizeInPoints.ToString(System.Globalization.CultureInfo.InvariantCulture)}pt;");
                if (font.Bold)
                {
                    styleBuilder.Append("font-weight:bold;");
                }
                if (font.Italic)
                {
                    styleBuilder.Append("font-style:italic;");
                }
                if (font.Underline)
                {
                    styleBuilder.Append("text-decoration:underline;");
                }
            }
            else
            {
                styleBuilder.Append("font-family: Consolas, monospace;");
            }

            var adjustedColor = AdjustColorForClipboard(color);
            styleBuilder.Append($"color: rgb({adjustedColor.R}, {adjustedColor.G}, {adjustedColor.B});");

            var encoded = System.Web.HttpUtility.HtmlEncode(text);
            return $"<span style=\"{styleBuilder}\">{encoded}</span>";
        }

        private static bool ColorsEqual(System.Drawing.Color first, System.Drawing.Color second)
        {
            return first.R == second.R && first.G == second.G && first.B == second.B;
        }

        private static bool FontsEqual(System.Drawing.Font first, System.Drawing.Font second)
        {
            if (first == null && second == null)
            {
                return true;
            }

            if (first == null || second == null)
            {
                return false;
            }

            return string.Equals(first.FontFamily.Name, second.FontFamily.Name, StringComparison.OrdinalIgnoreCase)
                && Math.Abs(first.SizeInPoints - second.SizeInPoints) < 0.1
                && first.Style == second.Style;
        }

        private static string NormalizeClipboardHtml(string htmlPayload)
        {
            const string StartMarker = "<!--StartFragment-->";
            const string EndMarker = "<!--EndFragment-->";

            if (string.IsNullOrWhiteSpace(htmlPayload))
            {
                return BuildClipboardHtml(string.Empty);
            }

            var startIndex = htmlPayload.IndexOf(StartMarker, StringComparison.OrdinalIgnoreCase);
            var endIndex = htmlPayload.IndexOf(EndMarker, StringComparison.OrdinalIgnoreCase);

            if (startIndex < 0 || endIndex <= startIndex)
            {
                return BuildClipboardHtml(WrapPlainText(htmlPayload));
            }

            var fragmentStart = startIndex + StartMarker.Length;
            var fragmentLength = endIndex - fragmentStart;
            var fragment = htmlPayload.Substring(fragmentStart, fragmentLength);

            var adjustedFragment = ReplaceLimeGreen(fragment);
            return BuildClipboardHtml(adjustedFragment);
        }

        private static string BuildClipboardHtml(string htmlBody)
        {
            const string HeaderTemplate = "Version:0.9\\r\\nStartHTML:{0:0000000000}\\r\\nEndHTML:{1:0000000000}\\r\\nStartFragment:{2:0000000000}\\r\\nEndFragment:{3:0000000000}\\r\\n";
            var encoding = Encoding.UTF8;
            var prefix = "<html><body>";
            var fragmentStart = "<!--StartFragment-->";
            var fragmentEnd = "<!--EndFragment-->";
            var suffix = "</body></html>";
            var safeHtmlBody = htmlBody ?? string.Empty;

            var fullHtmlBuilder = new StringBuilder();
            fullHtmlBuilder.Append(prefix);
            fullHtmlBuilder.Append(fragmentStart);
            fullHtmlBuilder.Append(safeHtmlBody);
            fullHtmlBuilder.Append(fragmentEnd);
            fullHtmlBuilder.Append(suffix);
            var fullHtml = fullHtmlBuilder.ToString();

            var headerPlaceholder = string.Format(HeaderTemplate, 0, 0, 0, 0);
            var startHtml = encoding.GetByteCount(headerPlaceholder);
            var startFragment = startHtml + encoding.GetByteCount(prefix) + encoding.GetByteCount(fragmentStart);
            var endFragment = startFragment + encoding.GetByteCount(safeHtmlBody);
            var endHtml = startHtml + encoding.GetByteCount(fullHtml);

            var header = string.Format(HeaderTemplate, startHtml, endHtml, startFragment, endFragment);
            var payload = header + fullHtml;

            // Defensive recalculation: if any offset is out of sync with the final payload
            // (e.g., due to unexpected encoding differences), rebuild the header using the
            // measured byte lengths so the clipboard HTML fragment markers remain valid.
            var encodedPayload = encoding.GetBytes(payload);
            var measuredStartHtml = encoding.GetByteCount(header);
            var measuredStartFragment = measuredStartHtml + encoding.GetByteCount(prefix) + encoding.GetByteCount(fragmentStart);
            var measuredEndFragment = measuredStartFragment + encoding.GetByteCount(safeHtmlBody);
            var measuredEndHtml = encodedPayload.Length;

            if (measuredStartHtml != startHtml || measuredStartFragment != startFragment || measuredEndFragment != endFragment || measuredEndHtml != endHtml)
            {
                header = string.Format(HeaderTemplate, measuredStartHtml, measuredEndHtml, measuredStartFragment, measuredEndFragment);
                payload = header + fullHtml;
            }

            return payload;
        }

        private static System.Drawing.Color AdjustColorForClipboard(System.Drawing.Color color)
        {
            // Only adjust lime-green text, which can be too bright on white backgrounds.
            return color.ToArgb() == System.Drawing.Color.Lime.ToArgb()
                ? System.Drawing.Color.FromArgb(0, 180, 0)
                : color;
        }

        private static string ReplaceLimeGreen(string htmlFragment)
        {
            if (string.IsNullOrEmpty(htmlFragment))
            {
                return htmlFragment;
            }

            var darkerGreen = "rgb(0, 180, 0)";
            return htmlFragment
                .Replace("rgb(0, 255, 0)", darkerGreen)
                .Replace("rgb(0,255,0)", darkerGreen);
        }
    }
}
