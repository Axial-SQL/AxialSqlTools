using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Xml;

namespace AxialSqlTools.JobQuickView
{
    internal static class JobCommandEditorSupport
    {
        // Visual Studio's Text Editor font category, shared by the SSMS T-SQL editor.
        private static readonly Guid TextEditorFontCategory = new Guid(FontsAndColorsCategory.TextEditor);

        internal static void ApplyHostEditorFont(TextEditor editor)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (editor == null) return;

            editor.FontFamily = new FontFamily("Consolas");
            editor.FontSize = 10.0 * 96.0 / 72.0;
            try
            {
                var storage = Package.GetGlobalService(typeof(SVsFontAndColorStorage)) as IVsFontAndColorStorage;
                if (storage == null) return;
                var category = TextEditorFontCategory;
                // Read default values too, without creating or changing host settings.
                uint flags = (uint)(__FCSTORAGEFLAGS.FCSF_READONLY | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS);
                if (storage.OpenCategory(ref category, flags) < 0) return;
                try
                {
                    var font = new FontInfo[1];
                    if (storage.GetFont(new LOGFONTW[1], font) < 0) return;
                    if (!string.IsNullOrWhiteSpace(font[0].bstrFaceName))
                        editor.FontFamily = new FontFamily(font[0].bstrFaceName);
                    // The host stores points; WPF uses 1/96-inch units and handles monitor DPI.
                    if (font[0].wPointSize > 0)
                        editor.FontSize = font[0].wPointSize * 96.0 / 72.0;
                }
                finally
                {
                    storage.CloseCategory();
                }
            }
            catch (Exception ex)
            {
                // Typography must not prevent the job window from opening.
                AxialSqlToolsPackage._logger?.Debug("Unable to read the SSMS editor font ({0}).", ex.GetType().Name);
            }
        }

        internal static void ApplyTheme(TextEditor editor, FrameworkElement scope, string subsystem)
        {
            // Reuse the host-aware selection, line number, SQL and high-contrast treatment.
            SqlEditorSupport.ApplyTheme(editor, scope);
            if (string.Equals(subsystem, "TSQL", StringComparison.OrdinalIgnoreCase)) return;

            string definition = string.Equals(subsystem, "PowerShell", StringComparison.OrdinalIgnoreCase) ? PowerShell :
                string.Equals(subsystem, "CmdExec", StringComparison.OrdinalIgnoreCase) ? CommandShell : null;
            if (definition == null)
            {
                // SSIS, replication and other subsystem commands are not T-SQL.
                editor.SyntaxHighlighting = null;
                return;
            }
            using (var textReader = new StringReader(definition))
            using (var xmlReader = XmlReader.Create(textReader))
            {
                var highlighting = HighlightingLoader.Load(xmlReader, HighlightingManager.Instance);
                var background = VsThemeBrushResolver.GetBrushColor(editor.Background, SystemColors.WindowColor);
                var foreground = VsThemeBrushResolver.GetBrushColor(editor.Foreground, SystemColors.WindowTextColor);
                foreach (var color in highlighting.NamedHighlightingColors)
                {
                    var original = color.Foreground?.GetColor(null) ?? foreground;
                    color.Foreground = new SimpleHighlightingBrush(SystemParameters.HighContrast ? foreground :
                        VsThemeBrushResolver.EnsureTextContrast(original, background, foreground));
                }
                editor.SyntaxHighlighting = highlighting;
            }
        }

        private const string PowerShell = @"<SyntaxDefinition name='PowerShell' xmlns='http://icsharpcode.net/sharpdevelop/syntaxdefinition/2008'>
  <Color name='Comment' foreground='#008000' />
  <Color name='String' foreground='#A31515' />
  <Color name='Keyword' foreground='#0000FF' fontWeight='bold' />
  <Color name='Variable' foreground='#795E26' />
  <Color name='Command' foreground='#AF00DB' />
  <RuleSet ignoreCase='true'>
    <Span color='Comment' multiline='true'><Begin>&lt;#</Begin><End>#&gt;</End></Span>
    <Span color='Comment'><Begin>\#</Begin></Span>
    <Span color='String' multiline='true'><Begin>@&quot;$</Begin><End>^&quot;@</End></Span>
    <Span color='String' multiline='true'><Begin>@'$</Begin><End>^'@</End></Span>
    <Span color='String'><Begin>&quot;</Begin><End>&quot;</End><RuleSet><Span begin='`.' end='' /></RuleSet></Span>
    <Span color='String'><Begin>'</Begin><End>'</End><RuleSet><Span begin='&apos;&apos;' end='' /></RuleSet></Span>
    <Keywords color='Keyword'><Word>begin</Word><Word>break</Word><Word>catch</Word><Word>class</Word><Word>continue</Word><Word>data</Word><Word>do</Word><Word>dynamicparam</Word><Word>else</Word><Word>elseif</Word><Word>end</Word><Word>exit</Word><Word>filter</Word><Word>finally</Word><Word>for</Word><Word>foreach</Word><Word>from</Word><Word>function</Word><Word>if</Word><Word>in</Word><Word>param</Word><Word>process</Word><Word>return</Word><Word>switch</Word><Word>throw</Word><Word>trap</Word><Word>try</Word><Word>until</Word><Word>using</Word><Word>while</Word></Keywords>
    <Rule color='Variable'>\$[\w:]+|\$\{[^}]+\}</Rule>
    <Rule color='Command'>\b[A-Za-z]+-[A-Za-z][\w-]*\b</Rule>
  </RuleSet>
</SyntaxDefinition>";

        private const string CommandShell = @"<SyntaxDefinition name='Command shell' xmlns='http://icsharpcode.net/sharpdevelop/syntaxdefinition/2008'>
  <Color name='Comment' foreground='#008000' />
  <Color name='String' foreground='#A31515' />
  <Color name='Keyword' foreground='#0000FF' fontWeight='bold' />
  <Color name='Variable' foreground='#795E26' />
  <RuleSet ignoreCase='true'>
    <Span color='Comment'><Begin>^\s*(?:@?rem\b|::)</Begin></Span>
    <Span color='String' begin='&quot;' end='&quot;' />
    <Keywords color='Keyword'><Word>call</Word><Word>cd</Word><Word>chdir</Word><Word>copy</Word><Word>del</Word><Word>dir</Word><Word>echo</Word><Word>else</Word><Word>endlocal</Word><Word>errorlevel</Word><Word>exist</Word><Word>exit</Word><Word>for</Word><Word>goto</Word><Word>if</Word><Word>in</Word><Word>mkdir</Word><Word>move</Word><Word>not</Word><Word>pause</Word><Word>popd</Word><Word>pushd</Word><Word>ren</Word><Word>rmdir</Word><Word>set</Word><Word>setlocal</Word><Word>start</Word><Word>type</Word></Keywords>
    <Rule color='Variable'>%[^%\r\n]+%|![^!\r\n]+!|%%?[A-Za-z0-9*]</Rule>
  </RuleSet>
</SyntaxDefinition>";
    }
}
