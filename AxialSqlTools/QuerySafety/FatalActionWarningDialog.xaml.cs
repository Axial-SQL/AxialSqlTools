using System;
using System.Linq;
using System.Windows;

namespace AxialSqlTools.QuerySafety
{
    public partial class FatalActionWarningDialog : Microsoft.VisualStudio.PlatformUI.DialogWindow
    {
        private readonly ToolWindowThemeController themeController;

        internal bool AllowRepeats { get; private set; }

        internal FatalActionWarningDialog(FatalActionAnalysis analysis, string context, bool canRemember)
        {
            InitializeComponent();
            themeController = new ToolWindowThemeController(this, () => ToolWindowThemeResources.ApplySharedTheme(this));
            ContextText.Text = context;
            ActionText.Text = string.Join(Environment.NewLine, analysis.Actions.Select(a =>
                "Line " + a.Line + ": " + a.Operation + " - " + a.Target));
            if (analysis.Limitation != null) ActionText.Text += Environment.NewLine + analysis.Limitation;
            RepeatButton.IsEnabled = canRemember;
            RepeatButton.IsDefault = canRemember;
            CancelButton.IsDefault = !canRemember;
            RepeatText.Text = canRemember
                ? "Allow repeats applies only to this exact SQL in this query window and connection. Changed SQL, a different database, reconnecting or closing the window requires approval again."
                : "Repeat approval is unavailable until the query can be fully checked and the active connection session can be identified. You can still run once.";
            Loaded += (sender, args) =>
            {
                if (canRemember) RepeatButton.Focus();
                else CancelButton.Focus();
            };
        }

        private void RunOnce_Click(object sender, RoutedEventArgs e) { DialogResult = true; }
        private void AllowRepeats_Click(object sender, RoutedEventArgs e)
        {
            if (!RepeatButton.IsEnabled) return;
            AllowRepeats = true;
            DialogResult = true;
        }
    }
}
