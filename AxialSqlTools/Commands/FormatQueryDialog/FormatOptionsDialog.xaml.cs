using System.Windows;
using static AxialSqlTools.SettingsManager;

namespace AxialSqlTools
{
    public partial class FormatOptionsDialog : Window
    {
        public TSqlCodeFormatSettings Settings { get; }

        public FormatOptionsDialog(TSqlCodeFormatSettings initial = null)
        {
            InitializeComponent();
            Settings = initial ?? new TSqlCodeFormatSettings();
            ApplySettingsToUi();
        }

        private void ApplySettingsToUi()
        {
            PreserveComments.IsChecked = Settings.preserveComments;
            RemoveNewLineAfterJoin.IsChecked = Settings.removeNewLineAfterJoin;
            AddTabAfterJoinOn.IsChecked = Settings.addTabAfterJoinOn;
            MoveCrossJoinToNewLine.IsChecked = Settings.moveCrossJoinToNewLine;
            FormatCaseAsMultiline.IsChecked = Settings.formatCaseAsMultiline;
            AddNewLineBetweenStatementsInBlocks.IsChecked = Settings.addNewLineBetweenStatementsInBlocks;
            BreakSprocParametersPerLine.IsChecked = Settings.breakSprocParametersPerLine;
            UppercaseBuiltInFunctions.IsChecked = Settings.uppercaseBuiltInFunctions;
            UnindentBeginEndBlocks.IsChecked = Settings.unindentBeginEndBlocks;
            BreakVariableDefinitionsPerLine.IsChecked = Settings.breakVariableDefinitionsPerLine;
            BreakSprocDefinitionParametersPerLine.IsChecked = Settings.breakSprocDefinitionParametersPerLine;
        }

        private void formatSetting_Checked(object sender, RoutedEventArgs e)
        {
            // SyncFromUi();
        }
        private void formatSetting_Unchecked(object sender, RoutedEventArgs e)
        {
            // SyncFromUi();
        }

        private void SyncFromUi()
        {
            Settings.preserveComments = PreserveComments.IsChecked == true;
            Settings.removeNewLineAfterJoin = RemoveNewLineAfterJoin.IsChecked == true;
            Settings.addTabAfterJoinOn = AddTabAfterJoinOn.IsChecked == true;
            Settings.moveCrossJoinToNewLine = MoveCrossJoinToNewLine.IsChecked == true;
            Settings.formatCaseAsMultiline = FormatCaseAsMultiline.IsChecked == true;
            Settings.addNewLineBetweenStatementsInBlocks = AddNewLineBetweenStatementsInBlocks.IsChecked == true;
            Settings.breakSprocParametersPerLine = BreakSprocParametersPerLine.IsChecked == true;
            Settings.uppercaseBuiltInFunctions = UppercaseBuiltInFunctions.IsChecked == true;
            Settings.unindentBeginEndBlocks = UnindentBeginEndBlocks.IsChecked == true;
            Settings.breakVariableDefinitionsPerLine = BreakVariableDefinitionsPerLine.IsChecked == true;
            Settings.breakSprocDefinitionParametersPerLine = BreakSprocDefinitionParametersPerLine.IsChecked == true;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            SyncFromUi();
            DialogResult = true;
        }
    }
}
