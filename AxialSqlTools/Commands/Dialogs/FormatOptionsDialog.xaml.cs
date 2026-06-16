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
            LeadingCommas.IsChecked = Settings.leadingCommas;
            SemicolonBeforeCte.IsChecked = Settings.semicolonBeforeCte;
            BreakSelectElementsPerLine.IsChecked = Settings.breakSelectElementsPerLine;
            UseAssignmentAliases.IsChecked = Settings.useAssignmentAliases;
            OmitAsForTableAliases.IsChecked = Settings.omitAsForTableAliases;
            OmitAsInDeclare.IsChecked = Settings.omitAsInDeclare;
            FormatTableDefinitionsMultiline.IsChecked = Settings.formatTableDefinitionsMultiline;
            PrefixUnicodeStrings.IsChecked = Settings.prefixUnicodeStrings;
        }

        private void formatSetting_Checked(object sender, RoutedEventArgs e)
        {
            // SyncFromUi();
        }
        private void formatSetting_Unchecked(object sender, RoutedEventArgs e)
        {
            // SyncFromUi();
        }

        private void SetAllFormatOptions(bool value)
        {
            PreserveComments.IsChecked = value;
            RemoveNewLineAfterJoin.IsChecked = value;
            AddTabAfterJoinOn.IsChecked = value;
            MoveCrossJoinToNewLine.IsChecked = value;
            FormatCaseAsMultiline.IsChecked = value;
            AddNewLineBetweenStatementsInBlocks.IsChecked = value;
            BreakSprocParametersPerLine.IsChecked = value;
            UppercaseBuiltInFunctions.IsChecked = value;
            UnindentBeginEndBlocks.IsChecked = value;
            BreakVariableDefinitionsPerLine.IsChecked = value;
            BreakSprocDefinitionParametersPerLine.IsChecked = value;
            LeadingCommas.IsChecked = value;
            SemicolonBeforeCte.IsChecked = value;
            BreakSelectElementsPerLine.IsChecked = value;
            UseAssignmentAliases.IsChecked = value;
            OmitAsForTableAliases.IsChecked = value;
            OmitAsInDeclare.IsChecked = value;
            FormatTableDefinitionsMultiline.IsChecked = value;
            PrefixUnicodeStrings.IsChecked = value;
        }

        private void CheckAllOptions_Click(object sender, RoutedEventArgs e)
        {
            SetAllFormatOptions(true);
        }

        private void UncheckAllOptions_Click(object sender, RoutedEventArgs e)
        {
            SetAllFormatOptions(false);
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
            Settings.leadingCommas = LeadingCommas.IsChecked == true;
            Settings.semicolonBeforeCte = SemicolonBeforeCte.IsChecked == true;
            Settings.breakSelectElementsPerLine = BreakSelectElementsPerLine.IsChecked == true;
            Settings.useAssignmentAliases = UseAssignmentAliases.IsChecked == true;
            Settings.omitAsForTableAliases = OmitAsForTableAliases.IsChecked == true;
            Settings.omitAsInDeclare = OmitAsInDeclare.IsChecked == true;
            Settings.formatTableDefinitionsMultiline = FormatTableDefinitionsMultiline.IsChecked == true;
            Settings.prefixUnicodeStrings = PrefixUnicodeStrings.IsChecked == true;
            Settings.removeSemicolonsFromDeclare = Settings.formatTableDefinitionsMultiline;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            SyncFromUi();
            DialogResult = true;
        }
    }
}
