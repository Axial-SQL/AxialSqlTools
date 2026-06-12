# Axial SQL Tools | SQL Server Management Studio 22 Productivity Add-in

Axial SQL Tools is an open-source productivity extension for SQL Server Management Studio (SSMS) 22. The project began in 2016 as a practical response to everyday SQL Server workflow friction and has grown into a community-supported tool maintained by contributors who use SSMS in real production environments.

The goal is straightforward: make common SSMS tasks faster, clearer, and less repetitive. Axial SQL Tools focuses on small, useful improvements such as clearer query status indicators, richer grid export options, query templates, formatting helpers, quick search, schema comparison, and other utilities that help database engineers stay in flow.

This extension may overlap with other free or commercial tools. It is not intended to replace every existing solution; instead, it provides a focused, open, and approachable set of enhancements for teams and individuals who want practical SSMS improvements that can evolve with community feedback.

Axial SQL Tools is free and openly available under the Apache-2.0 license. The maintainers welcome bug reports, ideas, documentation improvements, and code contributions that make the extension more useful for everyone.

If the extension helps streamline your SSMS workflow, please consider giving the repository a :star: star. Stars help other SQL Server users discover the project and show support for the contributors who keep it moving forward.

<img width="590" height="464" alt="image" src="https://github.com/user-attachments/assets/cada1f9e-4674-46e6-9fed-ca7a8724424a" />

## Main Features - (see Wiki for more details)

- [**Transaction and Column Encryption Setting Warning**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Transaction-and-Column-Encryption-Setting-Warning): <br/>
Immediately identify open transactions and clearly see when Always Encrypted is enabled.
<img src="https://github.com/Axial-SQL/AxialSqlTools/blob/main/pics/transaction-warning.png?raw=true"/> <img src="https://github.com/user-attachments/assets/bfb3becf-dd68-4ed9-a2fa-a6b08a7151ab" />

- **Precise Execution Time**: See milliseconds on the status bar.<br/>
    <img src="https://github.com/Axial-SQL/AxialSqlTools/blob/main/pics/query-duration.png?raw=true"/>

- [**Right Alignment for Numeric Values in Grid**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Align-numeric-values-in-the-grid-result-to-the-right): Automatically align numeric values in the grid result to the right. <br/>
<img src="https://github.com/Axial-SQL/AxialSqlTools/blob/main/pics/right-align-before.png?raw=true"/> -> <img src="https://github.com/Axial-SQL/AxialSqlTools/blob/main/pics/right-align-after.png?raw=true"/>

- [**Format Any TSQL Code**](https://github.com/Axial-SQL/AxialSqlTools/wiki/TSQL-Code-Formatting-with-Microsoft-ScriptDOM-library): Validate and format your TSQL code with Microsoft TSQL parser, making it more readable and maintainable.

- [**Copy Query for Web Paste**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Copy-Query-for-Web-Paste): Preserve the text format to ensure proper rendering when inserting into HTML-aware clients such as Outlook Online, Gmail, and others.
  
- [**Query Templates and Snippets**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Query-Templates-and-Snippets): Quickly access a saved collection of query templates for common tasks, reducing the time and effort required for routine work.

- [**Export Grid to Google Sheets**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Google-Sheets-Export): Quickly export results from the grid view directly into Google Sheets.
  
- [**Export Grid to Excel**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Export-Grid-To-Excel): Quickly export results from the grid view directly into Excel file.
  
- [**Export Grid to Email**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Export-Grid-to-Email): Similar to the previous feature, but with the added ability to send the file via email directly from SSMS.
  
- [**Export Grid as Temp Table**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Export-grid-results-as-a-temp-table): Convert the grid result(s) into temp table with insert statements.
  
- [**Quick Search**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Quick-Search): Quick Search helps you find SQL text quickly across one database or across all accessible databases on a server.
 
- [**Script Object Definition**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Script-Object-Definition): Quickly generate scripts for the definition of selected objects directly from the selected query text.
  
- [**Server Health Dashboard**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Server-Health-Dashboard): A quick overview of the server's most important metrics.

- [**BULK Data Transfer**](https://github.com/Axial-SQL/AxialSqlTools/wiki/BULK-Data-Transfer): A simplified UI for trivial data copy use-cases.
  
- [**Query History**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Query-History): A detailed log of executed queries, enabling auditing, tracking, and easy retrieval of past executions.

- [**Sync to GitHub**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Sync-to-GitHub): Rudimentary, low effort, manually triggered source control.

- [**Copy Column Names from Grid**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Copy-Column-Names-from-Grid): Quickly copy the column names from the results grid to the clipboard.
  
- [**Copy Cell Values As ...**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Copy-Grid-Cells-As-...): Copy cell values to the clipboard in multiple formats: INSERT, CSV, JSON, XML, or HTML.
  
## Query Library

The extension includes a query library compiled over time with ideas and best practices from SQL Server practitioners and community experts. These templates cover a wide range of scenarios and help users complete complex or repetitive tasks faster.

## Installation

To install the add-in, choose one of the following options:

1. **Compile from Source:** If you prefer to compile the tool from source code, please ensure you are using Visual Studio 2026 for compatibility. The build process will automatically copy all necessary files into the SSMS extension folder, and you will be ready to use it.

2. **Download the Release and Install the Extension:** For a quicker setup, download the most recent version from the [Releases](https://github.com/Axial-SQL/AxialSqlTools/releases) section of this GitHub repository. See the [Wiki](https://github.com/Axial-SQL/AxialSqlTools/wiki) for installation steps.

After installation and an SSMS restart, the **Axial SQL Tools** toolbar will appear in the list of available toolbars in SSMS, providing quick access to all features.<br/>
<img width="824" height="399" alt="image" src="https://github.com/user-attachments/assets/dc57882f-5a01-46dd-ac99-865ca246a255" />

## Contributing

Axial SQL Tools is maintained as an open-source project, and community participation is encouraged. Contributions can include bug reports, feature ideas, documentation improvements, query library additions, testing, and code changes.

### Submitting Ideas and Bugs

1. Submit bugs in the [Issues](https://github.com/Axial-SQL/AxialSqlTools/issues) section of this repository.
2. Add ideas in the [Discussions](https://github.com/Axial-SQL/AxialSqlTools/discussions) section.

All constructive feedback and contributions are appreciated.

## Acknowledgements

This project stands on the shoulders of the many developers, SQL Server professionals, and authors who have publicly shared their work. The maintainers are grateful for the resources, examples, discussions, and guidance published by the broader developer community. Axial SQL Tools aims to give back by keeping the extension open, useful, and accessible.

## License

This tool is distributed freely under the Apache-2.0 license, with no warranty implied.

## Disclaimer

This extension is provided "as is", with all faults, defects, and errors, and without warranty of any kind. The maintainers and contributors do not offer guaranteed support and are not responsible for any damage or loss resulting from the use of this tool.

---

Thank you for using the Axial SQL Tools SSMS Add-in. The project community hopes it makes your work with SQL Server significantly more efficient.
