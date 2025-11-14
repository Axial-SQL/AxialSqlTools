# Axial SQL Tools | SQL Server Management Studio 21 Productivity Addin

As an engineer with over two decades of experience in SQL Server, I've encountered my fair share of inefficiencies and limitations within SSMS. 
Motivated by these challenges, I started developing the SSMS extension back in 2016. This repository represents the fourth iteration of my work, tailored specifically for SSMS 21.
This project is a personal endeavor to solve the issues I've faced, but I would like to incorporate valuable community feedback and ideas as well.

While I recognize that this extension may overlap with existing tools, both free and paid, my passion lies in coding and tackling problems in my own unique way. The aim isn't to overshadow other solutions but to provide a complementary tool that addresses gaps and simplifies processes for engineers like myself.

I'm committed to keeping this tool free and openly available to the community. While formalized support won't be provided, I welcome feedback and suggestions to continually improve and evolve the tool based on user needs. 

If you find this extension useful and it helps streamline your workflow in SSMS, please consider giving it a :star: star on the repository. Your support not only motivates me but also helps others in the community discover this tool. 

Check it out and let me know what you think!

<img alt="image" src="https://github.com/user-attachments/assets/23ce2810-44cb-485d-852a-23f74ee0c3b0" />

## Main Features - (see Wiki for more details)

- **Transaction Warning**: See right away if you left any transaction open.<br/>
<img src="https://github.com/Axial-SQL/AxialSqlTools/blob/main/pics/transaction-warning.png?raw=true"/>

- **Precise Execution Time**: See milliseconds on the status bar.<br/>
<img src="https://github.com/Axial-SQL/AxialSqlTools/blob/main/pics/query-duration.png?raw=true"/>

- [**Format Any TSQL Code**](https://github.com/Axial-SQL/AxialSqlTools/wiki/TSQL-Code-Formatting-with-Microsoft-ScriptDOM-library): Validate and format your TSQL code with Microsoft TSQL parser, making it more readable and maintainable.
  
- [**Query Templates and Snippets**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Query-Templates-and-Snippets): Quickly access your personal collection of query templates for common tasks, reducing the time and effort required for routine work.
  
- [**Export Grid to Excel**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Export-Grid-To-Excel): Quickly export results from the grid view directly into Excel file.
  
- [**Export Grid to Email**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Export-Grid-to-Email): Similar to the previous feature, but with the added ability to send the file via email directly from SSMS.
  
- [**Export Grid as Temp Table**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Export-grid-results-as-a-temp-table): Convert the grid result(s) into temp table with insert statements.
  
- **Script Selected Object Definition**: Quickly generate scripts for the definition of selected objects directly from the selected query text.
  
- [**Server Health Dashboard**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Server-Health-Dashboard): A quick overview of the server's most important metrics..
![server-health](https://github.com/Axial-SQL/AxialSqlTools/assets/13791336/760dcb74-d73b-42c7-94fe-933e321d0044)

- [**Right Alignment for Numeric Values in Grid**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Align-numeric-values-in-the-grid-result-to-the-right): Automatically align numeric values in the grid result to the right. <br/>
<img src="https://github.com/Axial-SQL/AxialSqlTools/blob/main/pics/right-align-before.png?raw=true"/> -> <img src="https://github.com/Axial-SQL/AxialSqlTools/blob/main/pics/right-align-after.png?raw=true"/>

- [**BULK Data Transfer**](https://github.com/Axial-SQL/AxialSqlTools/wiki/BULK-Data-Transfer): A simplified UI for trivial data copy use-cases.
  
- [**Query History**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Query-History): A detailed log of executed queries, enabling auditing, tracking, and easy retrieval of past executions.

- [**Sync to GitHub**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Sync-to-GitHub): Rudimentary, low effort, manually triggered source control.

- [**Copy Column Names from Grid**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Copy-Column-Names-from-Grid): Quickly copy the column names from the results grid to the clipboard.
  
- [**Database Schema Compare**](https://github.com/Axial-SQL/AxialSqlTools/wiki/Database-Schema-Compare): DacFx-based database schema comparison tool.


## Query Library

The extension also comes with a set of queries, which have been compiled over time and incorporate ideas and best practices from many SQL experts. 
These templates are designed to cover a wide range of scenarios, helping you to execute complex tasks faster.

## Installation

To install the addin, you have two options:

1. **Compile from Source:** If you prefer to compile the tool from the source code, please ensure you are using Visual Studio 2022 for compatibility. 
The build process will automatically copy all necessary files into the SSMS extension folder, and you will be ready to use it.

2. **Download the Release:** For a quicker setup, download the most recent version from the [Releases](https://github.com/Axial-SQL/AxialSqlTools/releases) section of this GitHub repository.

After obtaining the compiled files (either by compiling from source or downloading from Releases), place the folder with the compiled files into the following directory:
`C:\Program Files\Microsoft SQL Server Management Studio 21\Release\Common7\IDE\Extensions`

> **⚠️ Warning**
>
> If your Windows security policy has blocked the `AxialSqlTools.dll` file, you may need to unlock it by using the checkbox in the file properties.

Restart SQL Server Management Studio after placing the files in the extensions directory.

After installation, the Axial SQL Tools toolbar will be available in the list of toolbars within SSMS, providing quick access to all the features.<br/>
<img src="https://github.com/Axial-SQL/AxialSqlTools/blob/main/pics/toolbar.png?raw=true"/>

## Contributing

We believe in the power of community and are open to any ideas or bug reports. 
Your contributions are what make this tool better every day. 
If you have ideas for new features or have encountered a bug, please feel free to submit them in this repository.

### Submitting Ideas and Bugs

1. Submit bugs in the [Issues](https://github.com/Axial-SQL/AxialSqlTools/issues) section of this repository.
2. Add ideas in the [Discussions](https://github.com/Axial-SQL/AxialSqlTools/discussions) section.

We appreciate your feedback and contributions!

## Acknowledgements

This project stands on the shoulders of the many developers and authors who have publicly shared their work. I am not an expert in C# or VSIX development, so the creation of this tool would not have been possible without the invaluable resources and examples published by others on GitHub. We do not claim any copyright over the work derived from these contributions. Our aim is to give back to the community by sharing this tool, and we are deeply grateful for the guidance and inspiration provided by the broader developer community.

## License
This tool is distributed freely under the Apache-2.0 license, with no warranty implied. 
It is an internal tool that we use daily at our organization. 
While we do not offer formal support or sell this tool, we are open to consulting and helping solve complex SQL Server challenges. 
Feel free to reach out for professional assistance: `albochkov03@gmail.com`.

## Disclaimer
This extension is provided "as is", with all faults, defects, and errors, and without warranty of any kind. 
The creators do not offer support and are not responsible for any damage or loss resulting from the use of this tool.

---

Thank you for using the Axial SQL Tools SSMS Addin! 
We hope it makes your work with SQL Server significantly more efficient!
