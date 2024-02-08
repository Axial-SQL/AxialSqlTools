# SSMS Productivity Addin SQL Server Management Studio 19

This VSIX extension is created for SQL Server Management Studio 19, focusing on productivity improvements. 
It streamlines daily routine tasks, making database management more efficient and less time-consuming.

<img src="https://github.com/Axial-SQL/AxialSqlTools/blob/main/pics/main.png?raw=true"/>

***Very much work in progress. Please submit bugs/ideas in this repo.***

## Features

So far we've developed these features:

- **Format Query**: Beautify your SQL scripts with Microsoft internal TSQL parser, making them more readable and maintainable.
- **Export Grid to Excel**: Export result from the grid view directly into Excel file for further analysis or reporting.
- **Export Grid as Inserts**: Convert your grid result into insert statements.
- **Script Selected Object Definition**: Quickly generate scripts for the definition of selected objects directly from the selected query text.
- **Quick Query Templates**: Access a collection of query templates for common tasks, saving time and effort on routine queries.

## Query Templates

The extension also comes with a set of query templates, which have been compiled over time and incorporate ideas and best practices from many SQL experts. 
These templates are designed to cover a wide range of scenarios, helping you to execute complex tasks faster.

## Installation

To install the addin, you have two options:

1. **Compile from Source:** If you prefer to compile the tool from the source code, please ensure you are using Visual Studio 2017 for compatibility. After compiling, you will have a folder with the necessary files to proceed with the installation.

2. **Download the Release:** For a quicker setup, download the most recent version from the [Releases](https://github.com/Axial-SQL/AxialSqlTools/releases) section of this GitHub repository.


After obtaining the compiled files (either by compiling from source or downloading from Releases), follow these steps to install the extension:

- For SSMS 19, place the folder with the compiled files into the following directory:
`C:\Program Files (x86)\Microsoft SQL Server Management Studio 19\Common7\IDE\Extensions` 

<!--
- For SSMS 18, use this directory instead:
`C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Extensions`
-->

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
Feel free to reach out for professional assistance: `info@axial-sql.com`.

## Disclaimer
This extension is provided "as is", with all faults, defects, and errors, and without warranty of any kind. 
The creators do not offer support and are not responsible for any damage or loss resulting from the use of this tool.

---

Thank you for using the SSMS Productivity Addin for SQL Server Management Studio 19. 
We hope it makes your work with SQL Server significantly more efficient!
