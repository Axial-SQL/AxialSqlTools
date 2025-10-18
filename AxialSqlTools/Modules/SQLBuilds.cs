using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using static AxialSqlTools.AxialSqlToolsPackage;

namespace AxialSqlTools
{
    internal class SQLBuilds
    {


        public static SQLBuildsData DownloadSqlServerBuildInfo()
        {
            try
            {
                return DownloadAndParseExcel();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "An exception occurred");

                return null;
            }

        }

        private static SQLBuildsData DownloadAndParseExcel()
        {
            SQLBuildsData data = new SQLBuildsData();
            string excelUrl = "https://aka.ms/sqlserverbuilds";

            // Download the Excel file into a byte array.
            byte[] excelBytes;
            using (WebClient webClient = new WebClient())
            {
                excelBytes = webClient.DownloadData(excelUrl);
            }

            List<string> sheetNames = new List<string> { "2025", "2022", "2019", "2017", "2016", "2014", "2012" }; // not going for older versions...

            // Load the Excel file into a MemoryStream.
            using (MemoryStream stream = new MemoryStream(excelBytes))
            {
                // Open the SpreadsheetDocument for read-only access.
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;

                    // Iterate over each sheet in the workbook.
                    foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                    {
                        string sheetName = sheet.Name;
                        WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                        if (worksheetPart == null)
                        {
                            continue;
                        }

                        if (!sheetNames.Contains(sheetName))
                        {
                            continue;
                        }

                        // Get the SheetData element.
                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                        if (sheetData == null)
                        {
                            continue;
                        }

                        var buildList = new List<SQLVersionInfo>();
                        Dictionary<int, string> headers = null;
                        bool isHeaderRow = true;

                        // Iterate over the rows in the worksheet.
                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            // The first row is assumed to be the header.
                            if (isHeaderRow)
                            {
                                headers = new Dictionary<int, string>();
                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    int colIndex = GetColumnIndex(cell.CellReference);
                                    string headerValue = GetCellValue(cell, sharedStringPart);
                                    headers[colIndex] = headerValue;
                                }
                                isHeaderRow = false;
                            }
                            else
                            {
                                // Build a dictionary of header -> cell value for this row.
                                Dictionary<string, string> rowData = new Dictionary<string, string>();

                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    int colIndex = GetColumnIndex(cell.CellReference);
                                    if (headers != null && headers.ContainsKey(colIndex))
                                    {
                                        string header = headers[colIndex];
                                        string cellValue = GetCellValue(cell, sharedStringPart);
                                        rowData[header] = cellValue;
                                    }
                                }

                                // Only add the row if it contains a non-empty "Build" value.
                                if (rowData.ContainsKey("Build Number") && !string.IsNullOrWhiteSpace(rowData["Build Number"]))
                                {
                                    SQLVersionInfo info = new SQLVersionInfo
                                    {
                                        SqlVersion = sheetName,
                                        UpdateName = rowData.ContainsKey("Cumulative Update or Security ID") ? rowData["Cumulative Update or Security ID"] : null,
                                        Url = rowData.ContainsKey("KB URL") ? rowData["KB URL"] : null,
                                        KbNumber = rowData.ContainsKey("KB Number") ? rowData["KB Number"] : null,
                                    };

                                    string BuildNumber = rowData.ContainsKey("Build Number") ? rowData["Build Number"] : null;
                                    if (BuildNumber != null)
                                    {
                                        try
                                        {
                                            info.BuildNumber = new Version(BuildNumber.Trim());
                                        }
                                        catch
                                        {
                                        }
                                    }
                                    // Parse the release date if present.
                                    if (rowData.ContainsKey("Release Date"))
                                    {
                                        string releaseDateStr = rowData["Release Date"];
                                        if (double.TryParse(releaseDateStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
                                        {
                                            try
                                            {
                                                info.ReleaseDate = DateTime.FromOADate(result);
                                            }
                                            catch { }
                                        }

                                    }
                                    buildList.Add(info);
                                }
                            }
                        }

                        // Store the build list keyed by the worksheet (SQL Server version) name.
                        data.Builds[sheetName] = buildList;
                    }
                }
            }

            return data;
        }

        /// <summary>
        /// Retrieves the text value of a cell. If the cell uses the shared string table,
        /// the method returns the appropriate string.
        /// </summary>
        /// <param name="cell">The cell to read.</param>
        /// <param name="sharedStringPart">The shared string table part (may be null).</param>
        /// <returns>The cell value as a string.</returns>
        private static string GetCellValue(Cell cell, SharedStringTablePart sharedStringPart)
        {
            if (cell == null)
                return null;

            string value = cell.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (sharedStringPart != null && int.TryParse(value, out int index))
                {
                    return sharedStringPart.SharedStringTable.ElementAt(index).InnerText;
                }
            }
            return value;
        }

        /// <summary>
        /// Converts a cell reference (e.g. "A1", "B2", "AA3") to a zero-based column index.
        /// </summary>
        /// <param name="cellReference">The cell reference string.</param>
        /// <returns>The zero-based column index.</returns>
        private static int GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
                return -1;

            // Extract only the letters.
            string columnReference = new string(cellReference.Where(c => Char.IsLetter(c)).ToArray());
            int columnIndex = 0;
            int factor = 1;
            for (int pos = columnReference.Length - 1; pos >= 0; pos--)
            {
                columnIndex += (columnReference[pos] - 'A' + 1) * factor;
                factor *= 26;
            }
            return columnIndex - 1; // Convert to zero-based index.
        }

    }
}
