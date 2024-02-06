using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    public static class ExcelExport
    {

        public static void SaveDataTableToExcel(DataTable dataTable, string filePath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                Sheet sheet = new Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "QueryResult"
                };
                sheets.Append(sheet);

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Add column headers
                Row headerRow = new Row();
                foreach (DataColumn column in dataTable.Columns)
                {
                    Cell cell = new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(column.ColumnName)
                    };
                    headerRow.AppendChild(cell);
                }
                sheetData.AppendChild(headerRow);

                // Add data rows
                foreach (DataRow dataRow in dataTable.Rows)
                {
                    Row newRow = new Row();
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        var ColumnType = CellValues.String;
                        if (dataTable.Columns[column.Ordinal].DataType == typeof(int) ||
                            dataTable.Columns[column.Ordinal].DataType == typeof(long))
                        { ColumnType = CellValues.Number; }
                        else if (dataTable.Columns[column.Ordinal].DataType == typeof(DateTime))
                        { ColumnType = CellValues.Date; }

                        Cell cell = new Cell
                        {
                            DataType = ColumnType,
                            CellValue = new CellValue(dataRow[column].ToString())
                        };
                        newRow.AppendChild(cell);
                    }
                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
            }
        }


    }
}
