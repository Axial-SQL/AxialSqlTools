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

        #region Create Stylesheet

        private static Stylesheet CreateStylesheet()
        {
            try
            {
                HorizontalAlignmentValues leftHorizontal = HorizontalAlignmentValues.Left;
                HorizontalAlignmentValues rightHorizontal = HorizontalAlignmentValues.Right;
                HorizontalAlignmentValues centerHorizontal = HorizontalAlignmentValues.Center;
                VerticalAlignmentValues topVertical = VerticalAlignmentValues.Top;
                VerticalAlignmentValues centerVertical = VerticalAlignmentValues.Center;

                return new Stylesheet(
                    new Fonts(
                    /*Index 0 - Black*/ CreateFont("000000", false),
                    /*Index 1 - Bold Black*/ CreateFont("000000", true),
                    /*Index 2 - Purple*/ CreateFont("660066", false),
                    /*Index 3 - Bold Purple*/ CreateFont("660066", true),
                    /*Index 4 - Red*/ CreateFont("990000", false),
                    /*Index 5 - Bold Red*/ CreateFont("990000", true),
                    /*Index 6 - Orange*/ CreateFont("FF6600", false),
                    /*Index 7 - Bold Orange*/ CreateFont("FF6600", true),
                    /*Index 8 - Blue*/ CreateFont("0066FF", false),
                    /*Index 9 - Bold Blue*/ CreateFont("0066FF", true),
                    /*Index 10 - Green*/ CreateFont("339900", false),
                    /*Index 11 - Bold Green*/ CreateFont("339900", true),
                    /*Index 12 - Bold Black Large*/ CreateFont("000000", true)
                        ),
                    new Fills(
                    /*Index 0 - Default Fill (None)*/ CreateFill(string.Empty, PatternValues.None),
                    /*Index 1 - Default Fill (Gray125)*/ CreateFill(string.Empty, PatternValues.Gray125),
                    /*Index 2 - Dark Gray Fill*/ CreateFill("BBBBBB", PatternValues.Solid),
                    /*Index 3 - Light Gray Fill*/ CreateFill("EEEEEE", PatternValues.Solid),
                    /*Index 4 - Yellow Gray Fill*/ CreateFill("FFCC00", PatternValues.Solid)
                        ),
                    new Borders(
                    /*Index 0 - Default Border (None)*/ CreateBorder(false, false, false, false),
                    /*Index 1 - All Borders*/ CreateBorder(true, true, true, true),
                    /*Index 2 - Top & Bottom Borders*/ CreateBorder(true, false, true, false)
                        ),
                    new CellFormats(
                    /*Index 0 - Black Font, No Fill, No Borders, Wrap Text*/ CreateCellFormat(0, 0, 0, leftHorizontal, null, true),
                    /*Index 1 - Black Font, No Fill, No Borders, Horizontally Centered*/ CreateCellFormat(0, 0, 0, centerHorizontal, null, false),
                    /*Index 2 - Bold Black Font, Dark Gray Fill, All Borders*/ CreateCellFormat(1, 2, 1, null, null, false),
                    /*Index 3 - Bold Black Font, Dark Gray Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(1, 2, 2, centerHorizontal, centerVertical, false),
                    /*Index 4 - Bold Black Font, Dark Gray Fill, All Borders, Centered*/ CreateCellFormat(1, 2, 1, centerHorizontal, centerVertical, false),
                    /*Index 5 - Bold Purple Font, Light Gray Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(3, 3, 2, centerHorizontal, centerVertical, false),
                    /*Index 6 - Bold Red Font, Light Gray Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(5, 3, 2, centerHorizontal, centerVertical, false),
                    /*Index 7 - Bold Orange Font, Light Gray Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(7, 3, 2, centerHorizontal, centerVertical, false),
                    /*Index 8 - Bold Blue Font, Light Gray Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(9, 3, 2, centerHorizontal, centerVertical, false),
                    /*Index 9 - Bold Green Font, Light Gray Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(11, 3, 2, centerHorizontal, centerVertical, false),
                    /*Index 10 - Purple Font, No Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(2, 0, 2, centerHorizontal, centerVertical, false),
                    /*Index 11 - Red Font, No Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(4, 0, 2, centerHorizontal, centerVertical, false),
                    /*Index 12 - Orange Font, No Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(6, 0, 2, centerHorizontal, centerVertical, false),
                    /*Index 13 - Blue Font , No Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(8, 0, 2, centerHorizontal, centerVertical, false),
                    /*Index 14 - Green Font, No Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(10, 0, 2, centerHorizontal, centerVertical, false),
                    /*Index 15 - Bold Black Font, Yellow Fill, All Borders, Centered, Wrap Text*/ CreateCellFormat(1, 4, 1, centerHorizontal, centerVertical, true),
                    /*Index 16 - Bold Black Font, No Fill, All Borders, Wrap Text*/ CreateCellFormat(1, 0, 1, centerHorizontal, centerVertical, true),
                    /*Index 17 - Bold Black Font, Light Gray Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(1, 3, 2, centerHorizontal, centerVertical, false),
                    /*Index 18 - Bold Black Font, No Fill, Top & Bottom Borders, Centered*/ CreateCellFormat(0, 0, 2, centerHorizontal, centerVertical, false),
                    /*Index 19 - Bold Black Font, Dark Gray Fill, Top & Bottom Borders, Centered Vertically*/ CreateCellFormat(1, 2, 1, null, centerVertical, false),
                    /*Index 20 - Black Font, No Fill, All Borders, Top Aligned, Wrap Text*/ CreateCellFormat(0, 0, 1, null, topVertical, true),
                    /*Index 21 - Black Font, No Fill, All Borders, Centered Vertically, Wrap Text*/ CreateCellFormat(0, 0, 1, null, centerVertical, true),
                    /*Index 22 - Black Font, No Fill, All Borders, Centered, Wrap Text*/ CreateCellFormat(0, 0, 1, centerHorizontal, centerVertical, true),
                    /*Index 23 - Black Font, No Fill, All Borders, Centered Vertically, Right Aligned*/ CreateCellFormat(0, 0, 1, rightHorizontal, centerVertical, false),
                    /*Index 24 - Black Font, No Fill, All Borders, Centered Horizontally, Top Aligned, Wrap Text*/ CreateCellFormat(0, 0, 1, centerHorizontal, topVertical, true),
                    /*Index 25 - Bold Black Font, No Fill, No Borders, Wrap Text*/ CreateCellFormat(1, 0, 0, leftHorizontal, topVertical, false),
                    /*Index 4 - Bold Black Font, Dark Gray Fill, All Borders, Centered, Wrap Text*/ CreateCellFormat(1, 2, 1, centerHorizontal, centerVertical, true)
                        )
                    );
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private static Font CreateFont(string fontColor, bool isBold)
        {
            try
            {
                Font font = new Font();
                font.FontSize = new FontSize() { Val = 10 };
                font.Color = new Color { Rgb = new HexBinaryValue() { Value = fontColor } };
                font.FontName = new FontName() { Val = "Calibri" };
                if (isBold)
                { font.Bold = new Bold(); }
                return font;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private static Fill CreateFill(string fillColor, PatternValues patternValue)
        {
            try
            {
                Fill fill = new Fill();
                PatternFill patternfill = new PatternFill();
                patternfill.PatternType = patternValue;
                if (!string.IsNullOrWhiteSpace(fillColor))
                { patternfill.ForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue { Value = fillColor } }; }
                fill.PatternFill = patternfill;

                return fill;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private static Border CreateBorder(bool topBorderRequired, bool rightBorderRequired, bool bottomBorderRequired, bool leftBorderRequired)
        {
            try
            {
                Border border = new Border();
                if (!topBorderRequired && !rightBorderRequired && !bottomBorderRequired && !leftBorderRequired)
                {
                    border.TopBorder = new TopBorder();
                    border.RightBorder = new RightBorder();
                    border.BottomBorder = new BottomBorder();
                    border.LeftBorder = new LeftBorder();
                    border.DiagonalBorder = new DiagonalBorder();
                }
                else
                {
                    if (topBorderRequired)
                    { border.TopBorder = new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin }; }
                    if (rightBorderRequired)
                    { border.RightBorder = new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin }; }
                    if (bottomBorderRequired)
                    { border.BottomBorder = new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin }; }
                    if (leftBorderRequired)
                    { border.LeftBorder = new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin }; }
                    border.DiagonalBorder = new DiagonalBorder();
                }
                return border;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private static CellFormat CreateCellFormat(UInt32Value fontId, UInt32Value fillId, UInt32Value borderId,
            HorizontalAlignmentValues? horizontalAlignment, VerticalAlignmentValues? verticalAlignment, bool wrapText)
        {
            try
            {
                CellFormat cellFormat = new CellFormat();
                Alignment alignment = new Alignment();
                if (horizontalAlignment != null)
                { alignment.Horizontal = horizontalAlignment; }
                if (verticalAlignment != null)
                { alignment.Vertical = verticalAlignment; }
                alignment.WrapText = wrapText;
                cellFormat.Alignment = alignment;
                cellFormat.FontId = fontId;
                cellFormat.FillId = fillId;
                cellFormat.BorderId = borderId;
                return cellFormat;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        #endregion Create Stylesheet

        public static void SaveDataTableToExcel(List<DataTable> dataTables, string filePath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = CreateStylesheet();               

                Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                uint i = 1;

                foreach (DataTable dataTable in dataTables)
                {
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    Sheet sheet = new Sheet()
                    {
                        Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = i,
                        Name = "QueryResult_" + i.ToString()
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
                            CellValue = new CellValue(column.ColumnName),
                            /*Index 4 - Bold Black Font, Dark Gray Fill, All Borders, Centered*/
                            StyleIndex = 4
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
                            //TODO - do more types 
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

                    //--------------------------------------------------------
                    // Create or get the SheetViews element
                    SheetViews sheetViews = worksheetPart.Worksheet.GetFirstChild<SheetViews>();
                    sheetViews = new SheetViews();
                    worksheetPart.Worksheet.InsertAt(sheetViews, 0); // Insert at the beginning of the Worksheet
                    // Create a SheetView element
                    SheetView sheetView = new SheetView() { TabSelected = true, WorkbookViewId = 0 };
                    // Create a Pane element that specifies freezing the first row
                    Pane pane = new Pane()
                    {
                        VerticalSplit = 1, // Freeze one row
                        TopLeftCell = "A2",
                        ActivePane = PaneValues.BottomLeft,
                        State = PaneStateValues.Frozen
                    };
                    sheetView.Append(pane);
                    sheetViews.Append(sheetView);
                    worksheetPart.Worksheet.Save();
                    //\\------------------------------------------------------

                    ////--------------------------------------------------------
                    // TODO - doesn't work properly 
                    //// try to create a new filtered table around data
                    //string columnNameStart = "A"; // Starting at column A
                    //string columnNameEnd = GetColumnName(dataTable.Columns.Count); // Get the last column name based on the DataTable column count
                    //int rowCount = dataTable.Rows.Count + 1; // Including the header row
                    //string reference = $"{columnNameStart}1:{columnNameEnd}{rowCount}";

                    //TableDefinitionPart tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
                    //tablePart.Table = new Table()
                    //{
                    //    Id = i,
                    //    Name = "Table_" + i,
                    //    DisplayName = "Table_" + i,
                    //    Reference = reference,
                    //    TotalsRowShown = false,
                    //    AutoFilter = new AutoFilter() { Reference = reference }
                    //};

                    //tablePart.Table.TableColumns = new TableColumns() { Count = (uint)dataTable.Columns.Count };
                    //uint columnIndex = 1;
                    //foreach (DataColumn column in dataTable.Columns)
                    //{
                    //    TableColumn tableColumn = new TableColumn()
                    //    {
                    //        Id = columnIndex,
                    //        Name = column.ColumnName
                    //    };
                    //    tablePart.Table.TableColumns.Append(tableColumn);
                    //    columnIndex++;
                    //}

                    //tablePart.Table.Save();
                    ////\\--------------------------------------------------------


                    i += 1;

                }

                workbookPart.Workbook.Save();
            }
        }

        // Helper method to convert column index to column name (e.g., 1 to A, 28 to AB)
        static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

    }
}
