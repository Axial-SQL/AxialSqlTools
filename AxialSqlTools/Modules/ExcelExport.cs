using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace AxialSqlTools
{
    public static class ExcelExport
    {

        #region Create Stylesheet

        private static Font CreateConsolasFont(string fontColor, bool isBold)
        {
            // reuse your existing CreateFont, then override the name
            Font font = CreateFont(fontColor, isBold);
            font.FontName = new FontName() { Val = "Consolas" };
            return font;
        }

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
                    /*Index 12 - Bold Black Large*/ CreateFont("000000", true),
                    /*Index 13 - Consolas, regular, black*/ CreateConsolasFont("000000", false)
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
                    /*Index 4 - Bold Black Font, Dark Gray Fill, All Borders, Centered, Wrap Text*/ CreateCellFormat(1, 2, 1, centerHorizontal, centerVertical, true),
                    /*Index 27 - Consolas Font, no fill or borders*/ CreateCellFormat(13, 0, 0, leftHorizontal, null, false)
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

        private static string GetSourceQueryText()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE dte = Package.GetGlobalService(typeof(DTE)) as DTE;

            if (dte?.ActiveDocument != null)
            {

                try
                {
                    TextSelection selection = dte.ActiveDocument.Selection as TextSelection;

                    string existingCommandText = selection.Text.Trim();

                    if (!string.IsNullOrEmpty(existingCommandText))
                    {
                        return existingCommandText;
                    }

                 
                    TextDocument textDoc = dte.ActiveDocument.Object("TextDocument") as TextDocument;
                    if (textDoc != null)
                    {
                        existingCommandText = textDoc.CreateEditPoint(textDoc.StartPoint).GetText(textDoc.EndPoint).Trim();

                        if (!string.IsNullOrEmpty(existingCommandText))
                        {
                            return existingCommandText;
                        }
                    }

                }
                catch 
                {                   

                }
               
            } 
            
            return string.Empty;
        }

        public static void SaveDataTableToExcel(List<DataTable> dataTables, string filePath)
        {

            var excelSettings = SettingsManager.GetExcelExportSettings();

            // detect shift state
            bool isShiftPressed = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);

            // include the source only when includeSourceQuery XOR shiftPressed is true
            bool includeSource = excelSettings.includeSourceQuery ^ isShiftPressed;

            // now pull it or leave it empty
            string sourceQuery = includeSource ? GetSourceQueryText() : string.Empty;

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                workbookPart.Workbook.AppendChild(new BookViews(new WorkbookView()));

                WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = CreateStylesheet();

                Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                uint sheetId = 1;

                // --- existing DataTable sheets ---
                foreach (DataTable dataTable in dataTables)
                {
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    // set up columns
                    Columns columns = new Columns();
                    uint colNum = 1;
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        Column col = new Column() { Min = colNum, Max = colNum };
                        if (column.ExtendedProperties.ContainsKey("columnWidthInPixels"))
                        {
                            col.Width = Convert.ToDouble(column.ExtendedProperties["columnWidthInPixels"]) * 0.17;
                            col.CustomWidth = true;
                        }
                        columns.Append(col);
                        colNum++;
                    }
                    worksheetPart.Worksheet.AppendChild(columns);

                    // add sheet data
                    SheetData sheetData = new SheetData();
                    worksheetPart.Worksheet.AppendChild(sheetData);

                    // append sheet to workbook
                    Sheet sheet = new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = sheetId,
                        Name = "QueryResult_" + sheetId
                    };
                    sheets.Append(sheet);

                    // header row
                    Row headerRow = new Row();
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        string columnName = column.ExtendedProperties.ContainsKey("columnName")
                            ? (string)column.ExtendedProperties["columnName"]
                            : column.ColumnName;

                        Cell cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(columnName),
                            StyleIndex = 4  // bold, fill, borders, etc.
                        };
                        headerRow.AppendChild(cell);
                    }
                    sheetData.AppendChild(headerRow);

                    // data rows
                    foreach (DataRow dataRow in dataTable.Rows)
                    {
                        Row newRow = new Row();
                        foreach (DataColumn column in dataTable.Columns)
                        {
                            Cell cell = new Cell();
                            object value = dataRow[column];

                            if (value is DBNull)
                            {
                                cell.DataType = CellValues.String;
                                cell.CellValue = new CellValue();
                            }
                            else if (column.DataType == typeof(int)
                                     || column.DataType == typeof(long)
                                     || column.DataType == typeof(double)
                                     || column.DataType == typeof(decimal)
                                     || column.DataType == typeof(short)
                                     || column.DataType == typeof(byte))
                            {
                                cell.DataType = CellValues.Number;
                                cell.CellValue = new CellValue(value.ToString());
                            }
                            else if (column.DataType == typeof(bool))
                            {
                                cell.DataType = CellValues.Boolean;
                                cell.CellValue = new CellValue((bool)value);
                            }
                            else
                            {
                                string text = value.ToString();
                                if (text.Length > 32767)
                                    text = text.Substring(0, 32767);
                                cell.DataType = CellValues.String;
                                cell.CellValue = new CellValue(text);
                            }

                            newRow.AppendChild(cell);
                        }
                        sheetData.AppendChild(newRow);
                    }


                    //--------------------------------------------
                    // Add AutoFilter to the header row
                    if (excelSettings.addAutofilter)
                    {
                        // 1) compute the last column letter
                        string lastCol = GetExcelColumnName(dataTable.Columns.Count);

                        // 2) build the filter range (header is always row 1)
                        string filterRef = $"A1:{lastCol}1";

                        // 3) insert the AutoFilter element
                        var autoFilter = new AutoFilter() { Reference = filterRef };
                        // make sure sheetData is the <sheetData> you appended rows into
                        worksheetPart.Worksheet.InsertAfter(autoFilter, sheetData);
                    }
                    //--------------------------------------------

                    // freeze header row
                    SheetViews sheetViews = new SheetViews();
                    worksheetPart.Worksheet.InsertAt(sheetViews, 0);
                    bool isActive = (sheetId == 1);
                    SheetView sheetView = new SheetView() { TabSelected = isActive, WorkbookViewId = 0 };
                    Pane pane = new Pane()
                    {
                        VerticalSplit = 1,
                        TopLeftCell = "A2",
                        ActivePane = PaneValues.BottomLeft,
                        State = PaneStateValues.Frozen
                    };
                    sheetView.Append(pane);
                    sheetViews.Append(sheetView);

                    worksheetPart.Worksheet.Save();
                    sheetId++;
                }

                // --- new Source sheet ---
                if (!string.IsNullOrEmpty(sourceQuery))
                {
                    WorksheetPart sourcePart = workbookPart.AddNewPart<WorksheetPart>();
                    // create a Worksheet with an empty SheetData
                    SheetData sourceData = new SheetData();
                    sourcePart.Worksheet = new Worksheet(sourceData);

                    // make column A about 10× the default 8.43-char width (≈84.3)
                    Columns sourceColumns = new Columns();
                    sourceColumns.Append(new Column()
                    {
                        Min = 1,
                        Max = 1,
                        Width = 8.43 * 10,
                        CustomWidth = true
                    });
                    // insert as the very first child of the worksheet
                    sourcePart.Worksheet.InsertAt(sourceColumns, 0);
                    //-------------------------------------------

                    Row headerRow = new Row();
                    Cell cellHeader = new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue("Source Query"),
                        StyleIndex = 4  // bold, fill, borders, etc.
                    };
                    headerRow.AppendChild(cellHeader);
                    sourceData.AppendChild(headerRow);

                    //-------------------------------------------
                    // Export source query text

                    // split on CRLF or LF
                    string[] lines = sourceQuery.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                    foreach (string line in lines)
                    {
                        // make an InlineString cell that preserves leading tabs/spaces
                        Cell cell = new Cell
                        {
                            DataType = CellValues.InlineString,
                            StyleIndex = 27    // use your Consolas style
                        };

                        InlineString inlineStr = new InlineString();
                        Text t = new Text(line){ Space = SpaceProcessingModeValues.Preserve };

                        inlineStr.AppendChild(t);
                        cell.AppendChild(inlineStr);

                        Row row = new Row();
                        row.AppendChild(cell);
                        sourceData.AppendChild(row);

                    }

                    // append the Source sheet
                    Sheet sourceSheet = new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(sourcePart),
                        SheetId = sheetId,
                        Name = "Source"
                    };
                    sheets.Append(sourceSheet);

                    sourcePart.Worksheet.Save();
                }

                workbookPart.Workbook.Save();
            }
        }
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = "";
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

    }


}
