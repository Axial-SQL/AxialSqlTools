using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace AxialSqlTools
{
    internal static class ExcelImport
    {
        internal sealed class WorksheetData
        {
            public DataTable Table { get; set; }

            public List<ExcelColumnMetadata> Columns { get; set; }

            public string WorksheetName { get; set; }
        }

        internal sealed class ExcelColumnMetadata
        {
            public string Name { get; set; }

            public string SqlType { get; set; }

            public Type ClrType { get; set; }

            public ColumnKind Kind { get; set; }
        }

        internal enum ColumnKind
        {
            Unknown,
            Empty,
            Boolean,
            Int32,
            Int64,
            Decimal,
            DateTime,
            Guid,
            String
        }

        private static readonly HashSet<uint> BuiltInDateFormats = new HashSet<uint>
        {
            14, 15, 16, 17, 18, 19, 20, 21, 22,
            27, 30, 36, 45, 46, 47, 50, 57
        };

        public static WorksheetData ReadWorksheet(string filePath, string worksheetName, bool firstRowHasHeaders)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Excel file path must be provided.", nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Excel file not found.", filePath);
            }

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("The workbook does not contain a workbook part.");

                Sheet sheet = ResolveWorksheet(workbookPart, worksheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(worksheetName)
                        ? "The workbook does not contain any worksheets."
                        : $"Worksheet '{worksheetName}' was not found in the workbook.");
                }

                WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                if (worksheetPart == null)
                {
                    throw new InvalidOperationException("Unable to open the worksheet part for the selected sheet.");
                }

                SharedStringTablePart sharedStrings = workbookPart.SharedStringTablePart;
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                if (sheetData == null)
                {
                    throw new InvalidOperationException("The worksheet does not contain any data.");
                }

                Dictionary<int, string> headerRow = null;
                List<Dictionary<int, string>> rows = new List<Dictionary<int, string>>();
                int maxColumnIndex = -1;

                foreach (Row row in sheetData.Elements<Row>())
                {
                    Dictionary<int, string> rowData = new Dictionary<int, string>();
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        int columnIndex = GetColumnIndex(cell.CellReference);
                        if (columnIndex < 0)
                        {
                            continue;
                        }

                        string value = GetCellValue(cell, sharedStrings, workbookPart);
                        if (value == null)
                        {
                            continue;
                        }

                        rowData[columnIndex] = value;
                        if (columnIndex > maxColumnIndex)
                        {
                            maxColumnIndex = columnIndex;
                        }
                    }

                    if (rowData.Count == 0)
                    {
                        continue;
                    }

                    if (firstRowHasHeaders && headerRow == null)
                    {
                        headerRow = rowData;
                        continue;
                    }

                    rows.Add(rowData);
                }

                if (maxColumnIndex < 0)
                {
                    throw new InvalidOperationException("The worksheet does not contain any populated cells.");
                }

                int columnCount = maxColumnIndex + 1;
                string[] columnNames = BuildColumnNames(headerRow, columnCount);
                List<ExcelColumnMetadata> metadata = BuildColumnMetadata(columnNames, rows);
                string sheetName = sheet.Name?.Value ?? "Worksheet";
                DataTable dataTable = BuildDataTable(sheetName, metadata, rows);

                return new WorksheetData
                {
                    WorksheetName = sheetName,
                    Columns = metadata,
                    Table = dataTable
                };
            }
        }

        private static Sheet ResolveWorksheet(WorkbookPart workbookPart, string worksheetName)
        {
            IEnumerable<Sheet> sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>();
            if (sheets == null)
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(worksheetName))
            {
                return sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, worksheetName, StringComparison.OrdinalIgnoreCase));
            }

            return sheets.FirstOrDefault();
        }

        private static string[] BuildColumnNames(Dictionary<int, string> headerRow, int columnCount)
        {
            string[] names = new string[columnCount];
            for (int i = 0; i < columnCount; i++)
            {
                string candidate = null;
                if (headerRow != null && headerRow.TryGetValue(i, out string headerValue))
                {
                    candidate = headerValue?.Trim();
                }

                if (string.IsNullOrWhiteSpace(candidate))
                {
                    candidate = $"Column{i + 1}";
                }

                names[i] = candidate;
            }

            HashSet<string> usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < names.Length; i++)
            {
                string baseName = names[i];
                string uniqueName = baseName;
                int suffix = 1;

                while (!usedNames.Add(uniqueName))
                {
                    uniqueName = $"{baseName}_{suffix++}";
                }

                names[i] = uniqueName;
            }

            return names;
        }

        private static List<ExcelColumnMetadata> BuildColumnMetadata(string[] columnNames, List<Dictionary<int, string>> rows)
        {
            var metadata = new List<ExcelColumnMetadata>(columnNames.Length);
            for (int i = 0; i < columnNames.Length; i++)
            {
                ColumnKind kind = DetermineColumnKind(i, rows);
                metadata.Add(new ExcelColumnMetadata
                {
                    Name = columnNames[i],
                    Kind = kind,
                    ClrType = GetClrType(kind),
                    SqlType = GetSqlType(kind)
                });
            }

            return metadata;
        }

        private static ColumnKind DetermineColumnKind(int columnIndex, List<Dictionary<int, string>> rows)
        {
            ColumnKind current = ColumnKind.Unknown;
            foreach (var row in rows)
            {
                if (!row.TryGetValue(columnIndex, out string value) || string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                ColumnKind valueKind = InferKind(value);
                if (current == ColumnKind.Unknown || current == ColumnKind.Empty)
                {
                    current = valueKind;
                    continue;
                }

                if (valueKind == ColumnKind.Empty)
                {
                    continue;
                }

                if (current == valueKind)
                {
                    continue;
                }

                ColumnKind promoted = PromoteKinds(current, valueKind);
                current = promoted;

                if (current == ColumnKind.String)
                {
                    break;
                }
            }

            if (current == ColumnKind.Unknown || current == ColumnKind.Empty)
            {
                current = ColumnKind.String;
            }

            return current;
        }

        private static ColumnKind PromoteKinds(ColumnKind existing, ColumnKind incoming)
        {
            if (existing == incoming)
            {
                return existing;
            }

            if (existing == ColumnKind.String || incoming == ColumnKind.String)
            {
                return ColumnKind.String;
            }

            if (existing == ColumnKind.DateTime || incoming == ColumnKind.DateTime)
            {
                return existing == ColumnKind.DateTime && incoming == ColumnKind.DateTime
                    ? ColumnKind.DateTime
                    : ColumnKind.String;
            }

            if (existing == ColumnKind.Guid || incoming == ColumnKind.Guid)
            {
                return existing == ColumnKind.Guid && incoming == ColumnKind.Guid
                    ? ColumnKind.Guid
                    : ColumnKind.String;
            }

            if (existing == ColumnKind.Boolean || incoming == ColumnKind.Boolean)
            {
                return existing == ColumnKind.Boolean && incoming == ColumnKind.Boolean
                    ? ColumnKind.Boolean
                    : ColumnKind.String;
            }

            // Numeric promotions
            if (IsNumeric(existing) && IsNumeric(incoming))
            {
                if (existing == ColumnKind.Decimal || incoming == ColumnKind.Decimal)
                {
                    return ColumnKind.Decimal;
                }

                if (existing == ColumnKind.Int64 || incoming == ColumnKind.Int64)
                {
                    return ColumnKind.Int64;
                }

                return ColumnKind.Int32;
            }

            return ColumnKind.String;
        }

        private static bool IsNumeric(ColumnKind kind)
        {
            return kind == ColumnKind.Int32 || kind == ColumnKind.Int64 || kind == ColumnKind.Decimal;
        }

        private static ColumnKind InferKind(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return ColumnKind.Empty;
            }

            if (TryParseBoolean(value, out _))
            {
                return ColumnKind.Boolean;
            }

            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _) ||
                int.TryParse(value, NumberStyles.Integer, CultureInfo.CurrentCulture, out _))
            {
                return ColumnKind.Int32;
            }

            if (long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _) ||
                long.TryParse(value, NumberStyles.Integer, CultureInfo.CurrentCulture, out _))
            {
                return ColumnKind.Int64;
            }

            if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out _) ||
                decimal.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out _))
            {
                return ColumnKind.Decimal;
            }

            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out _))
            {
                return ColumnKind.DateTime;
            }

            if (DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out _))
            {
                return ColumnKind.DateTime;
            }

            if (Guid.TryParse(value, out _))
            {
                return ColumnKind.Guid;
            }

            return ColumnKind.String;
        }

        private static DataTable BuildDataTable(string sheetName, List<ExcelColumnMetadata> metadata, List<Dictionary<int, string>> rows)
        {
            DataTable table = new DataTable(sheetName ?? "Worksheet");
            foreach (ExcelColumnMetadata column in metadata)
            {
                DataColumn dataColumn = new DataColumn(column.Name, column.ClrType ?? typeof(string))
                {
                    AllowDBNull = true
                };
                table.Columns.Add(dataColumn);
            }

            foreach (Dictionary<int, string> row in rows)
            {
                DataRow dataRow = table.NewRow();
                bool hasValue = false;

                for (int i = 0; i < metadata.Count; i++)
                {
                    if (!row.TryGetValue(i, out string rawValue) || string.IsNullOrWhiteSpace(rawValue))
                    {
                        dataRow[i] = DBNull.Value;
                        continue;
                    }

                    object typedValue = ConvertValue(rawValue, metadata[i].Kind);
                    if (typedValue == null)
                    {
                        dataRow[i] = DBNull.Value;
                        continue;
                    }

                    dataRow[i] = typedValue;
                    hasValue = true;
                }

                if (hasValue)
                {
                    table.Rows.Add(dataRow);
                }
            }

            return table;
        }

        private static object ConvertValue(string rawValue, ColumnKind kind)
        {
            switch (kind)
            {
                case ColumnKind.Boolean:
                    if (TryParseBoolean(rawValue, out bool boolValue))
                    {
                        return boolValue;
                    }
                    break;
                case ColumnKind.Int32:
                    if (int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int intValue))
                    {
                        return intValue;
                    }
                    if (int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.CurrentCulture, out intValue))
                    {
                        return intValue;
                    }
                    break;
                case ColumnKind.Int64:
                    if (long.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out long longValue))
                    {
                        return longValue;
                    }
                    if (long.TryParse(rawValue, NumberStyles.Integer, CultureInfo.CurrentCulture, out longValue))
                    {
                        return longValue;
                    }
                    break;
                case ColumnKind.Decimal:
                    if (decimal.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal decimalValue))
                    {
                        return decimalValue;
                    }
                    if (decimal.TryParse(rawValue, NumberStyles.Any, CultureInfo.CurrentCulture, out decimalValue))
                    {
                        return decimalValue;
                    }
                    break;
                case ColumnKind.DateTime:
                    if (DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime dateValue))
                    {
                        return dateValue;
                    }
                    if (DateTime.TryParse(rawValue, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out dateValue))
                    {
                        return dateValue;
                    }
                    break;
                case ColumnKind.Guid:
                    if (Guid.TryParse(rawValue, out Guid guidValue))
                    {
                        return guidValue;
                    }
                    break;
                default:
                    return rawValue;
            }

            // If parsing failed, return string to avoid data loss.
            return rawValue;
        }

        private static Type GetClrType(ColumnKind kind)
        {
            switch (kind)
            {
                case ColumnKind.Boolean:
                    return typeof(bool);
                case ColumnKind.Int32:
                    return typeof(int);
                case ColumnKind.Int64:
                    return typeof(long);
                case ColumnKind.Decimal:
                    return typeof(decimal);
                case ColumnKind.DateTime:
                    return typeof(DateTime);
                case ColumnKind.Guid:
                    return typeof(Guid);
                default:
                    return typeof(string);
            }
        }

        private static string GetSqlType(ColumnKind kind)
        {
            switch (kind)
            {
                case ColumnKind.Boolean:
                    return "BIT";
                case ColumnKind.Int32:
                    return "INT";
                case ColumnKind.Int64:
                    return "BIGINT";
                case ColumnKind.Decimal:
                    return "DECIMAL(38, 10)";
                case ColumnKind.DateTime:
                    return "DATETIME2";
                case ColumnKind.Guid:
                    return "UNIQUEIDENTIFIER";
                default:
                    return "NVARCHAR(MAX)";
            }
        }

        private static bool TryParseBoolean(string value, out bool boolValue)
        {
            if (bool.TryParse(value, out boolValue))
            {
                return true;
            }

            string trimmed = value.Trim();
            if (string.Equals(trimmed, "1", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "yes", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "y", StringComparison.OrdinalIgnoreCase))
            {
                boolValue = true;
                return true;
            }

            if (string.Equals(trimmed, "0", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "no", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "n", StringComparison.OrdinalIgnoreCase))
            {
                boolValue = false;
                return true;
            }

            boolValue = false;
            return false;
        }

        private static string GetCellValue(Cell cell, SharedStringTablePart sharedStrings, WorkbookPart workbookPart)
        {
            if (cell == null)
            {
                return null;
            }

            string text = cell.InnerText;

            if (cell.DataType == null)
                return text;

            CellValues type = cell.DataType.Value;

            if (type == CellValues.SharedString)
            {
                if (sharedStrings?.SharedStringTable != null && int.TryParse(text, out int index))
                {
                    if (index >= 0 && index < sharedStrings.SharedStringTable.ChildElements.Count)
                        return sharedStrings.SharedStringTable.ChildElements[index].InnerText;

                    return text;
                }
                return text;
            }

            if (type == CellValues.Boolean)
            {
                return text == "1" ? "TRUE" : text == "0" ? "FALSE" : text;
            }

            if (type == CellValues.Date)
            {
                if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double oaDate))
                    return DateTime.FromOADate(oaDate).ToString("o", CultureInfo.InvariantCulture);

                return text;
            }

            return text;
        }

        private static bool TryReadDateFromNumber(Cell cell, WorkbookPart workbookPart, out DateTime dateValue)
        {
            dateValue = default;
            if (cell?.StyleIndex == null)
            {
                return false;
            }

            WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart?.Stylesheet?.CellFormats == null)
            {
                return false;
            }

            uint styleIndex = cell.StyleIndex.Value;
            if (styleIndex >= stylesPart.Stylesheet.CellFormats.Count)
            {
                return false;
            }

            CellFormat cellFormat = stylesPart.Stylesheet.CellFormats.ElementAt((int)styleIndex) as CellFormat;
            if (cellFormat == null)
            {
                return false;
            }

            uint numberFormatId = cellFormat.NumberFormatId != null ? cellFormat.NumberFormatId.Value : 0;
            if (!IsDateFormat(numberFormatId, stylesPart))
            {
                return false;
            }

            if (!double.TryParse(cell.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out double oaDate))
            {
                return false;
            }

            dateValue = DateTime.FromOADate(oaDate);
            return true;
        }

        private static bool IsDateFormat(uint numberFormatId, WorkbookStylesPart stylesPart)
        {
            if (BuiltInDateFormats.Contains(numberFormatId))
            {
                return true;
            }

            if (stylesPart?.Stylesheet?.NumberingFormats == null)
            {
                return false;
            }

            foreach (NumberingFormat format in stylesPart.Stylesheet.NumberingFormats.OfType<NumberingFormat>())
            {
                if (format.NumberFormatId == null || format.NumberFormatId.Value != numberFormatId)
                {
                    continue;
                }

                string code = format.FormatCode?.Value;
                if (string.IsNullOrEmpty(code))
                {
                    continue;
                }

                string lower = code.ToLowerInvariant();
                if (lower.Contains("m") && lower.Contains("d"))
                {
                    return true;
                }

                if (lower.Contains("yy"))
                {
                    return true;
                }
            }

            return false;
        }

        private static int GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return -1;
            }

            int columnIndex = 0;
            int factor = 1;

            for (int pos = cellReference.Length - 1; pos >= 0; pos--)
            {
                char ch = cellReference[pos];
                if (char.IsLetter(ch))
                {
                    columnIndex += (char.ToUpperInvariant(ch) - 'A' + 1) * factor;
                    factor *= 26;
                }
            }

            return columnIndex - 1;
        }
    }
}
