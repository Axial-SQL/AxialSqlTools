using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using static AxialSqlTools.AxialSqlToolsPackage;

namespace AxialSqlTools
{
    public sealed class SQLBuildsLoadResult
    {
        public SQLBuildsData Data { get; set; } = new SQLBuildsData();
        public string Source { get; set; }
        public string Error { get; set; }
        public string Details { get; set; }
        public List<string> Warnings { get; } = new List<string>();
        public DateTime LoadedUtc { get; set; } = DateTime.UtcNow;
        public bool HasData => Data?.Builds != null && Data.Builds.Values.Any(v => v != null && v.Count > 0);
        internal void Warn(string message)
        {
            if (Warnings.Count < 100) Warnings.Add(message);
            else if (Warnings.Count == 100) Warnings.Add("Additional workbook issues were omitted.");
        }
    }

    internal static class SQLBuilds
    {
        internal const string SourceUrl = "https://aka.ms/sqlserverbuilds";
        private static readonly string[] SheetNames = { "2025", "2022", "2019", "2017", "2016", "2014", "2012" };

        public static SQLBuildsLoadResult DownloadSqlServerBuildInfo(string localFile = null)
        {
            string source = localFile ?? SourceUrl;
            try
            {
                byte[] bytes;
                if (localFile != null) bytes = File.ReadAllBytes(localFile);
                else using (var client = new TimedWebClient()) bytes = client.DownloadData(SourceUrl);
                return Parse(bytes, source);
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "Could not load SQL Server build information.");
                return new SQLBuildsLoadResult { Source = source,
                    Error = ex is WebException ? "Could not download SQL Server build information. Check your connection or proxy, then retry." :
                        "Could not read the selected workbook. Check that the file exists and is accessible.",
                    Details = ex.GetType().Name + ": " + ex.Message };
            }
        }

        internal static SQLBuildsLoadResult Parse(byte[] bytes, string source)
        {
            var result = new SQLBuildsLoadResult { Source = source };
            try
            {
                if (bytes == null || bytes.Length == 0) throw new InvalidDataException("The downloaded file is empty.");
                if (bytes.Length >= 8 && bytes.Take(8).SequenceEqual(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }))
                    throw new InvalidDataException("The file is an encrypted/protected Excel workbook or a legacy .xls file. An unencrypted .xlsx workbook is required.");
                if (bytes.Length < 4 || bytes[0] != 'P' || bytes[1] != 'K')
                    throw new InvalidDataException("The response is not an .xlsx workbook. The source may have returned an HTML page or another file format.");
                using (var stream = new MemoryStream(bytes, false))
                using (var document = SpreadsheetDocument.Open(stream, false))
                {
                    var workbook = document.WorkbookPart;
                    if (workbook?.Workbook?.Sheets == null) throw new InvalidDataException("The workbook has no worksheet list.");
                    var strings = workbook.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(s => s.InnerText).ToArray() ?? new string[0];
                    bool date1904 = workbook.Workbook.WorkbookProperties?.Date1904?.Value == true;
                    var seen = new HashSet<string>();
                    foreach (Sheet sheet in workbook.Workbook.Sheets.Elements<Sheet>())
                    {
                        string name = sheet.Name?.Value?.Trim();
                        if (!SheetNames.Contains(name)) continue;
                        seen.Add(name);
                        try
                        {
                            var part = workbook.GetPartById(sheet.Id.Value) as WorksheetPart;
                            if (part == null) throw new InvalidDataException("The worksheet part is missing.");
                            var rows = part.Worksheet.GetFirstChild<SheetData>()?.Elements<Row>().ToList();
                            if (rows == null) throw new InvalidDataException("The worksheet contains no rows.");
                            var builds = ParseSheet(rows, strings, part, name, date1904, result);
                            if (builds.Count > 0) result.Data.Builds[name] = builds.OrderByDescending(b => b.BuildNumber).ToList();
                        }
                        catch (Exception ex) { result.Warn("SQL Server " + name + ": " + ex.Message); }
                    }
                    foreach (string missing in SheetNames.Where(n => !seen.Contains(n))) result.Warn("Worksheet " + missing + " was not found.");
                }
                if (!result.HasData) result.Error = "The workbook contains no readable SQL Server builds. Its worksheets or column headers may have changed.";
            }
            catch (Exception ex)
            {
                result.Error = "Could not parse SQL Server build information. " + (ex is InvalidDataException ? ex.Message : "The workbook may be damaged, protected, or in an unsupported format.");
                result.Details = ex.GetType().Name + ": " + ex.Message;
                result.Data = new SQLBuildsData();
                _logger?.Error(ex, "Could not parse SQL Server build workbook.");
            }
            return result;
        }

        private static List<SQLVersionInfo> ParseSheet(List<Row> rows, string[] strings, WorksheetPart part, string sheet,
            bool date1904, SQLBuildsLoadResult result)
        {
            var builds = new List<SQLVersionInfo>();
            Dictionary<int, string> headers = null;
            int headerIndex = -1;
            for (int i = 0; i < Math.Min(50, rows.Count); i++)
            {
                var candidate = rows[i].Elements<Cell>().ToDictionary(c => ColumnIndex(c.CellReference?.Value), c => NormalizeHeader(CellText(c, strings)));
                if (candidate.Values.Contains("buildnumber")) { headers = candidate; headerIndex = i; break; }
            }
            if (headers == null) throw new InvalidDataException("Could not find the Build Number header in the first 50 rows.");
            foreach (string field in new[] { "releasedate", "kbnumber", "kburl", "cumulativeupdateorsecurityid" })
                if (!headers.Values.Contains(field)) result.Warn("SQL Server " + sheet + ": missing column " + field + ".");
            var links = part.HyperlinkRelationships.ToDictionary(h => h.Id, h => h.Uri.ToString());
            var hyperlinks = part.Worksheet.Descendants<Hyperlink>().Where(h => h.Reference != null && h.Id != null)
                .GroupBy(h => h.Reference.Value).ToDictionary(g => g.Key, g => links.TryGetValue(g.First().Id.Value, out string url) ? url : null);
            for (int i = headerIndex + 1; i < rows.Count; i++)
            {
                Row row = rows[i];
                string location = "SQL Server " + sheet + ", row " + (row.RowIndex?.Value.ToString() ?? (i + 1).ToString());
                try
                {
                    var values = new Dictionary<string, string>();
                    foreach (var cell in row.Elements<Cell>())
                    {
                        if (!headers.TryGetValue(ColumnIndex(cell.CellReference?.Value), out string header) || string.IsNullOrEmpty(header)) continue;
                        string value = CellText(cell, strings);
                        if (header == "kburl")
                        {
                            if (hyperlinks.TryGetValue(cell.CellReference.Value, out string link)) value = link;
                            else if (cell.CellFormula != null)
                            {
                                var match = Regex.Match(cell.CellFormula.Text ?? "", "^\\s*HYPERLINK\\s*\\(\\s*\"((?:[^\"]|\"\")*)\"", RegexOptions.IgnoreCase);
                                if (match.Success) value = match.Groups[1].Value.Replace("\"\"", "\"");
                            }
                        }
                        values[header] = value;
                    }
                    string build = Get(values, "buildnumber");
                    if (string.IsNullOrWhiteSpace(build)) continue;
                    if (!Version.TryParse(build.Trim(), out Version version)) { result.Warn(location + ": invalid build number '" + build + "'; row skipped."); continue; }
                    var info = new SQLVersionInfo { SqlVersion = sheet, BuildNumber = version, KbNumber = Get(values, "kbnumber"),
                        UpdateName = Get(values, "cumulativeupdateorsecurityid") };
                    string date = Get(values, "releasedate");
                    if (!TryDate(date, date1904, out DateTime releaseDate)) result.Warn(location + ": missing or invalid release date; shown as Unknown.");
                    info.ReleaseDate = releaseDate;
                    string url = Get(values, "kburl");
                    if (TryHttpUrl(url, out string validUrl)) info.Url = validUrl;
                    else if (!string.IsNullOrWhiteSpace(url)) result.Warn(location + ": invalid KB URL; link omitted.");
                    builds.Add(info);
                }
                catch (Exception ex) { result.Warn(location + ": " + ex.Message + "; row skipped."); }
            }
            if (builds.Count == 0) result.Warn("SQL Server " + sheet + ": no valid build rows found.");
            return builds;
        }

        private static string Get(Dictionary<string, string> values, string key) => values.TryGetValue(key, out string value) ? value : null;
        private static string NormalizeHeader(string value) => Regex.Replace(value ?? "", "\\s+", "").ToLowerInvariant();
        private static string CellText(Cell cell, string[] strings)
        {
            if (cell.DataType?.Value == CellValues.SharedString)
            {
                if (!int.TryParse(cell.CellValue?.Text, out int index) || index < 0 || index >= strings.Length)
                    throw new InvalidDataException("Invalid shared-string reference at " + cell.CellReference + ".");
                return strings[index];
            }
            // InnerText also includes formula text, so use the cached cell value explicitly.
            return cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? "";
        }
        private static int ColumnIndex(string reference)
        {
            if (string.IsNullOrWhiteSpace(reference)) throw new InvalidDataException("A cell has no column reference.");
            int column = 0;
            foreach (char c in reference.ToUpperInvariant().TakeWhile(char.IsLetter)) column = checked(column * 26 + c - 'A' + 1);
            if (column == 0) throw new InvalidDataException("Invalid cell reference: " + reference);
            return column - 1;
        }
        private static bool TryDate(string text, bool date1904, out DateTime value)
        {
            value = default(DateTime);
            if (string.IsNullOrWhiteSpace(text)) return false;
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double serial))
            {
                try { value = DateTime.FromOADate(serial + (date1904 ? 1462 : 0)); return true; } catch (ArgumentException) { return false; }
            }
            return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out value);
        }
        internal static bool TryHttpUrl(string value, out string url)
        {
            url = null;
            if (!Uri.TryCreate(value?.Trim(), UriKind.Absolute, out Uri uri) || (uri.Scheme != Uri.UriSchemeHttps && uri.Scheme != Uri.UriSchemeHttp)) return false;
            url = uri.AbsoluteUri; return true;
        }
        private sealed class TimedWebClient : WebClient
        {
            protected override WebRequest GetWebRequest(Uri address)
            {
                var request = base.GetWebRequest(address); request.Timeout = 30000;
                if (request is HttpWebRequest http) http.ReadWriteTimeout = 30000;
                return request;
            }
        }
    }
}
