using AxialSqlTools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using System.Text;

class Program
{
    static int count;
    static void Check(bool condition) { if (!condition) throw new Exception("Assertion failed"); }
    static void Test(string name, Action test) { test(); count++; Console.WriteLine("PASS " + name); }
    static Cell Text(string reference, string text) => new Cell { CellReference = reference, DataType = CellValues.InlineString, InlineString = new InlineString(new Text(text)) };
    static Row Header(uint index = 1) => new Row(new[] { Text("A" + index, " Build Number "), Text("C" + index, "Release Date"), Text("D" + index, "KB URL"), Text("E" + index, "KB Number"), Text("F" + index, "Cumulative Update or Security ID") }) { RowIndex = index };
    static Row Build(uint row = 2, string version = "16.0.1000.6", string date = "45000", string url = "https://example.com/kb") => new Row(Text("A" + row, version), Text("C" + row, date), Text("D" + row, url)) { RowIndex = row };
    static byte[] Book(Action<WorkbookPart, WorksheetPart, SheetData> edit = null, string name = "2022", bool date1904 = false)
    {
        using (var stream = new MemoryStream())
        {
            using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var wb = doc.AddWorkbookPart(); wb.Workbook = new Workbook(new WorkbookProperties { Date1904 = date1904 });
                var ws = wb.AddNewPart<WorksheetPart>(); var rows = new SheetData(Header(), Build()); ws.Worksheet = new Worksheet(rows);
                wb.Workbook.Append(new Sheets(new Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = name }));
                edit?.Invoke(wb, ws, rows); wb.Workbook.Save(); ws.Worksheet.Save();
            }
            return stream.ToArray();
        }
    }
    static SQLBuildsLoadResult Parse(byte[] bytes) => SQLBuilds.Parse(bytes, "fixture.xlsx");
    static AxialSqlTools.AxialSqlToolsPackage.SQLVersionInfo First(SQLBuildsLoadResult result) => result.Data.Builds.Values.SelectMany(x => x).First();
    static void Main(string[] args)
    {
        Test("empty response returns explicit non-null error result", () => { var r=Parse(new byte[0]); Check(r.Error.Contains("empty") && !r.HasData && r.Data!=null); });
        Test("HTML instead of Excel", () => Check(Parse(Encoding.UTF8.GetBytes("<html>Sign in</html>")).Error.Contains("not an .xlsx")));
        Test("OLE encrypted or legacy workbook identified", () => Check(Parse(new byte[] {0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1}).Error.Contains("encrypted/protected")));
        Test("damaged ZIP is reported", () => Check(Parse(new byte[] {80,75,3,4,0,0,0,0}).Error!=null));
        Test("inline strings and sparse column references", () => { var r=Parse(Book()); Check(r.HasData && First(r).BuildNumber.ToString()=="16.0.1000.6" && First(r).Url=="https://example.com/kb"); });
        Test("headers can follow an introductory row", () => { var r=Parse(Book((w,p,d)=>d.PrependChild(new Row(Text("A0","Build history"))))); Check(r.HasData); });
        Test("missing supported sheets are explicit warnings", () => { var r=Parse(Book()); Check(r.Warnings.Any(w=>w.Contains("2025"))); });
        Test("unsupported worksheets cannot produce silent empty success", () => { var r=Parse(Book(name:"Readme")); Check(!r.HasData && r.Error.Contains("no readable")); });
        Test("missing header reported with worksheet", () => { var r=Parse(Book((w,p,d)=>d.FirstChild.Remove())); Check(!r.HasData && r.Warnings.Any(w=>w.Contains("2022") && w.Contains("header"))); });
        Test("bad build rows skipped without losing valid rows", () => { var r=Parse(Book((w,p,d)=>d.Append(Build(3,"bad")))); Check(r.HasData && r.Data.Builds["2022"].Count==1 && r.Warnings.Any(w=>w.Contains("row 3"))); });
        Test("bad shared-string index isolated to row", () => { var r=Parse(Book((w,p,d)=>d.Append(new Row(new Cell { CellReference="A3", DataType=CellValues.SharedString, CellValue=new CellValue("999") }) { RowIndex=3 }))); Check(r.HasData && r.Warnings.Any(w=>w.Contains("shared-string"))); });
        Test("shared-string table is resolved", () => { var r=Parse(Book((w,p,d)=> { var shared=w.AddNewPart<SharedStringTablePart>(); shared.SharedStringTable=new SharedStringTable(new SharedStringItem(new Text("16.0.1111.1"))); var cell=d.LastChild.GetFirstChild<Cell>(); cell.RemoveAllChildren(); cell.DataType=CellValues.SharedString; cell.CellValue=new CellValue("0"); })); Check(First(r).BuildNumber.ToString()=="16.0.1111.1"); });
        Test("formula cached values do not include formula source", () => { var r=Parse(Book((w,p,d)=> { var cell=d.LastChild.GetFirstChild<Cell>(); cell.RemoveAllChildren(); cell.DataType=CellValues.String; cell.CellFormula=new CellFormula("CONCAT(16,\".0.1000.6\")"); cell.CellValue=new CellValue("16.0.1000.6"); })); Check(r.HasData); });
        Test("HYPERLINK formula yields URL not display label", () => { var r=Parse(Book((w,p,d)=> { var cell=d.LastChild.Elements<Cell>().Last(); cell.RemoveAllChildren(); cell.DataType=CellValues.String; cell.CellFormula=new CellFormula("HYPERLINK(\"https://example.com/formula\",\"KB\")"); cell.CellValue=new CellValue("KB"); })); Check(First(r).Url=="https://example.com/formula"); });
        Test("worksheet hyperlink relationships supported", () => { var r=Parse(Book((w,p,d)=> { var link=p.AddHyperlinkRelationship(new Uri("https://example.com/relationship"),true); p.Worksheet.Append(new Hyperlinks(new Hyperlink { Reference="D2",Id=link.Id })); })); Check(First(r).Url.EndsWith("relationship")); });
        Test("unsafe URLs are omitted without losing build", () => { var r=Parse(Book((w,p,d)=>{ d.LastChild.Elements<Cell>().Last().InlineString=new InlineString(new Text("file:///tmp/test.exe")); })); Check(r.HasData && First(r).Url==null && r.Warnings.Any(w=>w.Contains("invalid KB URL"))); });
        Test("invalid release date preserved as unknown with warning", () => { var r=Parse(Book((w,p,d)=> { d.LastChild.Elements<Cell>().ElementAt(1).InlineString=new InlineString(new Text("not a date")); })); Check(First(r).ReleaseDate==default(DateTime) && r.Warnings.Any(w=>w.Contains("release date"))); });
        Test("text dates accepted", () => { var r=Parse(Book((w,p,d)=> { d.LastChild.Elements<Cell>().ElementAt(1).InlineString=new InlineString(new Text("2026-09-11")); })); Check(First(r).ReleaseDate==new DateTime(2026,9,11)); });
        Test("1904 date system applied", () => { var a=First(Parse(Book())).ReleaseDate; var b=First(Parse(Book(date1904:true))).ReleaseDate; Check((b-a).TotalDays==1462); });
        Test("warnings bounded for badly damaged sheets", () => { var r=Parse(Book((w,p,d)=> { for(uint i=3;i<200;i++) d.Append(Build(i,"invalid")); })); Check(r.Warnings.Count==101); });
        Test("missing local file produces readable error", () => { var r=SQLBuilds.DownloadSqlServerBuildInfo(Path.Combine(Path.GetTempPath(),Guid.NewGuid()+".xlsx")); Check(!r.HasData && r.Error.Contains("read") && r.Details!=null); });
        if (args.Length>0) Test("current Microsoft download fails cleanly",()=> { var r=Parse(File.ReadAllBytes(args[0])); Check(r.Error!=null && r.Error.Contains("encrypted/protected") && !r.HasData); });
        Console.WriteLine(count+" tests passed.");
    }
}
