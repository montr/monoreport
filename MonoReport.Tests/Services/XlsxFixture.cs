using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

internal sealed class XlsxFixture
{
    internal byte[] Create(Action<WorkbookPart, WorksheetPart>? customize = null,
        SpreadsheetDocumentType kind = SpreadsheetDocumentType.Workbook)
    {
        using var output = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(output, kind))
        {
            var workbook = document.AddWorkbookPart();
            workbook.Workbook = new Workbook();
            var worksheet = workbook.AddNewPart<WorksheetPart>();
            worksheet.Worksheet = new Worksheet(new SheetData(new Row(Text("B3", "{Name}")) { RowIndex = 3 }));
            workbook.Workbook.Append(new Sheets(new Sheet { Name = "Report", SheetId = 1,
                Id = workbook.GetIdOfPart(worksheet) }));
            workbook.Workbook.Append(new DefinedNames(new DefinedName("'Report'!$B$3:$D$5")
                { Name = "_xlnm.Print_Area", LocalSheetId = 0 }));
            customize?.Invoke(workbook, worksheet);
        }
        return output.ToArray();
    }

    internal Cell Text(string address, string text) => new() { CellReference = address,
        DataType = CellValues.InlineString, InlineString = new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve }) };

    internal ServiceProvider Services() => new ServiceCollection().AddMonoReport().BuildServiceProvider();

    // Interpret the stored ST_Xstring once, independently of production encoding and SDK InnerText.
    internal string ReadCellText(byte[] source, string address)
    {
        using var package = SpreadsheetDocument.Open(new MemoryStream(source), false);
        using var stream = package.WorkbookPart!.WorksheetParts.Single().GetStream();
        var xml = System.Xml.Linq.XDocument.Load(stream);
        System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var cell = xml.Descendants(ns + "c").Single(item => (string?)item.Attribute("r") == address);
        var strings = cell.Element(ns + "is")!.Elements().SelectMany(item => item.Name == ns + "t"
            ? new[] { item } : item.Name == ns + "r" ? item.Elements(ns + "t") : []);
        return string.Concat(strings.Select(text => System.Text.RegularExpressions.Regex.Replace(text.Value,
            "_x([0-9a-fA-F]{4})_", match => ((char)ushort.Parse(match.Groups[1].Value,
                System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture)).ToString())));
    }

    internal byte[] RewritePart(byte[] source, string path, Func<string, string> rewrite)
    {
        using var buffer = new MemoryStream();
        buffer.Write(source);
        buffer.Position = 0;
        using (var zip = new System.IO.Compression.ZipArchive(buffer, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(path) ?? throw new InvalidOperationException("Fixture part not found: " + path);
            string xml;
            using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
            entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry(path).Open());
            writer.Write(rewrite(xml));
        }
        return buffer.ToArray();
    }

    internal MergeResult Merge(byte[] source, string json = "{\"Name\":\"Ada\"}", MergeOptions? options = null)
    {
        using var services = Services();
        using var input = new MemoryStream(source);
        using var data = JsonDocument.Parse(json);
        return services.GetRequiredService<MergeEngine>().Merge(input, data, options);
    }
}
