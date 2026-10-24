using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

// Shared inputs for Core/CLI/HTTP parity, using warnings produced by default options.
internal sealed class TransportFixture
{
    internal const string Data = "{\"Items\":[{\"$level\":10,\"Name\":\"Ada\"}]}";
    private readonly XlsxFixture _xlsx = new();

    internal byte[] Create(bool spreadsheet, int warningCount)
    {
        if (!spreadsheet)
        {
            var paragraphs = new List<System.Xml.Linq.XElement> { Docx.P(Docx.Run("Ada")) };
            foreach (var name in new[] { "First", "Second" }.Take(warningCount))
                paragraphs.AddRange([Docx.P(Docx.Run("{For:" + name + "}")), Docx.P(Docx.Run("{.}")), Docx.P(Docx.Run("{/For}"))]);
            var bytes = Docx.Create(paragraphs.ToArray());
            if (warningCount < 3) return bytes;
            using var buffer = new MemoryStream(); buffer.Write(bytes); buffer.Position = 0;
            using (var doc = WordprocessingDocument.Open(buffer, true))
            {
                var main = doc.MainDocumentPart!;
                var header = main.AddNewPart<HeaderPart>();
                header.Header = new DocumentFormat.OpenXml.Wordprocessing.Header(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("{For:Third}"))),
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("{/For}"))));
                main.Document!.Body!.Append(new DocumentFormat.OpenXml.Wordprocessing.SectionProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.HeaderReference { Id = main.GetIdOfPart(header), Type = DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default }));
            }
            return buffer.ToArray();
        }
        return _xlsx.Create((workbook, part) =>
        {
            var row = part.Worksheet!.Descendants<Row>().Single(); row.RemoveAllChildren();
            row.Append(_xlsx.Text("A3", "$l(10).Items"), _xlsx.Text("B3", warningCount > 0 ? "{../First}" : "{Name}"));
            if (warningCount > 1) row.Append(_xlsx.Text("C3", "{../../Second}"));
            if (warningCount < 3) return;
            var second = workbook.AddNewPart<WorksheetPart>();
            second.Worksheet = new Worksheet(new SheetData(new Row(_xlsx.Text("A1", "$l(10).Items"), _xlsx.Text("B1", "{../Third}")) { RowIndex = 1 }));
            workbook.Workbook!.Sheets!.Append(new Sheet { Name = "Second", SheetId = 2, Id = workbook.GetIdOfPart(second) });
            workbook.Workbook.DefinedNames!.Append(new DefinedName("'Second'!$B$1") { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
        });
    }

    internal MergeResult Merge(byte[] source)
    {
        using var services = _xlsx.Services(); using var input = new MemoryStream(source); using var data = JsonDocument.Parse(Data);
        return services.GetRequiredService<MergeEngine>().Merge(input, data);
    }

    internal void AssertPackage(byte[] bytes, bool spreadsheet)
    {
        if (spreadsheet)
        {
            using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        else Docx.AssertValid(bytes);
    }
}
