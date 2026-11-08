using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Cli.Tests;
using MonoReport.Tests.Xlsx;
using Xunit;
using Header = DocumentFormat.OpenXml.Wordprocessing.Header;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace MonoReport.Tests.Integration;

// Shared inputs for Core/CLI/HTTP parity, using warnings produced by default options.
internal sealed class TransportFixture
{
    internal const string Data = "{\"Items\":[{\"$level\":10,\"Name\":\"Ada\"}]}";
    internal const string SanitizationData = "{\"Name\":\"a\\u0001Café 😀\\uFFFE\"}";
    private readonly XlsxFixture _xlsx = new();

    internal byte[] Create(bool spreadsheet, int warningCount)
    {
        if (!spreadsheet)
        {
            var paragraphs = new List<XElement> {DocxFixture.P(DocxFixture.Run("Ada"))};
            foreach (var name in new[] {"First", "Second"}.Take(warningCount))
                paragraphs.AddRange([
                    DocxFixture.P(DocxFixture.Run("{For:" + name + "}")), DocxFixture.P(DocxFixture.Run("{.}")),
                    DocxFixture.P(DocxFixture.Run("{/For}"))
                ]);
            var bytes = DocxFixture.Create(paragraphs.ToArray());
            if (warningCount < 3) return bytes;
            using var buffer = new MemoryStream();
            buffer.Write(bytes);
            buffer.Position = 0;
            using (var doc = WordprocessingDocument.Open(buffer, true))
            {
                var main = doc.MainDocumentPart!;
                var header = main.AddNewPart<HeaderPart>();
                header.Header = new Header(
                    new Paragraph(new Run(new Text("{For:Third}"))),
                    new Paragraph(new Run(new Text("{/For}"))));
                main.Document!.Body!.Append(new SectionProperties(
                    new HeaderReference {Id = main.GetIdOfPart(header), Type = HeaderFooterValues.Default}));
            }

            return buffer.ToArray();
        }

        return _xlsx.Create((workbook, part) =>
        {
            var row = part.Worksheet!.Descendants<Row>().Single();
            row.RemoveAllChildren();
            row.Append(_xlsx.Text("A3", "$l(10).Items"), _xlsx.Text("B3", warningCount > 0 ? "{../First}" : "{Name}"));
            if (warningCount > 1) row.Append(_xlsx.Text("C3", "{../../Second}"));
            if (warningCount < 3) return;
            var second = workbook.AddNewPart<WorksheetPart>();
            second.Worksheet =
                new Worksheet(new SheetData(new Row(_xlsx.Text("A1", "$l(10).Items"), _xlsx.Text("B1", "{../Third}"))
                    {RowIndex = 1}));
            workbook.Workbook!.Sheets!.Append(new Sheet
                {Name = "Second", SheetId = 2, Id = workbook.GetIdOfPart(second)});
            workbook.Workbook.DefinedNames!.Append(new DefinedName("'Second'!$B$1")
                {Name = "_xlnm.Print_Area", LocalSheetId = 1});
        });
    }

    internal byte[] CreateSanitization(bool spreadsheet)
    {
        if (!spreadsheet) return DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Name upper}")));
        return _xlsx.Create((_, part) =>
        {
            var row = part.Worksheet!.Descendants<Row>().Single();
            row.RemoveAllChildren();
            row.Append(_xlsx.Text("B3", "{Name upper}"));
        });
    }

    internal GenerationResult Generate(byte[] source)
    {
        return Generate(source, Data);
    }

    internal GenerationResult Generate(byte[] source, string json)
    {
        using var services = _xlsx.Services();
        using var input = new MemoryStream(source);
        using var data = JsonDocument.Parse(json);
        return services.GetRequiredService<DocumentGenerator>().Generate(input, data);
    }

    internal void AssertPackage(byte[] bytes, bool spreadsheet)
    {
        if (spreadsheet)
        {
            using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        else
        {
            DocxFixture.AssertValid(bytes);
        }
    }
}
