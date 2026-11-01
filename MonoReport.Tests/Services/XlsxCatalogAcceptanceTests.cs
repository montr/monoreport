using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class XlsxCatalogAcceptanceTests
{
    [Theory]
    [InlineData("xlsx-tender")]
    [InlineData("xlsx-employees")]
    public void Catalog_workbooks_analyze_generate_and_preserve_native_types_formulas_and_styles(string id)
    {
        var template = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, id + ".template.xlsx"));
        var json = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, id + ".json"));
        using var services = new XlsxFixture().Services();
        using (var analysisInput = new MemoryStream(template))
        {
            var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(analysisInput);
            Assert.True(analysis.Valid, string.Join(Environment.NewLine, analysis.Errors.Select(error => error.Message)));
        }

        using var input = new MemoryStream(template);
        using var data = JsonDocument.Parse(json);
        var generated = services.GetRequiredService<MergeEngine>().Merge(input, data).Document;
        using var package = SpreadsheetDocument.Open(new MemoryStream(generated), false);
        var workbook = package.WorkbookPart!;
        var sheet = Assert.Single(workbook.WorksheetParts);
        var cells = sheet.Worksheet!.Descendants<Cell>().ToDictionary(
            cell => cell.CellReference?.Value ?? throw new InvalidOperationException("Catalog cell has no reference."), StringComparer.Ordinal);
        var printArea = Assert.Single(workbook.Workbook!.DefinedNames!.Elements<DefinedName>(),
            name => name.Name == "_xlnm.Print_Area");

        if (id == "xlsx-tender")
        {
            using var source = JsonDocument.Parse(json);
            Assert.Equal(2, source.RootElement.GetProperty("Bids").GetArrayLength());
            Assert.Equal(2, source.RootElement.GetProperty("Criteria").GetArrayLength());
            Assert.Equal(12, source.RootElement.GetProperty("Items").GetArrayLength());
            Assert.Equal("'Tender'!$A$1:$I$10", printArea.Text);
            Assert.Equal("Alpha", Text(cells["C3"]));
            Assert.Equal("Beta", Text(cells["F3"]));
            Assert.Equal("Unit price", Text(cells["D4"]));
            Assert.Equal("Unit price", Text(cells["G4"]));
            Assert.Equal("Laptop", Text(cells["C6"]));
            Assert.Equal("Dock", Text(cells["C7"]));
            Assert.Equal("2", Text(cells["B6"]));
            Assert.Equal("1200", Text(cells["D6"]));
            Assert.Equal("1150", Text(cells["G6"]));
            Assert.Equal(CellValues.Number, cells["B6"].DataType?.Value);
            Assert.Equal(CellValues.Number, cells["D6"].DataType?.Value);
            Assert.Equal("18", cells["D6"].StyleIndex?.Value.ToString());
            Assert.Equal("20", cells["D8"].StyleIndex?.Value.ToString());
            Assert.Equal(new[] { "SUM(D6:D7)", "SUM(E6:E7)", "SUM(G6:G7)", "SUM(H6:H7)" },
                sheet.Worksheet.Descendants<CellFormula>().Select(formula => formula.Text));
        }
        else
        {
            using var source = JsonDocument.Parse(json);
            var records = source.RootElement.GetProperty("Data").EnumerateArray().ToArray();
            Assert.Equal(110, records.Length);
            Assert.Equal(10, records.Count(record => record.GetProperty("$level").GetInt32() == 10));
            Assert.Equal(100, records.Count(record => record.GetProperty("$level").GetInt32() == 40));
            Assert.Equal("'Employees'!$A$1:$D$112", printArea.Text);
            Assert.Equal(10, cells.Values.Count(cell => cell.StyleIndex?.Value == 2));
            Assert.Equal(100, cells.Values.Count(cell => cell.StyleIndex?.Value == 6));
            Assert.Equal(100, cells.Values.Count(cell => cell.StyleIndex?.Value == 7));
            Assert.Equal("Engineering", Text(cells["B3"]));
            Assert.Equal("Avery Chen", Text(cells["C4"]));
            Assert.Equal("Software Engineer", Text(cells["D4"]));
            Assert.Equal(CellValues.InlineString, cells["C4"].DataType?.Value);
            Assert.Equal("6", cells["C4"].StyleIndex?.Value.ToString());
            Assert.Empty(sheet.Worksheet.Descendants<CellFormula>());
        }
    }

    private static string Text(Cell cell) => cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? string.Empty;
}
