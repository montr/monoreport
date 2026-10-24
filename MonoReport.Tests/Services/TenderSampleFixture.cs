using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace MonoReport.Tests.Services;

internal sealed class TenderSampleFixture
{
    internal string TemplatePath => Path.Combine(AppContext.BaseDirectory, "Samples/xlsx-tender.template.xlsx");
    internal byte[] Source => File.ReadAllBytes(TemplatePath);
    internal string Data => File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Samples/xlsx-tender.json"));

    internal void AssertGenerated(byte[] bytes)
    {
        using var document = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var workbook = document.WorkbookPart!;
        var sheet = Assert.Single(workbook.WorksheetParts).Worksheet!;
        var cells = sheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
        Assert.Equal("A1:I9", sheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal("Tender: Office equipment", cells["A1"].InlineString!.InnerText);
        Assert.Equal("Alpha", cells["C3"].InlineString!.InnerText);
        Assert.Equal("Beta", cells["F3"].InlineString!.InnerText);
        Assert.Equal("Unit price", cells["D4"].InlineString!.InnerText);
        Assert.Equal("Lead days", cells["E4"].InlineString!.InnerText);
        Assert.Equal("IT equipment", cells["A5"].InlineString!.InnerText);
        Assert.Equal("Laptop / IT equipment", cells["A6"].InlineString!.InnerText);
        Assert.Equal("Dock / IT equipment", cells["A7"].InlineString!.InnerText);
        Assert.Equal("2", cells["B6"].CellValue!.Text); Assert.Equal("3", cells["B7"].CellValue!.Text);
        foreach (var (address, value) in new[] { ("D6", "1200"), ("E6", "10"), ("G6", "1150"), ("H6", "12"), ("D7", "80"), ("E7", "5"), ("G7", "75"), ("H7", "7") })
        {
            Assert.Equal(value, cells[address].CellValue!.Text); Assert.Null(cells[address].InlineString);
        }
        Assert.Equal("Lot total: IT equipment", cells["A8"].InlineString!.InnerText);
        foreach (var column in new[] { "D", "E", "G", "H" })
        {
            Assert.Equal($"SUM({column}6:{column}7)", cells[column + "8"].CellFormula!.Text);
            Assert.Null(cells[column + "8"].CellValue);
        }
        Assert.Equal(new[] { "C3:E3", "F3:H3" }, sheet.Descendants<MergeCell>().Select(m => m.Reference!.Value));
        Assert.Equal("'Tender'!$A$1:$I$9", Assert.Single(workbook.Workbook!.DefinedNames!.Elements<DefinedName>()).Text);
        Assert.True(workbook.Workbook!.CalculationProperties!.FullCalculationOnLoad!.Value);
        Assert.DoesNotContain(sheet.Descendants<Text>(), t => t.Text!.Contains("$h(", StringComparison.Ordinal) || t.Text!.Contains("$l(", StringComparison.Ordinal));
    }
}
