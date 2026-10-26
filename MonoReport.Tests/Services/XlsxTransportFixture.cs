using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace MonoReport.Tests.Services;

// Shared transport input is generated from test-owned data; no Samples or examples-manifest dependency.
internal sealed class XlsxTransportFixture
{
    internal byte[] Source { get; } = new XlsxFixture().CreateRepeated();
    internal string Data => "{\"Items\":[{\"$level\":10,\"Name\":\"Alpha\",\"Price\":12.5},{\"$level\":10,\"Name\":\"Beta\",\"Price\":20}]}";

    internal void AssertGenerated(byte[] bytes)
    {
        using var document = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var workbook = document.WorkbookPart!;
        var sheet = Assert.Single(workbook.WorksheetParts).Worksheet!;
        var cells = sheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("A1:C4", sheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal("Alpha", cells["A1"].InlineString!.InnerText); Assert.Equal("Beta", cells["A2"].InlineString!.InnerText);
        Assert.Equal("12.5", cells["B1"].CellValue!.Text); Assert.Equal("20", cells["B2"].CellValue!.Text);
        Assert.Equal("B1*2", cells["C1"].CellFormula!.Text); Assert.Equal("B2*2", cells["C2"].CellFormula!.Text);
        Assert.Equal("SUM(B1:B2)", cells["C4"].CellFormula!.Text);
        Assert.Null(cells["C1"].CellValue); Assert.Null(cells["C4"].CellValue);
        Assert.Equal("'Report'!$A$1:$C$4", Assert.Single(workbook.Workbook!.DefinedNames!.Elements<DefinedName>()).Text);
        Assert.True(workbook.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);
    }
}
