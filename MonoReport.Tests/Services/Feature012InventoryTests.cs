using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed class Feature012InventoryTests
{
    private readonly XlsxFixture _fixture = new();
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Theory]
    [InlineData("missing-relationship")]
    [InlineData("duplicate-part")]
    [InlineData("truncated-xml")]
    [InlineData("missing-sheet-data")]
    [InlineData("duplicate-cell")]
    [InlineData("wrong-row")]
    [InlineData("column-overflow")]
    [InlineData("unknown-sheetdata-child")]
    [InlineData("unordered-cells")]
    [InlineData("invalid-date-flag")]
    public void Inventory_Malformed_sources_are_diagnostics_and_never_mutate_input(string scenario)
    {
        var source = _fixture.Create((workbook, _) =>
        {
            var sheet = workbook.Workbook!.Sheets!.Elements<Sheet>().Single();
            if (scenario == "invalid-date-flag") workbook.Workbook.InsertAt(new WorkbookProperties { Date1904 = false }, 0);
            if (scenario == "missing-relationship") sheet.Id = "rId404";
            if (scenario == "duplicate-part")
            {
                workbook.Workbook.Sheets.Append(new Sheet { Name = "Second", SheetId = 2, Id = sheet.Id!.Value });
                workbook.Workbook.DefinedNames!.Append(new DefinedName("'Second'!$B$3:$D$5")
                    { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
            }
        });
        if (scenario == "invalid-date-flag")
            source = _fixture.RewritePart(source, "xl/workbook.xml", xml =>
            {
                var workbook = XDocument.Parse(xml); workbook.Root!.Element(S + "workbookPr")!.SetAttributeValue("date1904", "invalid");
                return workbook.ToString();
            });
        else if (scenario is not ("missing-relationship" or "duplicate-part"))
            source = _fixture.RewritePart(source, "xl/worksheets/sheet1.xml", xml =>
            {
                if (scenario == "truncated-xml") return xml[..^10];
                var document = XDocument.Parse(xml);
                var cell = document.Descendants(S + "c").Single();
                if (scenario == "missing-sheet-data") document.Descendants(S + "sheetData").Remove();
                if (scenario == "duplicate-cell") cell.AddAfterSelf(new XElement(cell));
                if (scenario == "wrong-row") cell.SetAttributeValue("r", "B4");
                if (scenario == "column-overflow") cell.SetAttributeValue("r", "XFE3");
                if (scenario == "unknown-sheetdata-child") document.Descendants(S + "sheetData").Single().Add(new XElement(S + "unknownRow"));
                if (scenario == "unordered-cells") cell.AddAfterSelf(new XElement(S + "c", new XAttribute("r", "A3")));
                return document.ToString();
            });
        using var services = _fixture.Services();
        using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(analysis.Valid);
        var issue = Assert.Single(analysis.Errors);
        Assert.Equal(scenario == "truncated-xml" ? DiagnosticCodes.TemplateInvalidXml : DiagnosticCodes.XlsxUnsupportedFeature, issue.Code);
        Assert.NotNull(issue.Part);
        Assert.DoesNotContain("rId404", issue.Message);
        Assert.Equal(source, input.ToArray());
        Assert.True(input.CanRead);
        input.Position = 0;
        using var data = System.Text.Json.JsonDocument.Parse("{\"Name\":\"Ada\"}");
        var generation = Assert.Throws<TemplateException>(() => services.GetRequiredService<MergeEngine>().Merge(input, data));
        Assert.Equal(issue.Code, generation.Errors[0].Code);
        Assert.Equal(source, input.ToArray());
    }

    [Fact]
    public void Inventory_Phonetic_annotations_do_not_become_visible_text_and_actual_reads_are_counted()
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            var strings = workbook.AddNewPart<SharedStringTablePart>();
            strings.SharedStringTable = new SharedStringTable(new SharedStringItem(new Text("東京 {Name}"),
                new PhoneticRun(new Text("トウキョウ")) { BaseTextStartIndex = 0, EndingBaseIndex = 2 }));
            var cell = worksheet.Worksheet!.Descendants<Cell>().Single();
            cell.RemoveAllChildren(); cell.DataType = CellValues.SharedString; cell.CellValue = new CellValue("0");
        });
        using var input = new MemoryStream(source);
        using var package = SpreadsheetDocument.Open(input, false);
        var context = new XlsxOperationContext(new MergeOptions());
        var inventory = new XlsxPackageReader().Read(package, context);
        var cell = Assert.Single(inventory.Sheets[0].Cells);
        Assert.Equal("東京 {Name}", cell.Text);
        Assert.True(cell.RichText);
        var workbook = package.WorkbookPart!;
        var expectedBytes = new OpenXmlPart[] { workbook, workbook.WorksheetParts.Single(), workbook.SharedStringTablePart! }
            .Sum(part => { using var stream = part.GetStream(FileMode.Open, FileAccess.Read); return stream.Length; });
        Assert.Equal(expectedBytes, context.PartBytesRead);
    }

    [Fact]
    public void Inventory_Missing_target_parts_are_invalid_packages_in_both_operations()
    {
        var source = _fixture.RewritePart(_fixture.Create(), "xl/_rels/workbook.xml.rels",
            xml => xml.Replace("worksheets/sheet1.xml", "worksheets/missing.xml", StringComparison.Ordinal));
        using var services = _fixture.Services();
        using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(analysis.Valid);
        Assert.Equal(DiagnosticCodes.XlsxInvalidPackage, analysis.Errors[0].Code);
        input.Position = 0;
        using var data = System.Text.Json.JsonDocument.Parse("{}");
        var failure = Assert.Throws<TemplateException>(() => services.GetRequiredService<MergeEngine>().Merge(input, data));
        Assert.Equal(DiagnosticCodes.XlsxInvalidPackage, failure.Code);
        Assert.Equal(source, input.ToArray());
    }
}
