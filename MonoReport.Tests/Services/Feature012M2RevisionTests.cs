using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature012M2RevisionTests
{
    private readonly XlsxFixture _fixture = new();

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void F3_Static_preflight_preserves_one_error_per_formula_and_existing_operand_order(bool repeated)
    {
        var source = _fixture.Create((_, part) =>
        {
            var cell = part.Worksheet!.Descendants<Cell>().Single(); cell.RemoveAllChildren(); cell.DataType = null; cell.CellFormula = new CellFormula("A1+A2");
            if (repeated) part.Worksheet.GetFirstChild<SheetData>()!.Append(new Row(_fixture.Text("A4", "$l(10).Items"), _fixture.Text("B4", "{Name}")) { RowIndex = 4 });
        });
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var issue = Assert.Single(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Errors);
        Assert.Contains("Reference A2", issue.Message);
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source, "{\"Items\":[]}"));
        Assert.Single(error.Errors); Assert.Equal(issue.Message, error.Errors[0].Message);
    }

    [Theory]
    [InlineData(0, "C4", "D4", "C1", "D1")]
    [InlineData(1, "D4", "F4", "C1", "E1")]
    [InlineData(3, "F4", "J4", "C1", "G1")]
    public void F1_F2_Formulas_in_trailing_own_columns_follow_the_retained_merge_anchor(int criteria, string alphaCell, string betaCell, string alphaAnchor, string betaAnchor)
    {
        var source = HeaderTemplate(true, true);
        source = _fixture.RewritePart(source, "xl/worksheets/sheet1.xml", text =>
        {
            var xml = System.Xml.Linq.XDocument.Parse(text); System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var cell = xml.Descendants(ns + "c").Single(cell => (string?)cell.Attribute("r") == "E10");
            cell.RemoveNodes(); cell.Attribute("t")?.Remove(); cell.Add(new System.Xml.Linq.XElement(ns + "f", "LEN(D7)"));
            return xml.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        });
        var result = _fixture.Merge(source, Data(criteria));
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        var cells = output.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("LEN(" + alphaAnchor + ")", cells[alphaCell].CellFormula!.Text);
        Assert.Equal("LEN(" + betaAnchor + ")", cells[betaCell].CellFormula!.Text);
        Assert.Equal("Alpha", _fixture.ReadCellText(result.Document, alphaAnchor)); Assert.Equal("Beta", _fixture.ReadCellText(result.Document, betaAnchor));
    }

    [Theory]
    [InlineData(0, false)] [InlineData(0, true)] [InlineData(1, false)] [InlineData(3, false)]
    public void F1_Parent_merge_keeps_anchor_content_style_and_trailing_own_column_when_child_is_empty(int criteria, bool hidden)
    {
        var source = HeaderTemplate(childFirst: true, merged: true, hidden: hidden);
        var result = _fixture.Merge(source, Data(criteria));
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(output));
        var sheet = output.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        var cells = sheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        var beta = criteria switch { 0 => "D1", 1 => "E1", _ => "G1" };
        Assert.NotNull(cells["C1"].InlineString); Assert.NotNull(cells[beta].InlineString);
        Assert.Equal("Alpha", _fixture.ReadCellText(result.Document, "C1"));
        Assert.Equal("Beta", _fixture.ReadCellText(result.Document, beta));
        Assert.Equal(1U, cells["C1"].StyleIndex!.Value); Assert.Equal(1U, cells[beta].StyleIndex!.Value);
        Assert.Equal("descriptor", _fixture.ReadCellText(result.Document, "A4"));
        var columns = sheet.Descendants<Column>().Where(column => column.Width?.Value == 0).ToArray();
        Assert.Equal(new[] { (uint)(3 + criteria), (uint)(4 + 2 * criteria) }, columns.Select(column => column.Min!.Value));
        Assert.All(columns, column => Assert.Equal(hidden, column.Hidden?.Value ?? false));
        Assert.Equal(criteria == 0 ? 0 : 2, sheet.Descendants<MergeCell>().Count());
        if (criteria > 0)
            Assert.Equal(new[] { criteria == 1 ? "C1:D1" : "C1:F1", criteria == 1 ? "E1:F1" : "G1:J1" }, sheet.Descendants<MergeCell>().Select(merge => merge.Reference!.Value));
    }

    [Theory]
    [InlineData("E7*2", "D1*2", "H1*2")]
    [InlineData("SUM(E7:E7)", "SUM(D1:D1)", "SUM(H1:H1)")]
    [InlineData("SUM(D7:E7)", "SUM(C1:D1)", "SUM(G1:H1)")]
    public void F2_Leaf_formulas_bind_to_retained_parent_header_content_for_cells_and_ranges(string formula, string alpha, string beta)
    {
        var source = HeaderTemplate(childFirst: false, merged: false, formula: formula);
        var result = _fixture.Merge(source, Data(3));
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(output));
        var cells = output.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("10", cells["D1"].CellValue!.Text); Assert.Equal("20", cells["H1"].CellValue!.Text);
        Assert.All(new[] { "E1", "F1", "I1", "J1" }, address => Assert.Null(cells[address].CellValue));
        Assert.Equal(new[] { alpha, alpha, alpha, beta, beta, beta }, new[] { "D4", "E4", "F4", "H4", "I4", "J4" }.Select(address => cells[address].CellFormula!.Text));
    }

    [Theory]
    [InlineData("$E$7")]
    [InlineData("SUM(E7:E8)")]
    public void F2_Unsafe_header_owner_images_reject_without_selecting_a_placeholder(string formula)
    {
        var source = HeaderTemplate(false, false, formula);
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source, Data(3)));
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code); Assert.Contains("E10", error.Location);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void F3_Static_discarded_reference_rejects_in_both_interfaces_with_same_sheet_or_unrelated_repetition(bool mixedSheets)
    {
        var source = _fixture.Create((workbook, part) =>
        {
            var cell = part.Worksheet!.Descendants<Cell>().Single(); cell.RemoveAllChildren(); cell.DataType = null; cell.CellFormula = new CellFormula("A1");
            if (!mixedSheets) part.Worksheet.GetFirstChild<SheetData>()!.Append(new Row(_fixture.Text("A4", "$l(10).Items"), _fixture.Text("B4", "{Name}")) { RowIndex = 4 });
            else
            {
                var repeated = workbook.AddNewPart<WorksheetPart>();
                repeated.Worksheet = new Worksheet(new SheetData(new Row(_fixture.Text("A3", "$l(10).Items"), _fixture.Text("B3", "{Name}")) { RowIndex = 3 }));
                workbook.Workbook!.Sheets!.Append(new Sheet { Name = "Repeated", SheetId = 2, Id = workbook.GetIdOfPart(repeated) });
                workbook.Workbook.DefinedNames!.Append(new DefinedName("'Repeated'!$B$3:$D$5") { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
            }
        });
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(analysis.Valid); var issue = Assert.Single(analysis.Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code); Assert.Contains("sheet 'Report'", issue.Location); Assert.Contains("B3", issue.Location);
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, "{\"Items\":[{\"$level\":10,\"Name\":\"row\"}]}"));
        Assert.Null(result); Assert.Equal(issue.Code, error.Code); Assert.Equal(issue.Location, error.Errors[0].Location); Assert.Equal(issue.Message, error.Errors[0].Message);
        Assert.Equal(source, input.ToArray());
    }

    [Fact]
    public void F3_Formula_in_discarded_repeated_instance_is_not_evaluated_by_source_preflight()
    {
        var source = _fixture.Create((_, part) =>
        {
            var row = part.Worksheet!.Descendants<Row>().Single(); row.PrependChild(_fixture.Text("A3", "$l(10).Items"));
            var cell = row.Elements<Cell>().Single(cell => cell.CellReference!.Value == "B3"); cell.RemoveAllChildren(); cell.DataType = null; cell.CellFormula = new CellFormula("A1");
        });
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
        var result = _fixture.Merge(source, "{\"Items\":[]}");
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(output.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<CellFormula>());
    }

    [Fact]
    public void F3_Static_formula_over_repeated_content_keeps_data_dependent_validation_for_generation()
    {
        var source = _fixture.Create((_, part) =>
        {
            var cell = part.Worksheet!.Descendants<Cell>().Single(); cell.RemoveAllChildren(); cell.DataType = null; cell.CellFormula = new CellFormula("SUM(B4:B4)");
            part.Worksheet.GetFirstChild<SheetData>()!.Append(new Row(_fixture.Text("A4", "$l(10).Items"), _fixture.Text("B4", "{Name}")) { RowIndex = 4 });
        });
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
        var result = _fixture.Merge(source, "{\"Items\":[{\"$level\":10,\"Name\":3}]}");
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Equal("SUM(A2:A2)", Assert.Single(output.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<CellFormula>()).Text);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, Assert.Throws<TemplateException>(() => _fixture.Merge(source, "{\"Items\":[]}")).Code);
    }

    private byte[] HeaderTemplate(bool childFirst, bool merged, string? formula = null, bool hidden = false) => _fixture.Create((workbook, part) =>
    {
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$7:$F$12";
        workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font(), new Font(new Bold())),
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
            new Borders(new Border()), new CellStyleFormats(new CellFormat()), new CellFormats(new CellFormat(), new CellFormat { FontId = 1, ApplyFont = true }));
        var anchor = _fixture.Text(childFirst ? "D7" : "E7", childFirst ? "{Company}" : "{Rate}"); anchor.StyleIndex = 1;
        var own = childFirst ? 5U : 4U;
        var body = childFirst ? _fixture.Text("D10", "{Price}") : new Cell { CellReference = "E10", CellFormula = new CellFormula(formula ?? "E7*2") };
        part.Worksheet = new Worksheet(new Columns(new Column { Min = own, Max = own, Width = 0, CustomWidth = true, Hidden = hidden }),
            new SheetData(new Row(_fixture.Text("D1", "$h(1)"), _fixture.Text("E1", "$h(1)")) { RowIndex = 1 },
                new Row(_fixture.Text(childFirst ? "D2" : "E2", "$h(2)")) { RowIndex = 2 },
                new Row(new[] { _fixture.Text("A7", "$h(1).Bids"), anchor, childFirst ? new Cell { CellReference = "E7" } : null }.OfType<Cell>()) { RowIndex = 7 },
                new Row(_fixture.Text("A8", "$h(2).Criteria"), _fixture.Text(childFirst ? "D8" : "E8", "{Name}")) { RowIndex = 8 },
                new Row(new[] { _fixture.Text("A10", "$l(10).Items"), _fixture.Text("B10", "{Name}"), body, childFirst ? _fixture.Text("E10", "{Name}") : null }.OfType<Cell>()) { RowIndex = 10 }));
        if (merged) part.Worksheet.Append(new MergeCells(new MergeCell { Reference = "D7:E7" }));
    });

    private string Data(int criteria) => JsonSerializer.Serialize(new
    {
        Bids = new[] { new { Company = "Alpha", Rate = 10 }, new { Company = "Beta", Rate = 20 } },
        Criteria = Enumerable.Range(0, criteria).Select(index => new { Name = "criterion " + index }),
        Items = Enumerable.Range(0, Math.Max(1, 2 * criteria)).Select(_ => new Dictionary<string, object> { ["$level"] = 10, ["$path"] = "row", ["Name"] = "descriptor", ["Price"] = 100 })
    });
}
