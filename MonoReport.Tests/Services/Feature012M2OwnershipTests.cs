using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature012M2OwnershipTests
{
    private readonly XlsxFixture _fixture = new();

    [Theory]
    [InlineData(0)] [InlineData(1)]
    public void F5_Empty_or_single_parent_keeps_independent_rows_without_orphan_merge_content(int parents)
    {
        var result = _fixture.Merge(Template(bodyMerge: true), Data(3, parents));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var sheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal(3 * parents, sheet.Descendants<MergeCell>().Count());
        var cells = Cells(package);
        for (var logical = 0; logical < 2; logical++)
        {
            var row = 4 + logical;
            Assert.Equal("descriptor" + logical, cells[Address(row, 1)].InlineString!.InnerText);
            Assert.Equal("descriptor" + logical, cells[Address(row, 3 + parents * 4)].InlineString!.InnerText);
            Assert.Equal("900", cells[Address(row, 4 + parents * 4)].CellValue!.Text);
            if (parents == 0) continue;
            Assert.Equal("descriptor" + logical, cells["C" + row].InlineString!.InnerText);
            AssertBlank(cells["D" + row]); AssertBlank(cells["E" + row]); AssertBlank(cells["F" + row]);
        }
    }

    [Theory]
    [InlineData(0, false)] [InlineData(1, false)] [InlineData(3, false)]
    [InlineData(0, true)] [InlineData(1, true)] [InlineData(3, true)]
    public void F4_F5_Body_owner_references_and_independent_leaf_cells_keep_their_distinct_sources(int children, bool formulaAnchor)
    {
        var source = Template(bodyMerge: true, followRow: true, formulaAnchor: formulaAnchor);
        var result = _fixture.Merge(source, Data(children));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var cells = Cells(package);
        for (var logical = 0; logical < 2; logical++)
            for (var parent = 0; parent < 2; parent++)
            {
                var row = 4 + 2 * logical; var first = 3 + parent * (children + 1);
                var anchor = cells[Address(row, first)];
                if (formulaAnchor) Assert.Equal("LEN(A" + row + ")", anchor.CellFormula!.Text);
                else Assert.Equal("descriptor" + logical, anchor.InlineString!.InnerText);
                for (var column = first + 1; column <= first + children; column++) AssertBlank(cells[Address(row, column)]);
                Assert.Equal(Address(row, first), cells[Address(row + 1, first + children)].CellFormula!.Text);
                // The next source row is outside this merge: distinct tuple data must never be suppressed.
                for (var child = 0; child < children; child++)
                    Assert.Equal((100 * logical + parent * children + child).ToString(), cells[Address(row + 1, first + child)].CellValue!.Text);
                Assert.Equal("descriptor" + logical, cells[Address(row, 3 + 2 * (children + 1))].InlineString!.InnerText);
            }
    }

    [Theory]
    [InlineData(0, "E7*2")] [InlineData(1, "E7*2")] [InlineData(3, "E7*2")]
    [InlineData(0, "SUM(E7:E7)")] [InlineData(1, "SUM(E7:E7)")] [InlineData(3, "SUM(E7:E7)")]
    [InlineData(0, "COUNTBLANK(E7:E7)")] [InlineData(1, "COUNTBLANK(E7:E7)")] [InlineData(3, "COUNTBLANK(E7:E7)")]
    public void F4_Blank_nonanchor_references_never_alias_the_relocated_rate(int children, string formula)
    {
        var source = Template(formula: formula);
        if (children == 0)
        {
            AssertUnsafe(source, Data(children), "E10");
            return;
        }
        var result = _fixture.Merge(source, Data(children));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var cells = Cells(package);
        for (var parent = 0; parent < 2; parent++)
        {
            var anchor = Address(1, 3 + parent * (children + 1));
            var blank = Address(1, 3 + children + parent * (children + 1));
            Assert.Equal(parent == 0 ? "10" : "20", cells[anchor].CellValue!.Text);
            AssertBlank(cells[blank]);
            for (var row = 4; row <= 5; row++)
            {
                var actual = cells[Address(row, 3 + children + parent * (children + 1))].CellFormula!.Text;
                Assert.Equal(formula.Replace("E7", blank, StringComparison.Ordinal), actual);
            }
        }
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(3)]
    public void F4_Anchor_reference_follows_its_owner_across_parents_rows_and_child_counts(int children)
    {
        var result = _fixture.Merge(Template(formula: "D7*2+SUM(D7:D7)"), Data(children));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        var cells = Cells(package);
        for (var parent = 0; parent < 2; parent++)
        {
            var anchor = Address(1, 3 + parent * (children + 1));
            Assert.Equal(parent == 0 ? "10" : "20", cells[anchor].CellValue!.Text);
            for (var row = 4; row <= 5; row++)
                Assert.Equal(anchor + "*2+SUM(" + anchor + ":" + anchor + ")", cells[Address(row, 3 + children + parent * (children + 1))].CellFormula!.Text);
        }
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(3)]
    public void F4_Range_including_distinct_anchor_and_blank_requires_a_faithful_image(int children)
    {
        var source = Template(formula: "COUNTBLANK(D7:E7)");
        if (children != 1) { AssertUnsafe(source, Data(children), "E10"); return; }
        var result = _fixture.Merge(source, Data(children));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        var cells = Cells(package);
        Assert.Equal("COUNTBLANK(C1:D1)", cells["D4"].CellFormula!.Text);
        Assert.Equal("COUNTBLANK(E1:F1)", cells["F5"].CellFormula!.Text);
        AssertBlank(cells["D1"]); AssertBlank(cells["F1"]);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void F4_Displaced_source_identity_exists_even_without_an_authored_nonanchor_cell(bool omitNonanchor)
    {
        var source = Template(formula: "E7*2", omitNonanchor: omitNonanchor);
        AssertUnsafe(source, Data(0), "E10");
    }

    [Theory]
    [InlineData(0, true)] [InlineData(1, true)] [InlineData(3, true)]
    [InlineData(0, false)] [InlineData(1, false)] [InlineData(3, false)]
    public void F5_Each_body_merge_has_one_content_owner_preserving_nonanchor_styles_and_independent_fields(int children, bool childFirst)
    {
        var source = Template(bodyMerge: true, childFirst: childFirst);
        var result = _fixture.Merge(source, Data(children));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var cells = Cells(package);
        var sheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal(children == 0 ? 0 : 6, sheet.Descendants<MergeCell>().Count()); // Two headings plus four body instances.
        for (var logical = 0; logical < 2; logical++)
        {
            var row = logical + 4;
            Assert.Equal("descriptor" + logical, cells[Address(row, 1)].InlineString!.InnerText);
            for (var parent = 0; parent < 2; parent++)
            {
                var first = 3 + parent * (children + 1);
                Assert.Equal("descriptor" + logical, cells[Address(row, first)].InlineString!.InnerText);
                Assert.Equal(1U, cells[Address(row, first)].StyleIndex!.Value);
                for (var column = first + 1; column <= first + children; column++)
                {
                    AssertBlank(cells[Address(row, column)]);
                    Assert.Equal(!childFirst || column == first + children ? 2U : 1U, cells[Address(row, column)].StyleIndex!.Value);
                }
            }
            var outside = 3 + 2 * (children + 1);
            Assert.Equal("descriptor" + logical, cells[Address(row, outside)].InlineString!.InnerText);
            Assert.Equal("900", cells[Address(row, outside + 1)].CellValue!.Text);
        }
    }

    [Theory]
    [InlineData("text")] [InlineData("formula")] [InlineData("number")]
    public void F5_Authored_meaningful_nonanchor_content_is_rejected_in_both_interfaces(string kind)
    {
        var source = Template(bodyMerge: true, nonanchorContent: kind);
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        var issue = Assert.Single(analysis.Errors); Assert.Contains("nonanchor", issue.Message);
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source, Data(3)));
        Assert.Equal(issue.Code, error.Code); Assert.Equal(issue.Location, error.Errors[0].Location); Assert.Equal(issue.Message, error.Errors[0].Message);
    }

    private void AssertUnsafe(byte[] source, string data, string address)
    {
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        // These ownership errors depend on counts; Analyze must not invent JSON cardinalities.
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, data));
        Assert.Null(result); Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
        Assert.Contains("sheet 'Report'", error.Location); Assert.Contains(address, error.Location);
        Assert.Contains("Re-author", error.Errors[0].Message);
    }

    private static Dictionary<string, Cell> Cells(SpreadsheetDocument package) => package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
    private static string Address(int row, int column) => ((char)('A' + column - 1)).ToString() + row;
    private static void AssertBlank(Cell cell) { Assert.Null(cell.CellFormula); Assert.Null(cell.CellValue); Assert.Null(cell.InlineString); }

    private byte[] Template(bool bodyMerge = false, bool childFirst = true, string? formula = null, bool omitNonanchor = false,
        string? nonanchorContent = null, bool followRow = false, bool formulaAnchor = false) => _fixture.Create((workbook, part) =>
    {
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$7:$G$12";
        workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font(), new Font(new Bold())),
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
            new Borders(new Border(), new Border(new BottomBorder { Style = BorderStyleValues.Thin })), new CellStyleFormats(new CellFormat()),
            new CellFormats(new CellFormat(), new CellFormat { FontId = 1, ApplyFont = true }, new CellFormat { BorderId = 1, ApplyBorder = true }));
        var header = _fixture.Text("D7", "{Rate}"); header.StyleIndex = 1;
        var bodyAnchor = _fixture.Text("D10", bodyMerge ? "{Name}" : "{Price}"); bodyAnchor.StyleIndex = 1;
        if (formulaAnchor) { bodyAnchor.RemoveAllChildren(); bodyAnchor.DataType = null; bodyAnchor.CellFormula = new CellFormula("LEN(B10)"); }
        var nonanchor = new Cell { CellReference = "E10", StyleIndex = 2 };
        if (nonanchorContent == "text") { nonanchor.DataType = CellValues.InlineString; nonanchor.InlineString = new InlineString(new Text("independent")); }
        if (nonanchorContent == "formula") nonanchor.CellFormula = new CellFormula("1+1");
        if (nonanchorContent == "number") nonanchor.CellValue = new CellValue("77");
        if (!bodyMerge) { nonanchor.StyleIndex = null; nonanchor.CellFormula = new CellFormula(formula ?? "D7*2"); }
        part.Worksheet = new Worksheet(new SheetData(
            new Row(_fixture.Text("D1", "$h(1)"), _fixture.Text("E1", "$h(1)")) { RowIndex = 1 },
            new Row(_fixture.Text(childFirst ? "D2" : "E2", "$h(2)")) { RowIndex = 2 },
            new Row(new[] { _fixture.Text("A7", "$h(1).Bids"), header, omitNonanchor ? null : new Cell { CellReference = "E7", StyleIndex = 2 } }.OfType<Cell>()) { RowIndex = 7 },
            new Row(_fixture.Text("A8", "$h(2).Criteria"), _fixture.Text(childFirst ? "D8" : "E8", "{Name}")) { RowIndex = 8 },
            new Row(_fixture.Text("A10", "$l(10).Items"), _fixture.Text("B10", "{Name}"), bodyAnchor, nonanchor,
                _fixture.Text("F10", "{Name}"), _fixture.Text("G10", "{$.Independent}")) { RowIndex = 10 }),
            new MergeCells(bodyMerge ? new[] { new MergeCell { Reference = "D7:E7" }, new MergeCell { Reference = "D10:E10" } } : new[] { new MergeCell { Reference = "D7:E7" } }));
        if (followRow) part.Worksheet.GetFirstChild<SheetData>()!.Append(new Row(_fixture.Text("A11", "$level(10).Items"),
            _fixture.Text("D11", "{Price}"), new Cell { CellReference = "E11", CellFormula = new CellFormula("D10") }) { RowIndex = 11 });
    });

    private static string Data(int children, int parents = 2) => JsonSerializer.Serialize(new
    {
        Independent = 900,
        Bids = Enumerable.Range(1, parents).Select(parent => new { Rate = 10 * parent }),
        Criteria = Enumerable.Range(0, children).Select(i => new { Name = "criterion" + i }),
        Items = Enumerable.Range(0, 2).SelectMany(row => Enumerable.Range(0, Math.Max(1, parents * children)).Select(tuple =>
            new Dictionary<string, object> { ["$level"] = 10, ["$path"] = "row" + row, ["Name"] = tuple == 0 ? "descriptor" + row : "row" + row + "-tuple" + tuple, ["Price"] = 100 * row + tuple }))
    });
}
