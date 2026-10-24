using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using MonoReport.Models;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012NestedTests
{
    [Fact]
    public void Stage3_Three_level_branching_formulas_relocate_by_parent_and_shared_sibling_tuple_indices()
    {
        var source = Stage3BranchingTemplate();
        var rendered = RenderNested(source, Stage3BranchingData(2, 2));

        using var package = SpreadsheetDocument.Open(new MemoryStream(rendered.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal("A1:AH6", worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!, StringComparer.Ordinal);
        var formulas = cells.Values.Where(cell => cell.CellFormula is not null)
            .ToDictionary(cell => cell.CellReference!.Value!, cell => cell.CellFormula!.Text, StringComparer.Ordinal);
        Assert.Equal(new[] { "C3:R3", "S3:AH3" }, worksheet.Descendants<MergeCell>()
            .Select(merge => merge.Reference!.Value!).OrderBy(reference => reference, StringComparer.Ordinal));

        Assert.Equal(new Dictionary<string, string>
        {
            ["C1"] = "C3",
            ["S1"] = "S3",
            ["D1"] = "K$4+C$3+$A$2",
            ["G1"] = "N$4+C$3+$A$2",
            ["T1"] = "AA$4+S$3+$A$2",
            ["W1"] = "AD$4+S$3+$A$2",
            ["K1"] = "D$4+C$3+$A$2",
            ["N1"] = "G$4+C$3+$A$2",
            ["AA1"] = "T$4+S$3+$A$2",
            ["AD1"] = "W$4+S$3+$A$2",
            ["E1"] = "L$5+E$5+D$4+C$3+$A$2",
            ["F1"] = "M$5+F$5+D$4+C$3+$A$2",
            ["H1"] = "O$5+H$5+G$4+C$3+$A$2",
            ["I1"] = "P$5+I$5+G$4+C$3+$A$2",
            ["U1"] = "AB$5+U$5+T$4+S$3+$A$2",
            ["V1"] = "AC$5+V$5+T$4+S$3+$A$2",
            ["X1"] = "AE$5+X$5+W$4+S$3+$A$2",
            ["Y1"] = "AF$5+Y$5+W$4+S$3+$A$2"
        }, formulas);

        Assert.Equal("Root 0", ReadText(cells["C3"]));
        Assert.Equal("Root 1", ReadText(cells["S3"]));
        Assert.Equal("Branch 0", ReadText(cells["D4"]));
        Assert.Equal("Branch 1", ReadText(cells["G4"]));
        Assert.Equal("Branch 0", ReadText(cells["T4"]));
        Assert.Equal("Branch 1", ReadText(cells["W4"]));
        Assert.Equal("Leaf 0", ReadText(cells["E5"]));
        Assert.Equal("Leaf 1", ReadText(cells["F5"]));
        Assert.Equal("Leaf 0", ReadText(cells["L5"]));
        Assert.Equal("Leaf 1", ReadText(cells["M5"]));
        Assert.Equal("Leaf 0", ReadText(cells["U5"]));
        Assert.Equal("Leaf 1", ReadText(cells["V5"]));
        Assert.Equal("Leaf 0", ReadText(cells["AB5"]));
        Assert.Equal("Leaf 1", ReadText(cells["AC5"]));
        Assert.Equal("7", ReadText(cells["A2"]));
        Assert.Equal("C3", formulas["C1"]);
        Assert.Equal("S3", formulas["S1"]);
        Assert.Empty(rendered.Context.Errors);
    }

    [Theory]
    [InlineData(0, 2, 4, "C3", "G3")]
    [InlineData(1, 1, 8, "C3", "K3")]
    [InlineData(2, 2, 16, "C3", "S3")]
    [InlineData(2, 0, 8, "C3", "K3")]
    public void Stage3_Parent_formula_survives_zero_one_or_multiple_nested_collection_items(
        int branches, int leaves, int expectedParentWidth, string firstFormula, string secondFormula)
    {
        var rendered = RenderNested(Stage3BranchingTemplate(), Stage3BranchingData(branches, leaves));

        using var package = SpreadsheetDocument.Open(new MemoryStream(rendered.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!, StringComparer.Ordinal);
        Assert.Equal(2 + branches * 4 + branches * leaves * 2,
            worksheet.Descendants<Cell>().Count(cell => cell.CellFormula is not null));
        Assert.Equal(expectedParentWidth * 2 + 2, ColumnNumber(worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value!.Split(':')[1].TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9')));
        Assert.Equal(firstFormula, cells["C1"].CellFormula!.Text);
        var secondParentColumn = ColumnName(3 + expectedParentWidth);
        Assert.Equal(secondFormula, cells[secondParentColumn + "1"].CellFormula!.Text);
        Assert.Equal(new[] { "Root 0", "Root 1" }, new[] { "C3", secondParentColumn + "3" }.Select(address => ReadText(cells[address])));
        Assert.Empty(rendered.Context.Errors);
    }

    [Fact]
    public void Stage3_Absolute_reference_to_a_repeated_sibling_target_is_rejected_during_finalization()
    {
        var source = Stage3BranchingTemplate(absoluteSiblingTarget: true);
        var original = source.ToArray();

        var error = Assert.Throws<TemplateException>(() => RenderNested(source, Stage3BranchingData(2, 2)));

        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
        Assert.Contains("F5", error.Location);
        Assert.Contains("empty, discarded, disjoint or ambiguous image", error.Message);
        Assert.Contains("retained contiguous A1 range", error.Message);
        Assert.Equal(original, source);
    }

    private byte[] Stage3BranchingTemplate(bool absoluteSiblingTarget = false) => _fixture.Create((workbook, part) =>
    {
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$5:$K$10";
        var rows = part.Worksheet!.GetFirstChild<SheetData>()!;
        rows.RemoveAllChildren();
        rows.Append(
            new Row(Enumerable.Range(4, 8).Select(column => _fixture.Text(ColumnName(column) + "1", "$h(10)"))) { RowIndex = 1 },
            new Row(new[] { 5, 6, 8, 9 }.Select(column => _fixture.Text(ColumnName(column) + "2", "$h(20)"))) { RowIndex = 2 },
            new Row(new[] { 6, 9 }.Select(column => _fixture.Text(ColumnName(column) + "3", "$h(30)"))) { RowIndex = 3 },
            new Row(new Cell { CellReference = "D5", CellFormula = new CellFormula("D7") },
                new Cell { CellReference = "E5", CellFormula = new CellFormula("H$8+D$7+$B$6") },
                new Cell { CellReference = "F5", CellFormula = new CellFormula(absoluteSiblingTarget
                    ? "$I9" : "I$9+F$9+E$8+D$7+$B$6") },
                new Cell { CellReference = "H5", CellFormula = new CellFormula("E$8+D$7+$B$6") }) { RowIndex = 5 },
            new Row(_fixture.Text("B6", "7")) { RowIndex = 6 },
            new Row(_fixture.Text("A7", "$header(10).Roots"), _fixture.Text("D7", "{Name}")) { RowIndex = 7 },
            new Row(_fixture.Text("A8", "$h(20).Branches"), _fixture.Text("E8", "{Name}"), _fixture.Text("H8", "{Name}")) { RowIndex = 8 },
            new Row(_fixture.Text("A9", "$h(30).Leaves"), _fixture.Text("F9", "{Name}"), _fixture.Text("I9", "{Name}")) { RowIndex = 9 });
        part.Worksheet.Append(new MergeCells(new MergeCell { Reference = "D7:K7" }));
    });

    private static string Stage3BranchingData(int branches, int leaves) => System.Text.Json.JsonSerializer.Serialize(new
    {
        Roots = new[] { new { Name = "Root 0" }, new { Name = "Root 1" } },
        Branches = Enumerable.Range(0, branches).Select(index => new { Name = "Branch " + index }),
        Leaves = Enumerable.Range(0, leaves).Select(index => new { Name = "Leaf " + index })
    });

    private static string ReadText(Cell cell) => cell.InlineString!.InnerText;

    private static int ColumnNumber(string name) => name.Aggregate(0, (value, character) => checked(value * 26 + character - 'A' + 1));
}
