using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012NestedTests
{
    [Fact]
    public void Stage1_Three_level_recursive_geometry_renders_expected_width_values_and_sibling_tuple_reuse()
    {
        var source = ThreeLevelTemplate();
        var json = ThreeLevelData(2, 3, 4);

        var rendered = RenderNested(source, json);

        using var package = SpreadsheetDocument.Open(new MemoryStream(rendered.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal("A1:AF15", worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

        // Independent recurrence: 2 * (1 root column + 3 * (1 middle column + 4 leaf columns)) = 32.
        Assert.Equal(new[] { "Group 0", "Group 1" }, new[] { "A1", "Q1" }.Select(address => _fixture.ReadCellText(rendered.Document, address)));
        Assert.Equal(new[] { "Middle 0", "Middle 1", "Middle 2", "Middle 0", "Middle 1", "Middle 2" },
            new[] { "B2", "G2", "L2", "R2", "W2", "AB2" }.Select(address => _fixture.ReadCellText(rendered.Document, address)));
        Assert.Equal(new[] { "Leaf 0", "Leaf 1", "Leaf 2", "Leaf 3", "Leaf 0", "Leaf 1", "Leaf 2", "Leaf 3" },
            new[] { "C3", "D3", "E3", "F3", "X3", "Y3", "Z3", "AA3" }.Select(address => _fixture.ReadCellText(rendered.Document, address)));

        var firstSibling = Enumerable.Range(0, 24).Select(index => (index + 1).ToString()).ToArray();
        var secondSibling = Enumerable.Range(0, 24).Select(index => (index + 101).ToString()).ToArray();
        Assert.Equal(firstSibling, LeafColumns().Select(column => cells[column + "5"].CellValue!.Text));
        Assert.Equal(firstSibling, LeafColumns().Select(column => cells[column + "6"].CellValue!.Text));
        Assert.Equal(secondSibling, LeafColumns().Select(column => cells[column + "7"].CellValue!.Text));
        Assert.Equal(secondSibling, LeafColumns().Select(column => cells[column + "8"].CellValue!.Text));
        Assert.Equal(new[] { "Child 0/Group 0", "Child 1/Group 0", "Child 0/Group 1", "Child 1/Group 1" },
            new[] { "B5", "B7", "R11", "R13" }.Select(address => _fixture.ReadCellText(rendered.Document, address)));
        Assert.All(rendered.Bound[0].Rows[0].Where(row => row.Level == 40), row => Assert.Equal(24, row.Records.Length));
        Assert.Equal(96, rendered.Bound[0].Rows[0].Where(row => row.Level == 40).Sum(row => row.Records.Length));
        Assert.Equal(4, rendered.Bound[0].Rows[0].Count(row => row.Level == 40));
        Assert.Empty(rendered.Context.Errors);
    }

    [Fact]
    public void Stage1_Empty_leaf_product_keeps_parent_columns_and_renders_one_descriptor_per_vertical_item()
    {
        var source = ThreeLevelTemplate();
        var json = ThreeLevelData(2, 3, 0);

        var rendered = RenderNested(source, json);

        using var package = SpreadsheetDocument.Open(new MemoryStream(rendered.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal("A1:H15", worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal(new[] { "Group 0", "Group 1" }, new[] { "A1", "E1" }.Select(address => _fixture.ReadCellText(rendered.Document, address)));
        Assert.Equal(new[] { "Middle 0", "Middle 1", "Middle 2", "Middle 0", "Middle 1", "Middle 2" },
            new[] { "B2", "C2", "D2", "F2", "G2", "H2" }.Select(address => _fixture.ReadCellText(rendered.Document, address)));
        Assert.Equal("Child 0/Group 0", _fixture.ReadCellText(rendered.Document, "B5"));
        Assert.Equal("Child 1/Group 1", _fixture.ReadCellText(rendered.Document, "F13"));
        Assert.DoesNotContain(worksheet.Descendants<Cell>(), cell => cell.CellReference!.Value!.StartsWith("I", StringComparison.Ordinal));
        Assert.Equal(4, rendered.Bound[0].Rows[0].Count(row => row.Level == 40));
        Assert.All(rendered.Bound[0].Rows[0], row => Assert.Single(row.Records));
        Assert.Empty(rendered.Context.Errors);
    }

    private byte[] ThreeLevelTemplate() => _fixture.Create((workbook, part) =>
    {
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$D$7:$F$13";
        part.Worksheet = new Worksheet(new SheetData(
            new Row(_fixture.Text("D1", "$h(10)"), _fixture.Text("E1", "$h(10)"), _fixture.Text("F1", "$h(10)")) { RowIndex = 1 },
            new Row(_fixture.Text("E2", "$h(40)"), _fixture.Text("F2", "$h(40)")) { RowIndex = 2 },
            new Row(_fixture.Text("F3", "$h(90)")) { RowIndex = 3 },
            new Row(_fixture.Text("A7", "$header(10).Groups"), _fixture.Text("D7", "{Name}")) { RowIndex = 7 },
            new Row(_fixture.Text("A8", "$header(40).Middles"), _fixture.Text("E8", "{Name}")) { RowIndex = 8 },
            new Row(_fixture.Text("A9", "$header(90).Leaves"), _fixture.Text("F9", "{Name}")) { RowIndex = 9 },
            new Row(_fixture.Text("A10", "$l(10).Items"), _fixture.Text("D10", "{Name}")) { RowIndex = 10 },
            new Row(_fixture.Text("A11", "$l(40).Items"), _fixture.Text("E11", "{Name}/{../Name}"), _fixture.Text("F11", "{Value}")) { RowIndex = 11 },
            new Row(_fixture.Text("A12", "$l(40).Items"), _fixture.Text("E12", "{Name}/{../Name}"), _fixture.Text("F12", "{Value}")) { RowIndex = 12 },
            new Row(_fixture.Text("A13", "$l(10).Items"), _fixture.Text("D13", "{Name}")) { RowIndex = 13 }));
    });

    private static string ThreeLevelData(int groups, int middles, int leaves)
    {
        var records = new List<Dictionary<string, object>>();
        void AddGroup(int tupleCount, string path, string name, int valueOffset, int level)
        {
            for (var tuple = 0; tuple < tupleCount; tuple++)
                records.Add(new()
                {
                    ["$level"] = level,
                    ["$path"] = path,
                    ["Name"] = tuple == 0 ? name : "tuple must not replace ordinary name",
                    ["Value"] = valueOffset + tuple + 1
                });
        }

        var arity = groups == 0 || middles == 0 || leaves == 0 ? 1 : checked(groups * middles * leaves);
        for (var group = 0; group < groups; group++)
        {
            AddGroup(arity, "group" + group, "Group " + group, 0, 10);
            for (var sibling = 0; sibling < 2; sibling++)
            {
                AddGroup(arity, "group" + group + "-child" + sibling,
                    "Child " + sibling, sibling * 100, 40);
            }
        }

        return JsonSerializer.Serialize(new
        {
            Groups = Enumerable.Range(0, groups).Select(index => new { Name = "Group " + index }),
            Middles = Enumerable.Range(0, middles).Select(index => new { Name = "Middle " + index }),
            Leaves = Enumerable.Range(0, leaves).Select(index => new { Name = "Leaf " + index }),
            Items = records
        });
    }

    private static string[] LeafColumns() => [
        "C", "D", "E", "F", "H", "I", "J", "K", "M", "N", "O", "P",
        "S", "T", "U", "V", "X", "Y", "Z", "AA", "AC", "AD", "AE", "AF"];
}
