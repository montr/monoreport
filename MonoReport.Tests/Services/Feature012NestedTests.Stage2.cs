using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml;
using MonoReport.Models;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012NestedTests
{
    [Theory]
    [InlineData(0, "A1:B4", "")]
    [InlineData(1, "A1:P4", "A1:H1,I1:P1,A2:H2,I2:P2,A3:C3,I3:K3,E3:H3,M3:P3,A4:C4,I4:K4,E4:H4,M4:P4")]
    [InlineData(2, "A1:AD4", "A1:O1,P1:AD1,A2:O2,P2:AD2,A3:C3,D3:F3,P3:R3,S3:U3,H3:K3,L3:O3,W3:Z3,AA3:AD3,A4:C4,D4:F4,P4:R4,S4:U4,H4:K4,L4:O4,W4:Z4,AA4:AD4")]
    public void Stage2_Multirow_headers_project_parent_and_sibling_merges_with_styles_and_collapse(
        int childCount, string dimension, string mergeList)
    {
        var source = Stage2HeaderTemplate();
        var json = Stage2HeaderData(childCount);

        var rendered = RenderNested(source, json);

        using var package = SpreadsheetDocument.Open(new MemoryStream(rendered.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal(dimension, worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal(mergeList.Split(',', StringSplitOptions.RemoveEmptyEntries).OrderBy(reference => reference, StringComparer.Ordinal),
            worksheet.Descendants<MergeCell>().Select(merge => merge.Reference!.Value!).OrderBy(reference => reference, StringComparer.Ordinal));

        var parentCells = childCount switch { 0 => new[] { "A1", "B1" }, 1 => new[] { "A1", "I1" }, _ => new[] { "A1", "P1" } };
        var parentCodeCells = parentCells.Select(address => address.Replace("1", "2", StringComparison.Ordinal)).ToArray();
        Assert.Equal(new[] { "Parent 0", "Parent 1" }, parentCells.Select(address => _fixture.ReadCellText(rendered.Document, address)));
        Assert.Equal(new[] { "P0", "P1" }, parentCodeCells.Select(address => _fixture.ReadCellText(rendered.Document, address)));
        if (childCount > 0)
        {
            var childAnchors = childCount == 1
                ? new[] { "A3", "E3", "I3", "M3" }
                : new[] { "A3", "D3", "H3", "L3", "P3", "S3", "W3", "AA3" };
            var childCodeAnchors = childAnchors.Select(address => address.Replace("3", "4", StringComparison.Ordinal)).ToArray();
            Assert.Equal(childCount * 4, childAnchors.Length);
            var expectedChildren = Enumerable.Range(0, 4).SelectMany(_ => Enumerable.Range(0, childCount)
                .Select(index => "Child " + index));
            var expectedCodes = Enumerable.Range(0, 4).SelectMany(_ => Enumerable.Range(0, childCount)
                .Select(index => "C" + index));
            Assert.Equal(expectedChildren,
                childAnchors.Select(address => _fixture.ReadCellText(rendered.Document, address)));
            Assert.Equal(expectedCodes,
                childCodeAnchors.Select(address => _fixture.ReadCellText(rendered.Document, address)));
        }

        if (childCount == 0)
        {
            Assert.Empty(worksheet.Descendants<MergeCell>());
            Assert.Equal(new[] { "Parent 0", "Parent 1" }, parentCells.Select(address => _fixture.ReadCellText(rendered.Document, address)));
        }
        else
        {
            // E7 and J9 are styled non-anchor cells in source merges; their styles survive every copied merge.
            var styledNonAnchors = childCount == 1
                ? new[] { "B1", "J1", "B3", "J3", "G3", "O3" }
                : new[] { "B1", "E1", "Q1", "T1", "B3", "E3", "Q3", "T3", "J3", "N3", "Y3", "AC3" };
            Assert.All(styledNonAnchors, address => Assert.Equal(3U, cells[address].StyleIndex!.Value));
        }
        Assert.Empty(rendered.Context.Errors);
    }

    [Fact]
    public void Stage2_PrintArea_keeps_boundary_merges_and_discards_wholly_external_merges()
    {
        var source = _fixture.Create((_, part) =>
        {
            part.Worksheet!.Append(new MergeCells(
                new MergeCell { Reference = "A1:A2" },
                new MergeCell { Reference = "B3:C3" }));
        });

        var rendered = _fixture.Merge(source);

        using var package = SpreadsheetDocument.Open(new MemoryStream(rendered.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal("A1:B1", Assert.Single(worksheet.Descendants<MergeCell>()).Reference!.Value);
        Assert.Equal("A1:C3", worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
    }

    [Fact]
    public void Stage2_PrintArea_rejects_a_merge_crossing_the_crop_boundary_without_returning_a_workbook()
    {
        var source = _fixture.Create((_, part) =>
            part.Worksheet!.Append(new MergeCells(new MergeCell { Reference = "A3:C3" })));
        MergeResult? result = null;

        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source));

        Assert.Null(result);
        Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, error.Code);
        Assert.Contains("wholly inside or outside PrintArea", error.Message);
        Assert.Contains("A3:C3", error.Errors[0].Location);
    }

    private byte[] Stage2HeaderTemplate() => _fixture.Create((workbook, part) =>
    {
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$D$7:$K$10";
        workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()),
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })), new Borders(new Border()),
            new CellStyleFormats(new CellFormat()), new CellFormats(new CellFormat(),
                new CellFormat { FontId = 0, FillId = 0, BorderId = 0 },
                new CellFormat { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } },
                new CellFormat { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Right } },
                new CellFormat { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left } }));
        var rows = part.Worksheet!.GetFirstChild<SheetData>()!;
        rows.RemoveAllChildren();
        rows.Append(
            new Row(Enumerable.Range(4, 8).Select(column => _fixture.Text(ColumnName(column) + "1", "$h(10)"))) { RowIndex = 1 },
            new Row(Enumerable.Range(4, 3).Select(column => _fixture.Text(ColumnName(column) + "2", "$h(20)"))
                .Concat(Enumerable.Range(8, 4).Select(column => _fixture.Text(ColumnName(column) + "2", "$h(20)")))) { RowIndex = 2 },
            new Row(_fixture.Text("A7", "$h(10).Parents"), StyledText("D7", "{Name}", 1),
                new Cell { CellReference = "E7", StyleIndex = 3 }) { RowIndex = 7 },
            new Row(_fixture.Text("A8", "$h(10).Parents"), StyledText("D8", "{Code}", 1)) { RowIndex = 8 },
            new Row(_fixture.Text("A9", "$h(20).Children"), StyledText("D9", "{Name}", 2),
                new Cell { CellReference = "E9", StyleIndex = 3 }, StyledText("H9", "{Name}", 2),
                new Cell { CellReference = "J9", StyleIndex = 3 }) { RowIndex = 9 },
            new Row(_fixture.Text("A10", "$h(20).Children"), _fixture.Text("D10", "{Code}"),
                _fixture.Text("H10", "{Code}")) { RowIndex = 10 });
        part.Worksheet.Append(new MergeCells(
            new MergeCell { Reference = "D7:K7" }, new MergeCell { Reference = "D8:K8" },
            new MergeCell { Reference = "D9:F9" }, new MergeCell { Reference = "H9:K9" },
            new MergeCell { Reference = "D10:F10" }, new MergeCell { Reference = "H10:K10" }));
    });

    private Cell StyledText(string address, string text, uint style) => new()
    {
        CellReference = address, StyleIndex = style, DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve })
    };

    private static string Stage2HeaderData(int children) => System.Text.Json.JsonSerializer.Serialize(new
    {
        Parents = new[] { new { Name = "Parent 0", Code = "P0" }, new { Name = "Parent 1", Code = "P1" } },
        Children = Enumerable.Range(0, children).Select(index => new { Name = "Child " + index, Code = "C" + index })
    });

    private static string ColumnName(int number)
    {
        var result = "";
        while (number > 0)
        {
            number--;
            result = (char)('A' + number % 26) + result;
            number /= 26;
        }
        return result;
    }
}
