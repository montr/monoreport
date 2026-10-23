using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012FormulaTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OwnerRanges_Body_range_partially_overlapping_parent_header_rejects_incomplete_projection(bool mergedHeader)
    {
        var source = _fixture.Create((workbook, part) =>
        {
            part.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A2", "$l(10).Items"), _fixture.Text("B2", "{Name}")) { RowIndex = 2 },
                new Row(_fixture.Text("A3", "$l(40).Items"), _fixture.Text("B3", "{Price}"),
                    GroupFormula("C3", "SUM(B2:B3)")) { RowIndex = 3 }));
            if (mergedHeader) part.Worksheet.Append(new MergeCells(new MergeCell { Reference = "B2:C2" }));
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$2:$C$3";
        });
        var original = source.ToArray();
        MergeResult? result = null;

        // Exercise public generation, including source preflight and final owner checks.
        // Both B2 (parent header) and B3 (current body) are operands; dropping B2 is unsafe.
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source,
            """{"Items":[{"$level":10,"Name":"Parent"},{"$level":40,"Price":11}]}"""));

        Assert.Null(result);
        Assert.Equal(original, source);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
        var issue = Assert.Single(error.Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
        Assert.Contains("sheet 'Report'", issue.Location);
        Assert.EndsWith(", C3", issue.Location);
        Assert.Contains("Reference B2:B3", issue.Message);
        Assert.Contains("empty, discarded, disjoint or ambiguous image", issue.Message);
        Assert.Contains("Re-author it inside one compatible owner using a retained contiguous A1 range", issue.Message);
    }

    [Theory]
    [InlineData("B2:B3", false)] // Missing trailing child row, even though it has no authored B3 cell.
    [InlineData("B3:B2", false)]
    [InlineData("B2:B4", true)] // Header and footer survive contiguously around a missing interior row.
    [InlineData("B4:B2", true)]
    public void OwnerRanges_Missing_blank_child_rows_cannot_contract_group_ranges(string operand, bool footer)
    {
        var source = _fixture.Create((workbook, part) =>
        {
            var rows = new SheetData(
                new Row(_fixture.Text("A2", "$l(10).Items"), _fixture.Text("B2", "{Name}"),
                    GroupFormula("C2", "COUNTBLANK(" + operand + ")")) { RowIndex = 2 },
                new Row(_fixture.Text("A3", "$l(40).Items")) { RowIndex = 3 });
            if (footer) rows.Append(new Row(_fixture.Text("A4", "$l(10).Items"), _fixture.Text("B4", "{Name}")) { RowIndex = 4 });
            part.Worksheet = new Worksheet(rows);
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = footer ? "Report!$B$2:$C$4" : "Report!$B$2:$C$3";
        });
        // Q's child must not rescue P's incomplete projection.
        AssertOwnerRangeUnsafe(source,
            """{"Items":[{"$level":10,"Name":"P"},{"$level":10,"Name":"Q"},{"$level":40}]}""", "C2", operand);
    }

    [Theory]
    [InlineData("D9:E9")] // Trailing blank column.
    [InlineData("F9:D9")] // Interior blank column, descending.
    [InlineData("D7:F7")] // Ordinary cutout piece beside retained heading content.
    [InlineData("F7:D7")]
    public void OwnerRanges_Empty_horizontal_child_cannot_discard_blank_source_columns(string operand)
    {
        var fixture = new XlsxTwoAxisFixture();
        var source = fixture.Create((_, part) => part.Worksheet!.Descendants<Row>().Single(row => row.RowIndex!.Value == 9)
            .Append(GroupFormula("C9", "COUNTBLANK(" + operand + ")")));
        AssertOwnerRangeUnsafe(source, fixture.Data(0), "C9", operand);
    }

    [Theory]
    [InlineData("B$2:B3")]
    [InlineData("B3:B$2")]
    [InlineData("$B$2:C3")]
    [InlineData("C3:$B$2")]
    public void OwnerRanges_Mixed_endpoints_reject_incomplete_alternate_scope_before_geometry_comparison(string operand)
    {
        var source = OwnerRangeBodyWorkbook(operand);
        // The global projection covers both rows; the relative body scope cannot cover the parent header.
        AssertOwnerRangeUnsafe(source, """{"Items":[{"$level":10,"Name":"P"},{"$level":40,"Price":11}]}""", "C3", operand);
    }

    [Theory]
    [InlineData("B2:B3")]
    [InlineData("B3:B2")]
    public void OwnerRanges_Incomplete_primary_mixed_projection_uses_source_image_diagnostic(string operand)
    {
        var source = _fixture.Create((workbook, part) =>
        {
            part.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A2", "$l(10).Items"), _fixture.Text("B2", "{Name}"),
                    GroupFormula("C2", "COUNTBLANK(" + operand.Replace("B3", "B$3", StringComparison.Ordinal) + ")")) { RowIndex = 2 },
                new Row(_fixture.Text("A3", "$l(40).Items")) { RowIndex = 3 }));
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "Report!$B$2:$C$3";
        });
        AssertOwnerRangeUnsafe(source, """{"Items":[{"$level":10,"Name":"P"}]}""", "C2",
            operand.Replace("B3", "B$3", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("B2:D4")]
    [InlineData("D4:B2")]
    public void OwnerRanges_Deeper_body_cannot_borrow_incompatible_content_owners_from_ancestors_or_other_groups(string operand)
    {
        var source = _fixture.Create((workbook, part) =>
        {
            part.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A2", "$l(10).Items"), _fixture.Text("B2", "{Name}")) { RowIndex = 2 },
                new Row(_fixture.Text("A3", "$l(20).Items"), _fixture.Text("B3", "{Name}")) { RowIndex = 3 },
                new Row(_fixture.Text("A4", "$l(40).Items"), _fixture.Text("B4", "{Price}"),
                    GroupFormula("D4", "COUNTBLANK(" + operand + ")")) { RowIndex = 4 }),
                new MergeCells(new MergeCell { Reference = "B2:C2" }, new MergeCell { Reference = "B3:C3" },
                    new MergeCell { Reference = "B4:C4" }));
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "Report!$B$2:$D$4";
        });
        AssertOwnerRangeUnsafe(source,
            """{"Items":[{"$level":10,"Name":"P"},{"$level":20,"Name":"D"},{"$level":40,"Price":11},{"$level":40,"Price":11},{"$level":10,"Name":"Q"},{"$level":20,"Name":"D"},{"$level":40,"Price":11}]}""",
            "D4", operand);
    }

    [Theory]
    [InlineData("B3:B3", "A2:A2", "A3:A3", "A5:A5")]
    [InlineData("$B$2:$B$3", "$A$1:$A$5", "$A$1:$A$5", "$A$1:$A$5")]
    [InlineData("$B$3:$B$2", "$A$5:$A$1", "$A$5:$A$1", "$A$5:$A$1")]
    public void OwnerRanges_Complete_body_and_global_ranges_preserve_selected_repetitions(string operand, string first, string second, string third)
    {
        var source = OwnerRangeBodyWorkbook(operand);
        var original = source.ToArray();
        var result = RenderGroups(source,
            """{"Items":[{"$level":10,"Name":"P"},{"$level":40,"Price":11},{"$level":40,"Price":11},{"$level":10,"Name":"Q"},{"$level":40,"Price":11}]}""");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>()
            .Where(cell => cell.CellFormula is not null).ToArray();
        Assert.Equal(new[] { "B2", "B3", "B5" }, cells.Select(cell => cell.CellReference!.Value));
        Assert.Equal(new[] { first, second, third }.Select(range => "SUM(" + range + ")"), cells.Select(cell => cell.CellFormula!.Text));
        Assert.Equal(new[] { "0,1", "0,2", "3,4" }, result.Output.Sheets[0].Rows.SelectMany(row => row.Cells)
            .Where(cell => cell.Formula is not null).Select(cell => string.Join(",", cell.Id.Vertical!.LogicalRows)));
        Assert.All(cells, cell => Assert.Null(cell.CellValue));
        Assert.Equal(original, source);
        Assert.Empty(result.Context.Errors); Assert.Empty(result.Context.Warnings);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void OwnerRanges_Each_removed_address_requires_one_jointly_compatible_owner(bool retained)
    {
        var fixture = new XlsxTwoAxisFixture();
        var bytes = fixture.Create(); var original = bytes.ToArray();
        var source = fixture.Plan(bytes).Sheets[0].Source;
        var vertical = new VerticalInstancePath(0, [0]);
        var horizontal = new HorizontalInstancePath(0, 0, null);
        // Isolate the compatibility invariant: one address has separate vertical/horizontal matches,
        // another address is jointly compatible, and an ordinary source piece also survives.
        var owners = new List<XlsxContentOwner>
        {
            new(new(source.Id, new(7, 2)), vertical, horizontal with { ItemIndex = 1 }, new(1, 1)),
            new(new(source.Id, new(7, 2)), new(0, [1]), horizontal, new(1, 1)),
            new(new(source.Id, new(7, 3)), vertical, horizontal, new(1, 2))
        };
        if (retained) owners.Add(new(new(source.Id, new(7, 2)), vertical, horizontal, new(1, 1)));
        var geometry = new XlsxSheetGeometry(source, new(new(1, 1), new(6, 5)), -6, -1)
        {
            Expanded = true, RowSpans = [new(7, 12, 1)], ColumnSpans = [new(2, 6, 1)],
            ContentOwners = [.. owners]
        };
        using var services = _fixture.Services();
        var image = services.GetRequiredService<XlsxGeometryPlanner>().ProjectReference(new(new(7, 2), new(7, 4)), geometry, vertical, horizontal);
        Assert.Equal(XlsxRangeImageKind.Rectangle, image.Kind);
        Assert.Equal(retained, image.Complete);
        Assert.Equal(new XlsxRectangle(new(1, retained ? 1 : 2), new(1, 3)), image.Rectangle);
        var cell = source.Cells.Single(cell => cell.Id.Address == new XlsxAddress(7, 4));
        var placement = new XlsxCellPlacement(new(cell.Id, vertical, horizontal), new(1, 3), cell, null);
        var formula = services.GetRequiredService<XlsxFormulaReferences>().Read("SUM(B7:D7)", "D7");
        var planner = services.GetRequiredService<XlsxFormulaPlanner>();
        if (retained) Assert.Equal("SUM(A1:C1)", planner.BindFormula(formula, placement, geometry));
        else
        {
            var error = Assert.Throws<TemplateException>(() => planner.BindFormula(formula, placement, geometry));
            Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
            Assert.Contains("sheet 'Report', D7", error.Location);
            Assert.Contains("Reference B7:D7", error.Message);
            Assert.Contains("retained contiguous A1 range", error.Message);
        }
        Assert.Equal(original, bytes);
    }

    private byte[] OwnerRangeBodyWorkbook(string operand) => _fixture.Create((workbook, part) =>
    {
        part.Worksheet = new Worksheet(new SheetData(
            new Row(_fixture.Text("A2", "$l(10).Items"), _fixture.Text("B2", "{Name}")) { RowIndex = 2 },
            new Row(_fixture.Text("A3", "$l(40).Items"), _fixture.Text("B3", "{Price}"),
                GroupFormula("C3", "SUM(" + operand + ")")) { RowIndex = 3 }));
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "Report!$B$2:$C$3";
    });

    private void AssertOwnerRangeUnsafe(byte[] source, string data, string address, string operand)
    {
        var original = source.ToArray(); MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, data));
        Assert.Null(result); Assert.Equal(original, source);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
        Assert.All(error.Errors, issue => Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code));
        var issue = error.Errors.First(issue => issue.Location!.EndsWith(", " + address, StringComparison.Ordinal));
        Assert.Contains("sheet 'Report'", issue.Location);
        Assert.Contains("Reference " + operand, issue.Message);
        Assert.Contains("empty, discarded, disjoint or ambiguous image", issue.Message);
        Assert.Contains("Re-author it inside one compatible owner using a retained contiguous A1 range", issue.Message);
    }
}
