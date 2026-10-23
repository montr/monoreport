using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012NestedTests
{
    private readonly XlsxFixture _fixture = new();

    [Theory]
    [InlineData("10,40,90,40,10")]
    [InlineData("10,40,90")]
    [InlineData("90,40,10")]
    [InlineData("10,40,90,10")]
    [InlineData("10,90,40,10")]
    [InlineData("10,10,40,40,90,90,40,40,10,10")]
    [InlineData("-2147483648,0,2147483647,0,-2147483648")]
    public void Grammar_Parses_nonconsecutive_numeric_levels_with_optional_headers_and_footers(string sequence)
    {
        var levels = sequence.Split(',').Select(int.Parse).ToArray();
        var sheet = Plan(Create(levels.Select(level => "$l(" + level + ").Items").ToArray())).Sheets.Single();
        var region = Assert.Single(sheet.Regions);
        Assert.Equal((3, levels.Length + 2, levels.Min(), "$.Items"),
            (region.FirstRow, region.LastRow, region.Level, region.Collection));
    }

    [Fact]
    public void Grammar_Normalizes_keywords_and_forms_before_coalescing()
    {
        var sheet = Plan(Create("$level(+10).Items", "$L(10).Items", "$LEVEL(40).Items",
            "$l(40).Items", "$l(90).Items", "$level(40).Items", "$l(10).Items")).Sheets.Single();
        Assert.Equal((3, 9, 10, "$.Items"),
            (sheet.Regions[0].FirstRow, sheet.Regions[0].LastRow, sheet.Regions[0].Level, sheet.Regions[0].Collection));
        var parent = sheet.Regions[0].Root;
        Assert.Equal(10, parent.Level);
        AssertBlock(parent.Header, 3, 4); AssertBlock(parent.Footer, 9, 9); Assert.Null(parent.Body);
        var middle = Assert.Single(parent.Children);
        Assert.Equal(40, middle.Level);
        AssertBlock(middle.Header, 5, 6); AssertBlock(middle.Footer, 8, 8); Assert.Null(middle.Body);
        var leaf = Assert.Single(middle.Children);
        Assert.Equal(90, leaf.Level);
        AssertBlock(leaf.Body, 7, 7); Assert.Null(leaf.Header); Assert.Null(leaf.Footer); Assert.Empty(leaf.Children);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void Grammar_Assigns_header_only_and_footer_only_blocks_to_their_own_levels(bool footerOnly)
    {
        var levels = footerOnly ? new[] { 90, 40, 10 } : new[] { 10, 40, 90 };
        var root = Assert.Single(Plan(Create(levels.Select(level => "$l(" + level + ").Items").ToArray()))
            .Sheets.Single().Regions).Root;
        Assert.Equal(10, root.Level); Assert.Null(root.Body);
        AssertBlock(footerOnly ? root.Footer : root.Header, footerOnly ? 5 : 3, footerOnly ? 5 : 3);
        Assert.Null(footerOnly ? root.Header : root.Footer);
        var child = Assert.Single(root.Children);
        Assert.Equal(40, child.Level); Assert.Null(child.Body);
        AssertBlock(footerOnly ? child.Footer : child.Header, 4, 4);
        Assert.Null(footerOnly ? child.Header : child.Footer);
        var leaf = Assert.Single(child.Children);
        Assert.Equal(90, leaf.Level); Assert.Empty(leaf.Children);
        AssertBlock(leaf.Body, footerOnly ? 3 : 5, footerOnly ? 3 : 5);
        Assert.Null(leaf.Header); Assert.Null(leaf.Footer);
    }

    [Fact]
    public void Grammar_Deeper_body_before_middle_footer_belongs_to_that_middle_level()
    {
        var root = Assert.Single(Plan(Create("$l(10).Items", "$l(90).Items", "$l(40).Items", "$l(10).Items"))
            .Sheets.Single().Regions).Root;
        Assert.Equal(10, root.Level); AssertBlock(root.Header, 3, 3); AssertBlock(root.Footer, 6, 6);
        var middle = Assert.Single(root.Children);
        Assert.Equal(40, middle.Level); Assert.Null(middle.Header); Assert.Null(middle.Body);
        AssertBlock(middle.Footer, 5, 5);
        var body = Assert.Single(middle.Children);
        Assert.Equal(90, body.Level); AssertBlock(body.Body, 4, 4);
        Assert.Null(body.Header); Assert.Null(body.Footer); Assert.Empty(body.Children);
    }

    [Theory]
    [InlineData("40,10,40", "A4")]
    [InlineData("10,40,10,40", "A5")]
    [InlineData("10,40,10,40,10", "A5")]
    [InlineData("10,40,90,40,90,40,10", "A6")]
    public void Grammar_Rejects_ambiguous_interleavings_with_source_and_boundary_remedy(string sequence, string address)
    {
        var source = Create(sequence.Split(',').Select(level => "$l(" + level + ").Items").ToArray());
        AssertRejected(source, DiagnosticCodes.TemplateInvalidStructure, address, "unmarked control row");
    }

    [Theory]
    [InlineData("$l(40).items")]
    [InlineData("$l(40).Other")]
    public void Grammar_Does_not_coalesce_different_case_sensitive_collections(string child)
    {
        AssertRejected(Create("$l(10).Items", child, "$l(10).Items"),
            DiagnosticCodes.TemplateInvalidStructure, "A4", "unmarked control row");
    }

    [Fact]
    public void Grammar_Unmarked_rows_split_independent_runs_even_with_identical_markers()
    {
        var sheet = Plan(Create("$l(10).Items", "$l(40).Items", "$l(10).Items", null,
            "$level(10).Items", "$l(90).Items", "$l(10).Items")).Sheets.Single();
        Assert.Equal(new[] { (0, 3, 5, 10), (1, 7, 9, 10) },
            sheet.Regions.Select(region => (region.Id, region.FirstRow, region.LastRow, region.Level)));
        Assert.DoesNotContain(sheet.Controls, marker => marker.Cell.Id.Address.Row == 6);
    }

    [Fact]
    public void Grammar_Static_and_single_level_boundaries_still_parse()
    {
        Assert.Empty(Plan(Create((string?)null)).Sheets.Single().Regions);
        var region = Assert.Single(Plan(Create("$level(40).Items", "$L(40).Items")).Sheets.Single().Regions);
        Assert.Equal((3, 4, 40), (region.FirstRow, region.LastRow, region.Level));
        AssertBlock(region.Root.Body, 3, 4);
        Assert.Null(region.Root.Header); Assert.Null(region.Root.Footer); Assert.Empty(region.Root.Children);
    }

    [Fact]
    public void Grammar_Parsed_nested_source_analyzes_and_public_empty_generation_succeeds()
    {
        var source = Create("$l(10).Items", "$l(40).Items", "$l(10).Items");
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
        var result = _fixture.Merge(source, "{\"Items\":[]}");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>());
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package));
        Assert.Equal(source, input.ToArray());
    }

    private byte[] Create(params string?[] markers) => _fixture.Create((workbook, worksheet) =>
    {
        var data = worksheet.Worksheet!.GetFirstChild<SheetData>()!;
        data.RemoveAllChildren();
        for (var index = 0; index < markers.Length; index++)
        {
            var row = index + 3;
            var element = new Row { RowIndex = (uint)row };
            if (markers[index] is { } marker) element.Append(_fixture.Text("A" + row, marker));
            element.Append(_fixture.Text("B" + row, "Authored row " + row));
            data.Append(element);
        }
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$D$" + (markers.Length + 2);
    });

    private XlsxTemplatePlan Plan(byte[] source)
    {
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(source), false);
        var context = new XlsxOperationContext(new MergeOptions());
        return services.GetRequiredService<XlsxTemplatePlanner>().Plan(
            services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
    }

    private void AssertRejected(byte[] source, string code, string address, string remedy)
    {
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(analysis.Valid);
        var issue = Assert.Single(analysis.Errors);
        Assert.Equal(code, issue.Code);
        Assert.Contains("sheet 'Report'", issue.Location);
        Assert.Contains(address, issue.Location);
        Assert.Contains(remedy, issue.Message);
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source, "{\"Items\":[]}"));
        Assert.Equal(issue.Code, error.Code);
        Assert.Equal(issue.Location, error.Errors[0].Location);
        Assert.Equal(issue.Message, error.Errors[0].Message);
        Assert.Equal(source, input.ToArray());
    }

    private void AssertBlock(XlsxVerticalBlock? block, int first, int last)
    {
        Assert.NotNull(block);
        Assert.Equal((first, last), (block.FirstRow, block.LastRow));
    }
}
