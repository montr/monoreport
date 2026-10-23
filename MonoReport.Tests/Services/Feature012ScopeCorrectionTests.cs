using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed class Feature012ScopeCorrectionTests
{
    private readonly XlsxFixture _fixture = new();

    [Theory]
    [InlineData("{../Name}", "$.Items[].Name")]
    [InlineData("{../../Name}", "$.Items[].Name")]
    public void F1_Footer_first_single_ancestor_field_has_exact_canonical_path(string field, string expected)
    {
        var occurrence = Assert.Single(Prepare(Create([40, 10], 3, field)).Occurrences,
            item => item.Kind == "field" && item.Location!.EndsWith("B3"));
        Assert.Equal(expected, occurrence.Path);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void F1_Two_level_ancestor_fields_follow_tree_for_parent_first_and_footer_first(bool footerFirst)
    {
        var source = Create(footerFirst ? [40, 10] : [10, 40], footerFirst ? 3 : 4,
            "{Name}/{../Name}/{../../Name}/{$.Name}");
        var plan = Prepare(source);
        var row = footerFirst ? 3 : 4;
        var fields = plan.Occurrences.Where(item => item.Kind == "field" && item.Location!.EndsWith("B" + row)).ToArray();
        Assert.Equal(new[] { "{Name}", "{../Name}", "{../../Name}", "{$.Name}" }, fields.Select(item => item.Expression));
        Assert.Equal(new[] { "$.Items[].Name", "$.Items[].Name", "$.Items[].Name", "$.Name" }, fields.Select(item => item.Path));
        Assert.All(fields, field => Assert.Equal("/xl/worksheets/sheet1.xml", field.Part));
        AssertGuarded(source);
    }

    [Theory]
    [InlineData("10,40,90,40,10", 5)]
    [InlineData("90,40,10", 3)]
    [InlineData("10,90,40,10", 4)]
    [InlineData("90,90,40,40,10,10", 4)]
    public void F1_Multilevel_and_multirow_blocks_use_structural_ancestors(string sequence, int fieldRow)
    {
        var source = Create(sequence.Split(',').Select(value => (int?)int.Parse(value)).ToArray(), fieldRow,
            "{../Name}/{../../Name}/{../../../Name}/{$.Name}");
        var plan = Prepare(source);
        Assert.Equal(new[] { "$.Items[].Name", "$.Items[].Name", "$.Items[].Name", "$.Name" },
            plan.Occurrences.Where(item => item.Kind == "field" && item.Location!.EndsWith("B" + fieldRow)).Select(item => item.Path));
        AssertGuarded(source);
    }

    [Theory]
    [InlineData("10,40,90,40,10", 7)]
    [InlineData("90,40,10", 5)]
    public void M35_Parent_footer_has_data_relative_scope_independent_of_source_depth(string sequence, int footerRow)
    {
        var source = Create(sequence.Split(',').Select(value => (int?)int.Parse(value)).ToArray(), footerRow, "{../Name}");
        var occurrence = Assert.Single(Prepare(source).Occurrences, item => item.Kind == "field" && item.Expression == "{../Name}");
        Assert.Equal("$.Items[].Name", occurrence.Path);
        AssertGuarded(source);
    }

    [Theory]
    [InlineData("90,40,10", 3, "{../../../../Name}")]
    [InlineData("10,40,90,40,10", 7, "{../../Name}")]
    public void M35_Excess_source_parent_depth_is_deferred_to_data_binding(string sequence, int row, string field)
    {
        var source = Create(sequence.Split(',').Select(value => (int?)int.Parse(value)).ToArray(), row, field);
        var occurrence = Assert.Single(Prepare(source).Occurrences, item => item.Kind == "field" && item.Expression == field);
        Assert.Equal("$.Items[].Name", occurrence.Path);
        Assert.Equal("/xl/worksheets/sheet1.xml, sheet 'Report', B" + row, occurrence.Location);
        AssertGuarded(source);
    }

    [Fact]
    public void F1_Separate_vertical_regions_do_not_share_ancestor_scopes()
    {
        var source = Create([90, 40, 10, null, 40], 7, "{../Name}");
        var occurrence = Assert.Single(Prepare(source).Occurrences, item => item.Kind == "field" && item.Expression == "{../Name}");
        Assert.Equal("$.Items[].Name", occurrence.Path);
        AssertGuarded(source);
    }

    [Fact]
    public void F1_Horizontal_headers_keep_separate_current_parent_and_root_scopes()
    {
        var source = new XlsxTwoAxisFixture().Create((_, worksheet) =>
        {
            var row = worksheet.Worksheet!.Descendants<Row>().Single(item => item.RowIndex!.Value == 8);
            row.Elements<Cell>().Single(cell => cell.CellReference == "E8").InlineString!.Text!.Text =
                "{Name}/{../Company}/{../../Title}/{$.Title}";
        });
        var fields = Prepare(source).Occurrences.Where(item => item.Kind == "field" && item.Location!.EndsWith("E8")).ToArray();
        Assert.Equal(new[] { "$.Criteria[].Name", "$.Bids[].Company", "$.Title", "$.Title" }, fields.Select(item => item.Path));
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
    }

    [Fact]
    public void M35_Static_data_dependencies_preserve_original_source_tree()
    {
        var prepared = Prepare(Create([90, 40, 10], 3, "{../../Name}"));
        var sheet = Assert.Single(prepared.Sheets);
        var region = Assert.Single(sheet.Regions);
        using var services = _fixture.Services();
        var structured = services.GetRequiredService<XlsxTemplatePlanner>().Structure(sheet);
        Assert.Same(region, Assert.Single(structured.Regions));
        Assert.Equal(10, region.Root.Level);
        Assert.Equal(new XlsxVerticalBlock(5, 5), region.Root.Footer);
        var child = Assert.Single(region.Root.Children);
        Assert.Equal(40, child.Level); Assert.Equal(new XlsxVerticalBlock(4, 4), child.Footer);
        var leaf = Assert.Single(child.Children);
        Assert.Equal(90, leaf.Level); Assert.Equal(new XlsxVerticalBlock(3, 3), leaf.Body);
        Assert.Equal("$.Items[].Name", Assert.Single(prepared.Occurrences,
            item => item.Kind == "field" && item.Expression == "{../../Name}").Path);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void F2_Guard_names_pending_emission_and_placement_and_returns_no_report(bool footerFirst)
    {
        var source = Create(footerFirst ? [40, 10] : [10, 40], 3, "authored text");
        var issue = AssertGuarded(source);
        Assert.Contains("nested emission and placement", issue.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("hierarchy binding", issue.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("M3", issue.Message);
    }

    private XlsxTemplatePlan Prepare(byte[] source)
    {
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(source), false);
        var context = new XlsxOperationContext(new MergeOptions());
        return services.GetRequiredService<XlsxTemplatePlanner>().Prepare(
            services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
    }

    private GenerationDiagnostic AssertGuarded(byte[] source)
    {
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.True(analysis.Valid); Assert.Empty(analysis.Errors); Assert.Empty(analysis.Warnings);
        var issue = Assert.Single(Assert.Throws<TemplateException>(() => _fixture.Merge(source, "{\"Items\":[]}")).Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, issue.Code);
        Assert.Equal("/xl/worksheets/sheet1.xml, sheet 'Report', A3", issue.Location);
        foreach (var json in new[] { "{\"Items\":[]}", "{\"Items\":null}", "{\"Items\":[{\"$level\":10,\"Name\":\"parent\"},{\"$level\":40,\"Name\":\"child\"}]}" })
        {
            MergeResult? result = null;
            var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, json));
            Assert.Null(result);
            Assert.Equal(issue.Code, error.Code); Assert.Equal(issue.Message, error.Errors[0].Message);
            Assert.Equal(issue.Location, error.Errors[0].Location); Assert.False(error.LimitReached);
        }
        Assert.Equal(source, input.ToArray());
        return issue;
    }

    private byte[] Create(int?[] levels, int fieldRow, string field) => _fixture.Create((workbook, worksheet) =>
    {
        var data = worksheet.Worksheet!.GetFirstChild<SheetData>()!; data.RemoveAllChildren();
        for (var index = 0; index < levels.Length; index++)
        {
            var row = index + 3;
            var element = new Row { RowIndex = (uint)row };
            if (levels[index] is { } level) element.Append(_fixture.Text("A" + row, "$l(" + level + ").Items"));
            element.Append(_fixture.Text("B" + row, row == fieldRow ? field : "{Name}"));
            data.Append(element);
        }
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$D$" + (levels.Length + 2);
    });
}
