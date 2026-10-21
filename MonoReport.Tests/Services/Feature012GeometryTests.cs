using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services.Xlsx;
using MonoReport.Models.Xlsx;
using System.Text.Json;
using System.Xml.Linq;

namespace MonoReport.Tests.Services;

public sealed class Feature012GeometryTests
{
    private readonly XlsxFixture _fixture = new();
    private readonly XlsxTwoAxisFixture _twoAxis = new();

    [Theory]
    [InlineData("B7")] [InlineData("D8")]
    public void Controls_Header_fields_outside_their_matching_owner_reject_in_analysis_and_generation(string address)
    {
        var bytes = _twoAxis.Create((_, part) =>
        {
            var row = part.Worksheet!.Descendants<Row>().Single(row => row.RowIndex!.Value == uint.Parse(address[1..]));
            row.InsertAt(_fixture.Text(address, "{Name}"), 1);
        });
        using var services = _fixture.Services(); using var stream = new MemoryStream(bytes);
        var analysis = services.GetRequiredService<MonoReport.Services.TemplateAnalyzer>().Analyze(stream);
        var issue = Assert.Single(analysis.Errors); Assert.Contains(address, issue.Location);
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(bytes, _twoAxis.Data()));
        Assert.Equal(issue.Code, error.Code); Assert.Equal(issue.Message, error.Errors[0].Message);
        Assert.Contains("matching horizontal block", issue.Message);
    }

    [Theory]
    [InlineData(1, "C1:D1", "E1:F1")]
    [InlineData(3, "C1:F1", "G1:J1")]
    [InlineData(7, "C1:J1", "K1:R1")]
    public void Merges_Project_existing_parent_owned_heading_over_contained_children(int criteria, string first, string second)
    {
        var sheet = _twoAxis.Plan(_twoAxis.Create()).Sheets.Single();
        using var services = _fixture.Services(); using var json = JsonDocument.Parse(_twoAxis.Data(criteria));
        var context = new XlsxOperationContext(new MergeOptions()); var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
        var positions = geometry.ExpansionGeometry(sheet, planner.BindSheet(sheet, json.RootElement, context), context);
        var output = geometry.EmitCells(sheet, positions); geometry.ProjectMerges(sheet, output);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Assert.Equal(new[] { first, second }, output.Settings.Descendants(ns + "mergeCell").Select(cell => (string?)cell.Attribute("ref")));
    }

    [Theory]
    [InlineData(0, 5, 4)] [InlineData(1, 6, 5)] [InlineData(4, 9, 8)]
    public void Placement_Composes_repeated_rows_static_tail_and_crop_once(int count, int height, int below)
    {
        var sheet = _twoAxis.Plan(_twoAxis.Create()).Sheets.Single();
        using var services = _fixture.Services(); using var json = JsonDocument.Parse(_twoAxis.Data(3, count));
        var context = new XlsxOperationContext(new MergeOptions());
        var data = services.GetRequiredService<XlsxTemplatePlanner>().BindSheet(sheet, json.RootElement, context);
        var geometry = services.GetRequiredService<XlsxGeometryPlanner>().ExpansionGeometry(sheet, data, context);
        Assert.Equal((height, 11), (geometry.OutputArea.Height, geometry.OutputArea.Width));
        Assert.Equal(1, geometry.RowSpans[0].OutputFirst);
        Assert.Equal(below, geometry.RowSpans.Single(span => span.First == 11).OutputFirst);
        Assert.Equal(Enumerable.Range(0, count).Select(index => index * 6), geometry.RowSpans.Where(span => span.Vertical is not null).Select(span => span.Vertical!.LogicalRows[0]));
        Assert.Equal(Enumerable.Range(4, count), geometry.RowSpans.Where(span => span.Vertical is not null).Select(span => span.OutputFirst));
    }

    [Theory]
    [InlineData(0, 5)] [InlineData(1, 7)] [InlineData(3, 11)] [InlineData(7, 19)]
    public void Horizontal_Uses_collection_width_and_preserves_parent_own_columns(int criteria, int finalWidth)
    {
        var sheet = _twoAxis.Plan(_twoAxis.Create()).Sheets.Single();
        using var services = _fixture.Services(); using var json = JsonDocument.Parse(_twoAxis.Data(criteria));
        var data = services.GetRequiredService<XlsxTemplatePlanner>().BindSheet(sheet, json.RootElement, new(new MergeOptions()));
        var spans = services.GetRequiredService<XlsxGeometryPlanner>().HorizontalSpans(sheet, data);
        Assert.Equal(finalWidth, spans[^1].OutputLast);
        Assert.Equal(new[] { 3, 4 + criteria }, spans.Where(span => span.First == 4).Select(span => span.OutputFirst));
        Assert.Equal(2 * criteria, spans.Count(span => span.First == 5));
        Assert.All(spans.Where(span => span.First == 5), span => Assert.NotNull(span.Horizontal!.Parent));
        Assert.Equal(6, spans[^1].First); // ordinary trailing column survives
    }

    [Fact]
    public void Horizontal_Empty_outer_has_no_orphan_children()
    {
        var sheet = _twoAxis.Plan(_twoAxis.Create()).Sheets.Single();
        using var services = _fixture.Services(); using var json = JsonDocument.Parse(_twoAxis.Data(3, 1, 0));
        var data = services.GetRequiredService<XlsxTemplatePlanner>().BindSheet(sheet, json.RootElement, new(new MergeOptions()));
        var spans = services.GetRequiredService<XlsxGeometryPlanner>().HorizontalSpans(sheet, data);
        Assert.Equal(3, spans[^1].OutputLast); Assert.All(spans, span => Assert.Null(span.Horizontal));
    }

    [Fact]
    public void Controls_Normalize_long_short_forms_and_build_contained_parent_child_tree()
    {
        var sheet = _twoAxis.Plan(_twoAxis.Create()).Sheets.Single();
        Assert.Equal((4, 5, 1), (sheet.Horizontal!.FirstColumn, sheet.Horizontal.LastColumn, sheet.Horizontal.Level));
        var child = Assert.Single(sheet.Horizontal.Children);
        Assert.Equal((5, 5, 2), (child.FirstColumn, child.LastColumn, child.Level));
        var region = Assert.Single(sheet.Regions);
        Assert.Equal((10, 10, 10, "$.Items"), (region.FirstRow, region.LastRow, region.Level, region.Collection));
    }

    [Theory]
    [InlineData("F2", "$h(2)", "inside")]
    [InlineData("B1", "$h(1)", "one outer")]
    [InlineData("A9", "$h(1).Other", "case-sensitive")]
    public void Controls_Reject_orphans_second_roots_and_conflicting_associations(string address, string marker, string remedy)
    {
        var bytes = _twoAxis.Create((_, worksheet) =>
        {
            var row = worksheet.Worksheet!.Descendants<Row>().Single(item => item.RowIndex!.Value == uint.Parse(address[1..]));
            var cell = _fixture.Text(address, marker);
            if (address.StartsWith('A') || address.StartsWith('B')) row.PrependChild(cell); else row.Append(cell);
        });
        var error = Assert.Throws<TemplateException>(() => _twoAxis.Plan(bytes));
        Assert.Contains(remedy, error.Message);
    }
}
