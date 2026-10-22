using System.Collections.Immutable;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012NestedTests
{
    [Theory]
    [InlineData(3, 2, 0)] [InlineData(1, 1, 1)] [InlineData(3, 2, 3)]
    [InlineData(0, 2, 2)] [InlineData(3, 0, 2)]
    public void Schedule_Flat_multirow_blocks_keep_original_source_and_one_owner_per_tuple_set(int criteria, int companies, int count)
    {
        var fixture = new XlsxTwoAxisFixture();
        var sheet = fixture.Plan(fixture.Create((_, part) => part.Worksheet!.Descendants<Row>()
            .Single(row => row.RowIndex!.Value == 11).PrependChild(_fixture.Text("A11", "$level(10).Items")))).Sheets.Single();
        using var services = _fixture.Services(); using var json = JsonDocument.Parse(fixture.Data(criteria, count, companies));
        var planner = services.GetRequiredService<XlsxTemplatePlanner>(); var context = new XlsxOperationContext(new MergeOptions());
        var bound = planner.BindSheet(sheet, json.RootElement, context);
        var schedule = planner.ScheduleSheet(sheet, bound, context);
        Assert.Equal(sheet.Source.Id, schedule.SheetId);
        var region = Assert.Single(schedule.Regions);
        Assert.Equal(0, region.RegionId); Assert.Equal(count, region.Blocks.Length);
        var arity = Math.Max(1, companies * criteria);
        Assert.Equal(Enumerable.Range(0, count).Select(index => index * arity), region.Blocks.Select(block => Assert.Single(block.Owner.LogicalRows)));
        Assert.All(region.Blocks, block =>
        {
            Assert.Same(sheet.Regions[0].Root.Body, block.Source);
            Assert.Equal(new XlsxVerticalBlock(10, 11), block.Source);
            Assert.Equal(XlsxLogicalBlockRole.Body, block.Role); Assert.Equal(0, block.Owner.RegionId);
            Assert.Equal(arity, bound.Rows[0].Single(row => row.FirstIndex == block.Owner.LogicalRows[0]).Records.Length);
        });
        Assert.Empty(context.Warnings); Assert.Empty(context.Errors);
    }

    [Fact]
    public void Schedule_Omissions_keep_exact_warning_content_source_and_input_order()
    {
        var sheet = Plan(Create("$l(40).Items", "$level(40).Items")).Sheets.Single();
        using var services = _fixture.Services();
        using var json = JsonDocument.Parse("""{"Items":[{"$level":10},{"$level":40},{"$level":90},{"$level":40}]}""");
        var planner = services.GetRequiredService<XlsxTemplatePlanner>(); var context = new XlsxOperationContext(new MergeOptions());
        var bound = planner.BindSheet(sheet, json.RootElement, context);
        var schedule = planner.ScheduleSheet(sheet, bound, context);
        Assert.Equal(new[] { 1, 3 }, Assert.Single(schedule.Regions).Blocks.Select(block => Assert.Single(block.Owner.LogicalRows)));
        Assert.Equal(new[] { "$.Items[0]", "$.Items[2]" }, context.Warnings.Select(warning => warning.DataPath));
        var location = sheet.Controls.Single(marker => marker.Cell.Id.Address == new XlsxAddress(3, 1)).Location;
        Assert.All(context.Warnings, warning =>
        {
            Assert.Equal(DiagnosticCodes.XlsxMissingLevel, warning.Code);
            Assert.Equal("No template block matches this data level; its own rows were omitted.", warning.Message);
            Assert.Equal("$level", warning.Field); Assert.Equal("$level", warning.FieldCode);
            Assert.Equal(location, warning.Location); Assert.Equal(sheet.Source.Part, warning.Part);
        });
    }

    [Fact]
    public void Schedule_Independent_regions_keep_source_order_and_distinct_owners_without_static_rows()
    {
        var sheet = Plan(Create("$l(10).Items", null, "$l(10).Items")).Sheets.Single();
        using var services = _fixture.Services(); using var json = JsonDocument.Parse("""{"Items":[{"$level":10},{"$level":10}]}""");
        var planner = services.GetRequiredService<XlsxTemplatePlanner>(); var context = new XlsxOperationContext(new MergeOptions());
        var schedule = planner.ScheduleSheet(sheet, planner.BindSheet(sheet, json.RootElement, context), context);
        Assert.Equal(new[] { 0, 1 }, schedule.Regions.Select(region => region.RegionId));
        Assert.Equal(new[] { 3, 3, 5, 5 }, schedule.Regions.SelectMany(region => region.Blocks).Select(block => block.Source.FirstRow));
        Assert.Equal(new[] { 0, 1, 0, 1 }, schedule.Regions.SelectMany(region => region.Blocks).Select(block => Assert.Single(block.Owner.LogicalRows)));
        Assert.NotSame(schedule.Regions[0].Blocks[0].Owner, schedule.Regions[1].Blocks[0].Owner);
        var staticSheet = Plan(Create((string?)null)).Sheets.Single();
        Assert.Empty(planner.ScheduleSheet(staticSheet, planner.BindSheet(staticSheet, json.RootElement, context), context).Regions);
    }

    [Fact]
    public void Schedule_Geometry_consumes_only_supplied_blocks_order_and_shared_owners()
    {
        var fixture = new XlsxTwoAxisFixture();
        var sheet = fixture.Plan(fixture.Create((_, part) => part.Worksheet!.Descendants<Row>()
            .Single(row => row.RowIndex!.Value == 11).PrependChild(_fixture.Text("A11", "$l(10).Items")))).Sheets.Single();
        using var services = _fixture.Services(); using var json = JsonDocument.Parse(fixture.Data(3, 3));
        var planner = services.GetRequiredService<XlsxTemplatePlanner>(); var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
        var context = new XlsxOperationContext(new MergeOptions()); var data = planner.BindSheet(sheet, json.RootElement, context);
        var scheduled = planner.ScheduleSheet(sheet, data, context);
        // Select a subset in a different order: geometry cannot recover this sequence from data levels.
        var selected = ImmutableArray.Create(scheduled.Regions[0].Blocks[2], scheduled.Regions[0].Blocks[0]);
        var schedule = scheduled with { Regions = [new(0, selected)] };
        var output = geometry.ExpansionGeometry(sheet, schedule, geometry.HorizontalSpans(sheet, data));
        Assert.Equal(new XlsxRectangle(new(1, 1), new(8, 11)), output.OutputArea);
        Assert.Equal(new[] { (7, 9, 1), (10, 11, 4), (10, 11, 6), (12, 12, 8) },
            output.RowSpans.Select(span => (span.First, span.Last, span.OutputFirst)));
        Assert.Same(selected[0].Owner, output.RowSpans[1].Vertical);
        Assert.Same(selected[1].Owner, output.RowSpans[2].Vertical);
        var cells = geometry.EmitCells(sheet, output).Rows.SelectMany(row => row.Cells).Where(cell => cell.Id.Vertical is not null).ToArray();
        Assert.All(cells.Where(cell => cell.Output.Row is 4 or 5), cell => Assert.Same(selected[0].Owner, cell.Id.Vertical));
        Assert.All(cells.Where(cell => cell.Output.Row is 6 or 7), cell => Assert.Same(selected[1].Owner, cell.Id.Vertical));
        Assert.Empty(context.Warnings);
    }

    [Fact]
    public void Multirow_Flat_output_preserves_tuple_values_row_styles_merged_names_and_static_tail()
    {
        var fixture = new XlsxTwoAxisFixture();
        var bytes = fixture.Create((workbook, part) =>
        {
            workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()),
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })), new Borders(new Border()),
                new CellStyleFormats(new CellFormat()), new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 14 }));
            var row = part.Worksheet!.Descendants<Row>().Single(row => row.RowIndex!.Value == 11);
            row.Height = 31; row.CustomHeight = true;
            row.RemoveAllChildren(); row.Append(_fixture.Text("A11", "$level(10).Items"),
                new Cell { CellReference = "B11", StyleIndex = 1 }, _fixture.Text("E11", "{Price}"));
            part.Worksheet.GetFirstChild<MergeCells>()!.Append(new MergeCell { Reference = "B10:B11" });
        });
        var result = _fixture.Merge(bytes, fixture.Data(3, 2));
        using var document = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(document)); Assert.Empty(result.Warnings);
        var sheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal("A1:K8", sheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal(new[] { "C1:F1", "G1:J1", "A4:A5", "A6:A7" }, sheet.Descendants<MergeCell>().Select(merge => merge.Reference!.Value));
        Assert.Equal("Row 0", _fixture.ReadCellText(result.Document, "A4")); Assert.Equal("Row 1", _fixture.ReadCellText(result.Document, "A6"));
        Assert.Equal("tail", _fixture.ReadCellText(result.Document, "A8"));
        var cells = sheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.All(new[] { "A5", "A7" }, address => Assert.Equal(1U, cells[address].StyleIndex!.Value));
        Assert.Equal(new[] { 18d, 31d, 18d, 31d }, sheet.Descendants<Row>().Where(row => row.RowIndex!.Value is >= 4 and <= 7).Select(row => row.Height!.Value));
        foreach (var outputRow in new[] { 4, 5, 6, 7 })
        {
            var start = outputRow < 6 ? 100 : 200;
            Assert.Equal(Enumerable.Range(start, 6).Select(value => value.ToString()),
                new[] { "D", "E", "F", "H", "I", "J" }.Select(column => cells[column + outputRow].CellValue!.Text));
        }
    }

    [Fact]
    public void Multirow_Parsed_nested_child_reuses_bound_owner_and_tuple_records_without_traversal()
    {
        var fixture = new XlsxTwoAxisFixture();
        var sheet = fixture.Plan(fixture.Create((_, part) =>
        {
            var rows = part.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
            rows[9].PrependChild(_fixture.Text("A9", "$l(10).Items"));
            rows[10].Elements<Cell>().First().InlineString!.Text!.Text = "$l(40).Items";
            rows[11].PrependChild(_fixture.Text("A11", "$level(40).Items"));
            rows[11].Append(_fixture.Text("E11", "{Price}"));
            rows[12].PrependChild(_fixture.Text("A12", "$l(10).Items"));
        })).Sheets.Single();
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(new
        {
            Bids = new[] { new { Company = "A" }, new { Company = "B" } },
            Criteria = new[] { new { Name = "X" }, new { Name = "Y" }, new { Name = "Z" } },
            Items = Enumerable.Range(0, 3).SelectMany(item => Enumerable.Range(0, 6).Select(tuple =>
                new Dictionary<string, object> { ["$level"] = item == 0 ? 10 : 40, ["$path"] = "opaque" + item,
                    ["Name"] = "item" + item, ["Quantity"] = 1, ["Price"] = item * 100 + tuple }))
        }));
        using var services = _fixture.Services(); var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var context = new XlsxOperationContext(new MergeOptions()); var data = planner.BindSheet(sheet, json.RootElement, context);
        var region = Assert.Single(sheet.Regions); var child = Assert.Single(region.Root.Children);
        AssertBlock(child.Body, 10, 11);
        Assert.Equal(new[] { 0, 6, 12 }, data.Rows[0].Select(row => row.FirstIndex));
        Assert.Equal(new[] { 6, 12 }, data.Rows[0][0].ChildFirstIndices);
        var fields = sheet.Cells.Where(cell => cell.Source.Id.Address.Column == 5 &&
            cell.Source.Id.Address.Row >= child.Body!.FirstRow && cell.Source.Id.Address.Row <= child.Body.LastRow).ToArray();
        Assert.Equal(new[] { 10, 11 }, fields.Select(cell => cell.Source.Id.Address.Row));
        foreach (var row in data.Rows[0].Skip(1))
        {
            Assert.Same(data.Rows[0][0], row.Parent); Assert.Equal(6, row.Records.Length);
            for (var tuple = 0; tuple < 6; tuple++)
            {
                var scope = planner.RowScope(data.Root, row, region.Collection, tuple);
                var expected = (row.FirstIndex / 6 * 100 + tuple).ToString();
                Assert.All(fields, field => Assert.Equal(new XlsxPreparedValue("n", expected), planner.PrepareValue(field, scope, context)));
            }
        }
        Assert.Empty(context.Errors); Assert.Empty(context.Warnings);
        Assert.Contains("M3.4", Assert.Throws<NotSupportedException>(() => planner.ScheduleSheet(sheet, data, context)).Message);
    }
}
