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
    [InlineData(2, 2)] [InlineData(0, 2)] [InlineData(2, 0)]
    public void HeadersFooters_Render_complete_two_row_children_with_shared_four_record_tuples_and_authored_fidelity(int bids, int criteria)
    {
        var source = NestedTupleTemplate();
        var arity = Math.Max(1, bids * criteria);
        var json = JsonSerializer.Serialize(new
        {
            Bids = Enumerable.Range(0, bids).Select(index => new { Company = "Bid " + index }),
            Criteria = Enumerable.Range(0, criteria).Select(index => new { Name = "Criterion " + index }),
            Items = new[] { (10, "P", 0), (40, "C1", 100), (40, "C2", 200), (10, "Q", 0), (40, "C3", 300) }
                .SelectMany((item, index) => Enumerable.Range(0, arity).Select(tuple => new Dictionary<string, object>
                {
                    ["$level"] = item.Item1, ["$path"] = "opaque" + index,
                    ["Name"] = tuple == 0 ? item.Item2 : "tuple name must not become ordinary name",
                    ["Quantity"] = tuple == 0 ? index + 1 : 999,
                    ["Price"] = item.Item3 + tuple + 1
                }))
        });
        var result = RenderNested(source, json);
        var schedule = Assert.Single(Assert.Single(result.Schedules).Regions).Blocks;
        Assert.Equal(new[] { XlsxLogicalBlockRole.Header, XlsxLogicalBlockRole.Body, XlsxLogicalBlockRole.Body,
            XlsxLogicalBlockRole.Footer, XlsxLogicalBlockRole.Header, XlsxLogicalBlockRole.Body, XlsxLogicalBlockRole.Footer }, schedule.Select(block => block.Role));
        var expectedPaths = new[] { new[] { 0 }, new[] { 0, arity }, new[] { 0, arity * 2 }, new[] { 0 },
            new[] { arity * 3 }, new[] { arity * 3, arity * 4 }, new[] { arity * 3 } };
        for (var index = 0; index < schedule.Length; index++) Assert.Equal(expectedPaths[index], schedule[index].Owner.LogicalRows);
        Assert.Same(schedule[0].Owner, schedule[3].Owner); Assert.Same(schedule[4].Owner, schedule[6].Owner);
        var region = Assert.Single(result.Template.Sheets[0].Regions);
        Assert.Same(region.Root.Header, schedule[0].Source); Assert.Same(region.Root.Footer, schedule[3].Source);
        Assert.All(new[] { 1, 2, 5 }, index => Assert.Same(region.Root.Children[0].Body, schedule[index].Source));
        Assert.Equal(arity * 5, result.Bound[0].Rows[0].Sum(row => row.Records.Length));
        Assert.All(result.Bound[0].Rows[0], row => Assert.Equal(arity, row.Records.Length));
        var verticalSpans = result.Output.Sheets[0].Geometry.RowSpans.Where(span => span.Vertical is not null).ToArray();
        Assert.Equal(new[] { (9, 9, 3), (10, 11, 4), (10, 11, 6), (12, 12, 8), (9, 9, 9), (10, 11, 10), (12, 12, 12) },
            verticalSpans.Select(span => (span.First, span.Last, span.OutputFirst)));
        for (var index = 0; index < schedule.Length; index++) Assert.Same(schedule[index].Owner, verticalSpans[index].Vertical);
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package)); Assert.Empty(result.Context.Errors); Assert.Empty(result.Context.Warnings);
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal("A1:" + (bids == 0 ? "C" : criteria == 0 ? "E" : "I") + "12", worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal(new[] { "H:P", "C1", "C2", "F:P", "H:Q", "C3", "F:Q" },
            new[] { 3, 4, 6, 8, 9, 10, 12 }.Select(row => _fixture.ReadCellText(result.Document, "A" + row)));
        var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal(new[] { "P", "P", "Q" }, new[] { 5, 7, 11 }.Select(row => _fixture.ReadCellText(result.Document, "B" + row)));
        Assert.Equal(new[] { "2", "3", "5" }, new[] { 4, 6, 10 }.Select(row => cells["B" + row].CellValue!.Text));
        Assert.Equal(new[] { 21d, 18d, 31d, 18d, 31d, 27d, 21d, 18d, 31d, 27d },
            worksheet.Descendants<Row>().Where(row => row.RowIndex!.Value >= 3).Select(row => row.Height!.Value));
        Assert.All(new[] { 4, 6, 10 }, row => Assert.Equal(1U, cells["A" + row].StyleIndex!.Value));
        Assert.All(new[] { 5, 7, 11 }, row => Assert.Equal(2U, cells["A" + row].StyleIndex!.Value));
        var expectedMerges = new List<string>();
        if (bids > 0 && criteria > 0) expectedMerges.AddRange(new[] { "C1:E1", "F1:H1" });
        expectedMerges.AddRange(new[] { "A3:B3", "A9:B9", "A4:A5", "A6:A7", "A10:A11", "A8:B8", "A12:B12" });
        Assert.Equal(expectedMerges, worksheet.Descendants<MergeCell>().Select(merge => merge.Reference!.Value));
        if (bids > 0 && criteria > 0)
            foreach (var (row, firstPrice) in new[] { (4, 101), (6, 201), (10, 301) })
                foreach (var physicalRow in new[] { row, row + 1 })
                {
                    Assert.Equal(Enumerable.Range(firstPrice, 4).Select(price => price.ToString()),
                        new[] { "D", "E", "G", "H" }.Select(column => cells[column + physicalRow].CellValue!.Text));
                    Assert.All(new[] { "D", "E", "G", "H" }, column =>
                        Assert.Equal(physicalRow == row ? 3U : 4U, cells[column + physicalRow].StyleIndex!.Value));
                }
        if (bids == 2 && criteria == 2 && Environment.GetEnvironmentVariable("MONOREPORT_XLSX_EVIDENCE") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, "nested-two-row-template.xlsx"), source);
            File.WriteAllText(Path.Combine(directory, "nested-two-row-data.json"), json);
            File.WriteAllBytes(Path.Combine(directory, "nested-two-row-generated.xlsx"), result.Document);
        }
    }

    [Theory]
    [InlineData("10,40,10", "10:Company,20:Department X,40:Ivan,40:Anna,20:Department Y,40:Ivan", "Company,Ivan/Department X,Anna/Department X,Ivan/Department Y,Company", "0;0,1,2;0,1,3;0,4,5;0")]
    [InlineData("10,20,40,20,10", "10:Company,40:Ivan,40:Anna", "Company,Ivan/Company,Anna/Company,Company", "0;0,1;0,2;0")]
    [InlineData("10,20,40,20,10", "40:Ivan,40:Anna", "Ivan/,Anna/", "0;1")]
    [InlineData("90,40,10", "10:P,40:C,90:G", "G/C,C/P,P", "0,1,2;0,1;0")]
    [InlineData("10,40,20", "10:P,20:X,40:Y", "P,Y/X,X/P", "0;0,1,2;0,1")]
    [InlineData("10,40,20", "10:P,40:Y,20:X", "P,Y/P,X/P", "0;0,1;0,2")]
    [InlineData("3,40", "5:A,3:B,40:C", "B,C/B", "1;1,2")]
    [InlineData("10,40,10", "", "", "")]
    public void HeadersFooters_Data_tree_controls_projection_even_when_template_ancestry_differs(string levels, string items, string expected, string owners)
    {
        var source = ProjectionTemplate(levels.Split(',').Select(int.Parse).ToArray());
        var json = JsonSerializer.Serialize(new { Items = items.Split(',', StringSplitOptions.RemoveEmptyEntries).Select(item =>
            new Dictionary<string, object> { ["$level"] = int.Parse(item.Split(':')[0]), ["Name"] = item.Split(':')[1] }) });
        var result = RenderNested(source, json);
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var addresses = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Select(cell => cell.CellReference!.Value!).ToArray();
        Assert.Equal(expected.Split(',', StringSplitOptions.RemoveEmptyEntries), addresses.Select(address => _fixture.ReadCellText(result.Document, address)));
        Assert.Equal(owners.Split(';', StringSplitOptions.RemoveEmptyEntries),
            result.Schedules[0].Regions[0].Blocks.Select(block => string.Join(',', block.Owner.LogicalRows)));
        Assert.Empty(result.Context.Errors);
        if (items == "40:Ivan,40:Anna")
        {
            Assert.Equal(2, result.Context.Warnings.Count);
            Assert.All(result.Context.Warnings, warning => Assert.Contains("data ancestor", warning.Message));
        }
        else Assert.Empty(result.Context.Warnings);
    }

    [Fact]
    public void HeadersFooters_Merge_projection_deduplicates_equal_owners_and_filters_absent_source_blocks()
    {
        var source = ProjectionTemplate([10, 40, 40, 10], (workbook, part) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$C$6";
            part.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference == "B5").Remove();
            part.Worksheet.Append(new MergeCells(new MergeCell { Reference = "B3:C3" }, new MergeCell { Reference = "B4:B5" }, new MergeCell { Reference = "B6:C6" }));
        });
        // The parent shares a Header/Footer owner, while standalone children have no source-row-3 image.
        var result = RenderNested(source, """{"Items":[{"$level":10,"Name":"P"},{"$level":40,"Name":"C"},{"$level":5,"Name":"Unrendered"},{"$level":40,"Name":"D"}]}""");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Assert.Equal(new[] { "A1:B1", "A2:A3", "A5:A6", "A4:B4" }, package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<MergeCell>().Select(merge => merge.Reference!.Value));
        Assert.Equal(new[] { "P", "C/P", "P", "D/Unrendered" }, new[] { "A1", "A2", "A4", "A5" }.Select(address => _fixture.ReadCellText(result.Document, address)));
        using var services = _fixture.Services();
        var geometry = services.GetRequiredService<XlsxGeometryPlanner>(); var sheet = result.Template.Sheets[0];
        var positions = result.Output.Sheets[0].Geometry;
        // Split a multirow source span into equivalent physical intervals with independently allocated equal paths.
        // Both rows of each merged child then contribute a candidate; structural equality must deduplicate them.
        positions = positions with { ContentOwners = [], DisplacedSources = [], RowSpans = positions.RowSpans.SelectMany(span =>
            Enumerable.Range(span.First, span.Last - span.First + 1).Select(row => new XlsxAxisSpan(row, row,
                span.OutputFirst + row - span.First, span.Vertical is null ? null :
                    new(span.Vertical.RegionId, ImmutableArray.CreateRange(span.Vertical.LogicalRows.ToArray()))))).ToImmutableArray() };
        var projected = geometry.ProjectMerges(sheet, geometry.EmitCells(sheet, positions));
        Assert.Equal(4, projected.Settings.Descendants().Count(element => element.Name.LocalName == "mergeCell"));
        Assert.Equal(4, projected.Geometry.ContentOwners.Length);
    }

    [Fact]
    public void HeadersFooters_Reject_cross_source_block_merges_even_if_their_output_happens_to_be_adjacent()
    {
        var source = ProjectionTemplate([10, 40, 10], (workbook, part) =>
        {
            part.Worksheet!.Descendants<Row>().Single(row => row.RowIndex!.Value == 4).Elements<Cell>().Single(cell => cell.CellReference == "B4").Remove();
            part.Worksheet.Append(new MergeCells(new MergeCell { Reference = "B3:B4" }));
        });
        NestedRenderResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = RenderNested(source, """{"Items":[{"$level":10,"Name":"P"},{"$level":40,"Name":"C"}]}"""));
        Assert.Null(result); Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, error.Code);
        Assert.Contains("one vertical source block", error.Message); Assert.Contains("B3:B4", error.Errors[0].Location);
    }

    [Fact]
    public void HeadersFooters_Duplicate_children_keep_exact_formula_and_merge_owners_after_crop()
    {
        var source = _fixture.Create((workbook, part) =>
        {
            var rows = part.Worksheet!.GetFirstChild<SheetData>()!; rows.RemoveAllChildren();
            rows.Append(new Row(_fixture.Text("A7", "$l(10).Items"), _fixture.Text("B7", "{Name}")) { RowIndex = 7 },
                new Row(_fixture.Text("A8", "$l(40).Items"), _fixture.Text("B8", "{Name}"), _fixture.Text("C8", "{Price}")) { RowIndex = 8 },
                new Row(_fixture.Text("A9", "$l(40).Items"), new Cell { CellReference = "C9", CellFormula = new CellFormula("C8") }) { RowIndex = 9 },
                new Row(_fixture.Text("A10", "$l(10).Items"), _fixture.Text("B10", "{Name}")) { RowIndex = 10 });
            part.Worksheet.Append(new MergeCells(new MergeCell { Reference = "B8:B9" }));
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$7:$C$10";
        });
        var result = RenderNested(source, """{"Items":[{"$level":10,"Name":"Company"},{"$level":20,"Name":"X"},{"$level":40,"Name":"Ivan","Price":11},{"$level":20,"Name":"Y"},{"$level":40,"Name":"Ivan","Price":22}]}""");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal(new[] { "A2:A3", "A4:A5" }, worksheet.Descendants<MergeCell>().Select(merge => merge.Reference!.Value));
        var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("11", cells["B2"].CellValue!.Text); Assert.Equal("22", cells["B4"].CellValue!.Text);
        Assert.Equal("B2", cells["B3"].CellFormula!.Text); Assert.Equal("B4", cells["B5"].CellFormula!.Text);
        Assert.Equal("Ivan", _fixture.ReadCellText(result.Document, "A2")); Assert.Equal("Ivan", _fixture.ReadCellText(result.Document, "A4"));
        Assert.Equal(new[] { "0", "0,1,2", "0,3,4", "0" }, result.Schedules[0].Regions[0].Blocks.Select(block => string.Join(',', block.Owner.LogicalRows)));
    }

    [Fact]
    public void HeadersFooters_Surviving_formula_cannot_borrow_an_absent_block_operand_and_returns_no_workbook()
    {
        var source = ProjectionTemplate([10, 40, 10], (_, part) =>
        {
            var footer = part.Worksheet!.Descendants<Row>().Single(row => row.RowIndex!.Value == 5);
            footer.Elements<Cell>().Single(cell => cell.CellReference == "B5").Remove();
            footer.Append(new Cell { CellReference = "B5", CellFormula = new CellFormula("B4") });
        });
        NestedRenderResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = RenderNested(source, """{"Items":[{"$level":10,"Name":"P"}]}"""));
        Assert.Null(result); var issue = Assert.Single(error.Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code); Assert.Contains("B5", issue.Location); Assert.Contains("B4", issue.Message);
    }

    [Fact]
    public void Schedule_Mismatched_sheet_or_region_order_is_an_internal_invariant_failure()
    {
        var sheet = Plan(Create("$l(10).Items", null, "$l(10).Items")).Sheets[0];
        using var services = _fixture.Services(); using var json = JsonDocument.Parse("""{"Items":[]}""");
        var planner = services.GetRequiredService<XlsxTemplatePlanner>(); var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
        var context = new XlsxOperationContext(new MergeOptions()); var bound = planner.BindSheet(sheet, json.RootElement, context);
        var schedule = planner.ScheduleSheet(sheet, bound, context); var columns = geometry.HorizontalSpans(sheet, bound);
        Assert.Throws<InvalidOperationException>(() => geometry.ExpansionGeometry(sheet, schedule with { SheetId = 999 }, columns));
        Assert.Throws<InvalidOperationException>(() => geometry.ExpansionGeometry(sheet, schedule with { Regions = schedule.Regions.Reverse().ToImmutableArray() }, columns));
    }

    private byte[] NestedTupleTemplate() => new XlsxTwoAxisFixture().Create((workbook, part) =>
    {
        workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()),
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })), new Borders(new Border()),
            new CellStyleFormats(new CellFormat()), new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 1 },
                new CellFormat { NumberFormatId = 2 }, new CellFormat { NumberFormatId = 3 }, new CellFormat { NumberFormatId = 4 }));
        var rows = part.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
        rows[9].RemoveAllChildren(); rows[9].Height = 21; rows[9].CustomHeight = true;
        rows[9].Append(_fixture.Text("A9", "$l(10).Items"), _fixture.Text("B9", "H:{Name}"));
        rows[10].Elements<Cell>().Single(cell => cell.CellReference == "A10").InlineString!.Text!.Text = "$l(40).Items";
        rows[10].Elements<Cell>().Single(cell => cell.CellReference == "B10").StyleIndex = 1;
        rows[10].Elements<Cell>().Single(cell => cell.CellReference == "E10").StyleIndex = 3;
        rows[11].RemoveAllChildren(); rows[11].Height = 31; rows[11].CustomHeight = true;
        var price = _fixture.Text("E11", "{Price}"); price.StyleIndex = 4;
        rows[11].Append(_fixture.Text("A11", "$l(40).Items"), new Cell { CellReference = "B11", StyleIndex = 2 }, _fixture.Text("C11", "{../Name}"), price);
        rows[12].RemoveAllChildren(); rows[12].Height = 27; rows[12].CustomHeight = true;
        rows[12].Append(_fixture.Text("A12", "$l(10).Items"), _fixture.Text("B12", "F:{Name}"));
        part.Worksheet.GetFirstChild<MergeCells>()!.Append(new MergeCell { Reference = "B9:C9" }, new MergeCell { Reference = "B10:B11" }, new MergeCell { Reference = "B12:C12" });
    });

    private byte[] ProjectionTemplate(int[] levels, Action<WorkbookPart, WorksheetPart>? customize = null) => _fixture.Create((workbook, part) =>
    {
        var rows = part.Worksheet!.GetFirstChild<SheetData>()!; rows.RemoveAllChildren();
        for (var index = 0; index < levels.Length; index++)
        {
            var row = index + 3;
            // Roots use no ancestor lookup; every nonroot field reports its actual data parent.
            rows.Append(new Row(_fixture.Text("A" + row, "$l(" + levels[index] + ").Items"),
                _fixture.Text("B" + row, levels[index] == levels.Min() ? "{Name}" : "{Name}/{../Name}")) { RowIndex = (uint)row });
        }
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$B$" + (levels.Length + 2);
        customize?.Invoke(workbook, part);
    });

    private NestedRenderResult RenderNested(byte[] source, string json)
    {
        using var services = _fixture.Services(); using var data = JsonDocument.Parse(json);
        var context = new XlsxOperationContext(new MergeOptions());
        using var stream = new MemoryStream(); stream.Write(source); stream.Position = 0;
        XlsxTemplatePlan template; XlsxWorkbookPlan output;
        ImmutableArray<XlsxBoundSheet> bound; ImmutableArray<XlsxLogicalSheetSchedule> schedules;
        using (var package = SpreadsheetDocument.Open(stream, true))
        {
            var planner = services.GetRequiredService<XlsxTemplatePlanner>(); var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
            template = planner.Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
            bound = template.Sheets.Select(sheet => planner.BindSheet(sheet, data.RootElement, context)).ToImmutableArray();
            schedules = template.Sheets.Select((sheet, index) => planner.ScheduleSheet(sheet, bound[index], context)).ToImmutableArray();
            output = geometry.PlanReports(template, bound, schedules, context);
            output = output with { Sheets = output.Sheets.Select((sheet, index) => planner.BindInstances(template.Sheets[index], sheet, bound[index], context)).ToImmutableArray() };
            output = services.GetRequiredService<XlsxFormulaPlanner>().Finalize(template, output, context);
            services.GetRequiredService<XlsxDocumentRenderer>().Render(package, output, context);
        }
        return new(stream.ToArray(), template, bound, schedules, output, context);
    }

    private sealed record NestedRenderResult(byte[] Document, XlsxTemplatePlan Template, ImmutableArray<XlsxBoundSheet> Bound,
        ImmutableArray<XlsxLogicalSheetSchedule> Schedules, XlsxWorkbookPlan Output, XlsxOperationContext Context);
}
