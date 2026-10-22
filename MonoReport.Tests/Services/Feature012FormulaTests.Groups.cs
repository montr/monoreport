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

public sealed partial class Feature012FormulaTests
{
    [Theory]
    [InlineData("", "")] [InlineData("$", "$")] [InlineData("$", "")] [InlineData("", "$")]
    public void Groups_Headers_and_footers_select_full_data_descendants_without_next_group_leakage(string firstFlag, string lastFlag)
    {
        var source = GroupTemplate(firstFlag, lastFlag);
        var result = RenderGroups(source, GroupData());
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal("A1:C20", worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal("'Report'!$A$1:$C$20", package.WorkbookPart.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text);
        var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("COUNT(B2:B19)", cells["B1"].CellFormula!.Text);
        Assert.Equal("SUM(B2:B19)", cells["B20"].CellFormula!.Text);
        foreach (var (header, footer, first, last) in new[] { (2, 13, 3, 12), (14, 19, 15, 18) })
            foreach (var row in new[] { header, footer })
                Assert.Equal($"SUBTOTAL(9,B{firstFlag}{first}:B{lastFlag}{last})", cells["B" + row].CellFormula!.Text);
        foreach (var (header, footer, first, last) in new[] { (3, 8, 4, 7), (9, 12, 10, 11), (15, 18, 16, 17) })
        {
            Assert.Equal($"SUM(B{firstFlag}{first}:B{lastFlag}{last})", cells["B" + header].CellFormula!.Text);
            Assert.Equal($"SUBTOTAL(9,B{firstFlag}{first}:B{lastFlag}{last})", cells["B" + footer].CellFormula!.Text);
            Assert.Equal($"COUNTA(A{first}:A{last})", cells["C" + header].CellFormula!.Text);
        }
        Assert.Equal(new[] { "11", "22", "33", "99" }, new[] { 4, 6, 10, 16 }.Select(row => cells["B" + row].CellValue!.Text));
        Assert.Equal(new[] { "B4*2", "B6*2", "B10*2", "B16*2" }, new[] { 5, 7, 11, 17 }.Select(row => cells["B" + row].CellFormula!.Text));
        Assert.Equal(new[] { "A4:A5", "A6:A7", "A10:A11", "A16:A17" }, worksheet.Descendants<MergeCell>().Select(merge => merge.Reference!.Value));
        var placements = result.Output.Sheets[0].Rows.SelectMany(row => row.Cells).Where(cell => cell.Formula is not null).ToDictionary(cell => cell.Output.ToString());
        Assert.Equal(new[] { 0 }, placements["B2"].Id.Vertical!.LogicalRows);
        Assert.Equal(placements["B2"].Id.Vertical, placements["B13"].Id.Vertical);
        Assert.Equal(new[] { 0, 1, 2 }, placements["B3"].Id.Vertical!.LogicalRows);
        Assert.Equal(new[] { 0, 1, 2, 3 }, placements["B5"].Id.Vertical!.LogicalRows);
        Assert.Equal(new[] { 0, 5, 6 }, placements["B9"].Id.Vertical!.LogicalRows);
        Assert.Equal(new[] { 8, 9, 10 }, placements["B15"].Id.Vertical!.LogicalRows);
        Assert.Null(placements["B1"].Id.Vertical); Assert.Null(placements["B20"].Id.Vertical);
        Assert.All(cells.Values.Where(cell => cell.CellFormula is not null), cell => Assert.Null(cell.CellValue));
        Assert.Empty(result.Context.Errors); Assert.Empty(result.Context.Warnings);
        if (firstFlag.Length == 0 && lastFlag.Length == 0) ExportGroups(source, GroupData(), result.Document);
    }

    [Theory]
    [InlineData("C10:C10", false)] // Multirow child gaps cannot be filled with the child's formula row.
    [InlineData("C10:C11", true)] // Separate team headers/footers cannot be filled into a company range.
    [InlineData("B10:B10", false)] // Retained merged-name owners are still disjoint between children.
    [InlineData("C10", false)] // A scalar operand must never acquire descendant-range behavior.
    [InlineData("$C$10", false)] // Absolute scalar anchors cannot select one arbitrary child.
    [InlineData("A1:A1", false)]
    public void Groups_Unsafe_descendant_operands_reject_without_bounding_or_returning_a_workbook(string operand, bool company)
    {
        var source = GroupTemplate(customize: (_, part) =>
        {
            var target = part.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference == (company ? "C8" : "C9"));
            target.CellFormula!.Text = "SUM(" + operand + ")";
        });
        var original = source.ToArray(); GroupRenderResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = RenderGroups(source, GroupData()));
        Assert.Null(result); Assert.Equal(original, source);
        Assert.All(error.Errors, issue => Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code));
        var issue = error.Errors.First(issue => issue.Location!.EndsWith(company ? "C8" : "C9", StringComparison.Ordinal));
        Assert.Contains(operand, issue.Message); Assert.Contains("retained contiguous A1 range", issue.Message);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void Groups_Empty_own_descendants_cannot_borrow_from_the_next_group_or_become_zero(bool absentSourceLevel)
    {
        var source = GroupTemplate();
        var items = absentSourceLevel
            ? """[{"$level":10,"Name":"P"},{"$level":30,"Name":"Empty team"},{"$level":50,"Name":"Unrendered child"},{"$level":10,"Name":"Q"},{"$level":30,"Name":"Team"},{"$level":40,"Name":"Ivan","Price":99}]"""
            : """[{"$level":10,"Name":"P"},{"$level":30,"Name":"Empty team"},{"$level":10,"Name":"Q"},{"$level":30,"Name":"Team"},{"$level":40,"Name":"Ivan","Price":99}]""";
        GroupRenderResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = RenderGroups(source, "{\"Items\":" + items + "}"));
        Assert.Null(result); Assert.Contains(error.Errors, issue => issue.Code == DiagnosticCodes.XlsxUnsafeFormula && issue.Location!.Contains("C9") && issue.Message.Contains("C10:C11"));
    }

    [Fact]
    public void Groups_Ranges_can_include_current_owner_roles_and_keep_explicit_union_operands_separate()
    {
        var source = GroupTemplate(customize: (_, part) =>
        {
            var rows = part.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
            rows[8].Append(GroupFormula("D8", "COUNT(C8:C13)"));
            rows[13].Append(GroupFormula("D13", "SUM(C8:C8,C9:C12)"));
        });
        var result = RenderGroups(source, GroupData());
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("COUNT(B2:B13)", cells["C2"].CellFormula!.Text);
        Assert.Equal("COUNT(B14:B19)", cells["C14"].CellFormula!.Text);
        Assert.Equal("SUM(B2:B2,B3:B12)", cells["C13"].CellFormula!.Text);
        Assert.Equal("SUM(B14:B14,B15:B18)", cells["C19"].CellFormula!.Text);
    }

    [Fact]
    public void Groups_Empty_input_removes_group_formulas_with_their_owners()
    {
        var result = RenderGroups(SimpleGroupWorkbook(), """{"Items":[]}""");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Assert.Empty(package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>());
        Assert.Empty(result.Context.Errors); Assert.Empty(result.Context.Warnings);
    }

    [Fact]
    public void Groups_Footer_only_projection_uses_actual_data_parenthood_instead_of_source_ancestry()
    {
        var source = SimpleGroupWorkbook(footerOnly: true);
        var result = RenderGroups(source, """{"Items":[{"$level":10,"Name":"P"},{"$level":20,"Name":"Y"},{"$level":40,"Name":"Ivan","Price":11},{"$level":40,"Name":"Ivan","Price":22}]}""");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("SUM(B2:B3)", cells["B4"].CellFormula!.Text);
        Assert.Equal(new[] { 0, 1 }, result.Output.Sheets[0].Rows.SelectMany(row => row.Cells).Single(cell => cell.Output == new XlsxAddress(4, 2)).Id.Vertical!.LogicalRows);
        GroupRenderResult? rejected = null;
        var error = Assert.Throws<TemplateException>(() => rejected = RenderGroups(source, """{"Items":[{"$level":10,"Name":"P"},{"$level":40,"Name":"Ivan","Price":11},{"$level":20,"Name":"Y"}]}"""));
        Assert.Null(rejected); Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code); Assert.Contains("C9", error.Errors[0].Location);
    }

    [Fact]
    public void Groups_Independent_regions_and_worksheets_keep_local_owners_and_global_separator_totals()
    {
        var source = SimpleGroupWorkbook(multiple: true);
        var result = RenderGroups(source, """{"Items":[{"$level":10,"Name":"P"},{"$level":40,"Name":"Ivan","Price":11},{"$level":40,"Name":"Ivan","Price":22},{"$level":10,"Name":"Q"},{"$level":40,"Name":"Ivan","Price":99}],"Other":[{"$level":10,"Name":"R"},{"$level":40,"Name":"Ivan","Price":7}]}""");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var worksheets = package.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Select(sheet => ((WorksheetPart)package.WorkbookPart.GetPartById(sheet.Id!)).Worksheet!).ToArray();
        var cells = worksheets[0].Descendants<Cell>().Where(cell => cell.CellFormula is not null).ToDictionary(cell => cell.CellReference!.Value!, cell => cell.CellFormula!.Text);
        Assert.Equal(new Dictionary<string, string> { ["B1"] = "SUM(B2:B12)", ["B2"] = "SUM(B3:B4)", ["B5"] = "SUBTOTAL(9,B3:B4)",
            ["B6"] = "SUM(B7:B7)", ["B8"] = "SUBTOTAL(9,B7:B7)", ["B9"] = "COUNT(B2:B12)",
            ["B10"] = "SUM(B11:B11)", ["B12"] = "SUBTOTAL(9,B11:B11)", ["B13"] = "COUNT(B2:B12)" }, cells);
        Assert.Equal(new[] { "SUM(B2:B2)", "SUBTOTAL(9,B2:B2)" }, worksheets[1].Descendants<Cell>().Where(cell => cell.CellFormula is not null).Select(cell => cell.CellFormula!.Text));
        Assert.Equal(new[] { "A1:B13", "A1:B3" }, worksheets.Select(sheet => sheet.GetFirstChild<SheetDimension>()!.Reference!.Value));
        var owners = result.Output.Sheets[0].Rows.SelectMany(row => row.Cells).Where(cell => cell.Formula is not null && cell.Id.Vertical is not null).ToArray();
        Assert.Equal(new[] { 0, 0, 0, 0, 1, 1 }, owners.Select(cell => cell.Id.Vertical!.RegionId));
        Assert.All(result.Output.Sheets[1].Rows.SelectMany(row => row.Cells), cell => Assert.Equal(7U, cell.Source.Id.SheetId));
        Assert.Empty(result.Context.Warnings);
    }

    [Theory]
    [InlineData(2, 2)] [InlineData(1, 1)]
    public void Groups_Horizontal_tuple_and_parent_ranges_keep_vertical_group_ownership(int bids, int criteria)
    {
        var source = TupleGroupTemplate(); var arity = bids * criteria;
        var json = TupleGroupData(bids, criteria);
        var result = RenderGroups(source, json);
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        var leafColumns = bids == 2 ? new[] { "D", "E", "G", "H" } : new[] { "D" };
        foreach (var column in leafColumns)
        {
            Assert.Equal($"SUM({column}4:{column}7)", cells[column + "3"].CellFormula!.Text);
            Assert.Equal($"SUBTOTAL(9,{column}4:{column}7)", cells[column + "8"].CellFormula!.Text);
            Assert.Equal($"SUM({column}10:{column}11)", cells[column + "9"].CellFormula!.Text);
            Assert.Equal($"SUBTOTAL(9,{column}10:{column}11)", cells[column + "12"].CellFormula!.Text);
            Assert.Equal(column + "4*2", cells[column + "5"].CellFormula!.Text);
        }
        Assert.Equal(criteria == 2 ? "SUM(D4:E7)" : "SUM(D4:D7)", cells["C3"].CellFormula!.Text);
        if (bids == 2) Assert.Equal("SUM(G4:H7)", cells["F3"].CellFormula!.Text);
        var placements = result.Output.Sheets[0].Rows.SelectMany(row => row.Cells).ToArray();
        Assert.Equal(new[] { 0, arity, arity * 2 }, placements.Single(cell => cell.Output == new XlsxAddress(5, 4)).Id.Vertical!.LogicalRows);
        Assert.Equal(new[] { arity * 4 }, placements.Single(cell => cell.Output == new XlsxAddress(9, 4)).Id.Vertical!.LogicalRows);
        Assert.Equal(Enumerable.Range(101, arity).Select(value => value.ToString()), leafColumns.Select(column => cells[column + "4"].CellValue!.Text));
        Assert.Equal(Enumerable.Range(201, arity).Select(value => value.ToString()), leafColumns.Select(column => cells[column + "6"].CellValue!.Text));
        Assert.Empty(result.Context.Warnings);
    }

    [Theory]
    [InlineData(2, 0)] [InlineData(0, 2)]
    public void Groups_Surviving_formulas_reject_removed_horizontal_descendants(int bids, int criteria)
    {
        var source = TupleGroupTemplate(); GroupRenderResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = RenderGroups(source, TupleGroupData(bids, criteria)));
        Assert.Null(result); Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
        Assert.Contains(bids == 0 ? "F9" : "D9", error.Errors[0].Location);
        Assert.Contains(bids == 0 ? "D10:E11" : "E10:E11", error.Errors[0].Message);
    }

    [Theory]
    [InlineData("E10:E11")] // The two companies have a parent-only column between their leaf columns.
    [InlineData("$E10:E11")] // Mixed column flags select different horizontal owner images.
    public void Groups_Descendant_selection_preserves_horizontal_disjoint_and_mixed_flag_safety(string operand)
    {
        var source = TupleGroupTemplate();
        source = _fixture.RewritePart(source, "xl/worksheets/sheet1.xml", text =>
        {
            var xml = System.Xml.Linq.XDocument.Parse(text); System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            xml.Descendants(ns + "c").Single(cell => (string?)cell.Attribute("r") == (operand.StartsWith('$') ? "E9" : "F9")).Element(ns + "f")!.Value = "SUM(" + operand + ")";
            return xml.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        });
        GroupRenderResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = RenderGroups(source, TupleGroupData(2, 2)));
        Assert.Null(result); Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
        Assert.Contains(operand.StartsWith('$') ? "E9" : "F9", error.Errors[0].Location);
        Assert.Contains(operand.StartsWith('$') ? "Mixed endpoint copy flags" : "disjoint", error.Errors[0].Message);
    }

    private byte[] GroupTemplate(string firstFlag = "", string lastFlag = "", Action<WorkbookPart, WorksheetPart>? customize = null) => _fixture.Create((workbook, part) =>
    {
        string Range(int first, int last) => $"C{firstFlag}{first}:C{lastFlag}{last}";
        part.Worksheet = new Worksheet(new SheetData(
            new Row(_fixture.Text("B7", "top"), GroupFormula("C7", "COUNT(C8:C13)")) { RowIndex = 7 },
            new Row(_fixture.Text("A8", "$l(10).Items"), _fixture.Text("B8", "H:{Name}"), GroupFormula("C8", "SUBTOTAL(9," + Range(9, 12) + ")")) { RowIndex = 8 },
            new Row(_fixture.Text("A9", "$l(30).Items"), _fixture.Text("B9", "H:{Name}"), GroupFormula("C9", "SUM(" + Range(10, 11) + ")"), GroupFormula("D9", "COUNTA(B10:B11)")) { RowIndex = 9 },
            new Row(_fixture.Text("A10", "$l(40).Items"), _fixture.Text("B10", "{Name}"), _fixture.Text("C10", "{Price}")) { RowIndex = 10 },
            new Row(_fixture.Text("A11", "$l(40).Items"), GroupFormula("C11", "C10*2")) { RowIndex = 11 },
            new Row(_fixture.Text("A12", "$l(30).Items"), _fixture.Text("B12", "F:{Name}"), GroupFormula("C12", "SUBTOTAL(9," + Range(10, 11) + ")")) { RowIndex = 12 },
            new Row(_fixture.Text("A13", "$l(10).Items"), _fixture.Text("B13", "F:{Name}"), GroupFormula("C13", "SUBTOTAL(9," + Range(9, 12) + ")")) { RowIndex = 13 },
            new Row(_fixture.Text("B14", "bottom"), GroupFormula("C14", "SUM(C8:C13)")) { RowIndex = 14 }),
            new MergeCells(new MergeCell { Reference = "B10:B11" }));
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$7:$D$14";
        customize?.Invoke(workbook, part);
    });

    private string GroupData() => JsonSerializer.Serialize(new { Items = new[] { (10, "P", 0), (20, "Department X", 0), (30, "Team", 0), (40, "Ivan", 11), (40, "Ivan", 22),
        (20, "Department Y", 0), (30, "Team", 0), (40, "Ivan", 33), (10, "Q", 0), (20, "Department Z", 0), (30, "Team", 0), (40, "Ivan", 99) }
        .Select(item => new Dictionary<string, object> { ["$level"] = item.Item1, ["Name"] = item.Item2, ["Price"] = item.Item3 }) });

    private byte[] SimpleGroupWorkbook(bool footerOnly = false, bool multiple = false) => _fixture.Create((workbook, part) =>
    {
        var rows = new SheetData();
        void Region(SheetData target, int first, string collection, int firstColumn = 2)
        {
            var name = firstColumn == 2 ? "B" : "D"; var value = firstColumn == 2 ? "C" : "E";
            target.Append(new Row(_fixture.Text("A" + first, "$l(10)." + collection), _fixture.Text(name + first, "H:{Name}"),
                    GroupFormula(value + first, "SUM(" + value + (first + 1) + ":" + value + (first + 1) + ")")) { RowIndex = (uint)first },
                new Row(_fixture.Text("A" + (first + 1), "$l(40)." + collection), _fixture.Text(name + (first + 1), "{Name}"), _fixture.Text(value + (first + 1), "{Price}")) { RowIndex = (uint)first + 1 },
                new Row(_fixture.Text("A" + (first + 2), "$l(10)." + collection), _fixture.Text(name + (first + 2), "F:{Name}"),
                    GroupFormula(value + (first + 2), "SUBTOTAL(9," + value + (first + 1) + ":" + value + (first + 1) + ")")) { RowIndex = (uint)first + 2 });
        }
        if (multiple) rows.Append(new Row(GroupFormula("C6", "SUM(C7:C13)")) { RowIndex = 6 });
        Region(rows, 7, "Items");
        if (footerOnly)
        {
            rows.Elements<Row>().First().Elements<Cell>().Last().Remove();
            rows.Elements<Row>().Last().Elements<Cell>().First().InlineString!.Text!.Text = "$l(20).Items";
            rows.Elements<Row>().Last().Elements<Cell>().Last().CellFormula!.Text = "SUM(C8:C8)";
        }
        if (multiple)
        {
            rows.Append(new Row(_fixture.Text("B10", "separator"), GroupFormula("C10", "COUNT(C7:C13)")) { RowIndex = 10, Height = 0, CustomHeight = true });
            Region(rows, 11, "Other");
            rows.Append(new Row(GroupFormula("C14", "COUNT(C7:C13)")) { RowIndex = 14 });
        }
        part.Worksheet = new Worksheet(rows);
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = multiple ? "'Report'!$B$6:$C$14" : "'Report'!$B$7:$C$9";
        if (!multiple) return;
        var other = workbook.AddNewPart<WorksheetPart>(); var otherRows = new SheetData(); Region(otherRows, 5, "Other", 4);
        other.Worksheet = new Worksheet(otherRows);
        workbook.Workbook.Sheets!.Append(new Sheet { Name = "Independent", SheetId = 7, Id = workbook.GetIdOfPart(other) });
        workbook.Workbook.DefinedNames.Append(new DefinedName("'Independent'!$D$5:$E$7") { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
    });

    private byte[] TupleGroupTemplate() => new XlsxTwoAxisFixture().Create((_, part) =>
    {
        var rows = part.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
        rows[9].RemoveAllChildren(); rows[9].Append(_fixture.Text("A9", "$l(10).Items"), _fixture.Text("B9", "H:{Name}"),
            GroupFormula("D9", "SUM(E10:E11)"), GroupFormula("E9", "SUM(E10:E11)"), GroupFormula("F9", "SUM(E10:E11)"));
        rows[10].Elements<Cell>().First().InlineString!.Text!.Text = "$l(40).Items";
        rows[11].RemoveAllChildren(); rows[11].Append(_fixture.Text("A11", "$l(40).Items"), GroupFormula("E11", "E10*2"));
        rows[12].RemoveAllChildren(); rows[12].Append(_fixture.Text("A12", "$l(10).Items"), _fixture.Text("B12", "F:{Name}"), GroupFormula("E12", "SUBTOTAL(9,E10:E11)"));
        // The static trailing column must not bound over parent-only gaps for a positive Cartesian product.
        rows[9].Elements<Cell>().Last().CellFormula!.Text = "SUM(D10:E11)";
    });

    private string TupleGroupData(int bids, int criteria) => JsonSerializer.Serialize(new
    {
        Bids = Enumerable.Range(0, bids).Select(index => new { Company = "Bid " + index }), Criteria = Enumerable.Range(0, criteria).Select(index => new { Name = "Criterion " + index }),
        Items = new[] { (10, "P", 0), (20, "Department", 0), (40, "Ivan", 100), (40, "Ivan", 200), (10, "Q", 0), (40, "Ivan", 300) }
            .SelectMany((item, index) => Enumerable.Range(0, Math.Max(1, bids * criteria)).Select(tuple => new Dictionary<string, object>
            { ["$level"] = item.Item1, ["$path"] = "item" + index, ["Name"] = item.Item2, ["Price"] = item.Item3 + tuple + 1 }))
    });

    private Cell GroupFormula(string address, string text) => new() { CellReference = address, CellFormula = new CellFormula(text), CellValue = new CellValue("999999") };

    private GroupRenderResult RenderGroups(byte[] source, string json)
    {
        using var services = _fixture.Services(); using var data = JsonDocument.Parse(json);
        var context = new XlsxOperationContext(new MergeOptions());
        using var stream = new MemoryStream(); stream.Write(source); stream.Position = 0;
        XlsxWorkbookPlan output;
        using (var package = SpreadsheetDocument.Open(stream, true))
        {
            var planner = services.GetRequiredService<XlsxTemplatePlanner>(); var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
            var template = planner.Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
            var bound = template.Sheets.Select(sheet => planner.BindSheet(sheet, data.RootElement, context)).ToImmutableArray();
            var schedules = template.Sheets.Select((sheet, index) => planner.ScheduleSheet(sheet, bound[index], context)).ToImmutableArray();
            output = geometry.PlanReports(template, bound, schedules, context);
            output = output with { Sheets = output.Sheets.Select((sheet, index) => planner.BindInstances(template.Sheets[index], sheet, bound[index], context)).ToImmutableArray() };
            output = services.GetRequiredService<XlsxFormulaPlanner>().Finalize(template, output, context);
            services.GetRequiredService<XlsxDocumentRenderer>().Render(package, output, context);
        }
        return new(stream.ToArray(), output, context);
    }

    private void ExportGroups(byte[] source, string json, byte[] output)
    {
        if (Environment.GetEnvironmentVariable("MONOREPORT_XLSX_EVIDENCE") is not { Length: > 0 } directory) return;
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, "groups-template.xlsx"), source); File.WriteAllText(Path.Combine(directory, "groups-data.json"), json);
        File.WriteAllBytes(Path.Combine(directory, "groups-generated.xlsx"), output);
    }

    private sealed record GroupRenderResult(byte[] Document, XlsxWorkbookPlan Output, XlsxOperationContext Context);
}
