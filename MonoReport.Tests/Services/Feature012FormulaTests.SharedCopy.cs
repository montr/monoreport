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
    [Fact]
    public void SharedCopy_Followers_before_master_use_source_offsets_and_preserve_each_dollar_flag()
    {
        var bytes = SharedInventoryWorkbook();
        using var services = _fixture.Services();
        using var input = new MemoryStream(bytes);
        using var package = SpreadsheetDocument.Open(input, false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var context = new XlsxOperationContext(new MergeOptions());
        var source = services.GetRequiredService<XlsxPackageReader>().Read(package, context);
        var xml = source.Sheets[0].Xml.ToString();
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(source, context);
        var cells = plan.Sheets[0].Cells.ToDictionary(cell => cell.Source.Id.Address.ToString());
        Assert.Equal("B3+$D$4+$D3+B$4+SUM(B3:C3)+\"B3\"", cells["B3"].Formula!.Text);
        Assert.Equal("D3+$D$4+$D3+D$4+SUM(D3:E3)+\"B3\"", cells["D3"].Formula!.Text);
        Assert.Equal("D4+$D$4+$D4+D$4+SUM(D4:E4)+\"B3\"", cells["D4"].Formula!.Text);
        var references = cells["B3"].Formula!.References;
        Assert.Equal(new[] { (false, false), (true, true), (false, true), (true, false), (false, false) },
            references.Select(reference => (reference.First.RowAbsolute, reference.First.ColumnAbsolute)));
        Assert.Equal(new XlsxAddress(3, 2), references[4].First.Address);
        Assert.Equal(new XlsxAddress(3, 3), references[4].Last.Address);
        Assert.True(references[4].IsRange);
        Assert.All(references, reference => Assert.Equal(1u, reference.TargetSheetId));
        Assert.Equal("1+1", cells["C3"].Formula!.Text); // Ordinary override is independent of the group.
        Assert.Equal(4, cells.Count); // A master rectangle does not synthesize missing cells.
        Assert.Equal(xml, source.Sheets[0].Xml.ToString());
        Assert.Equal(bytes, input.ToArray());
    }

    [Fact]
    public void SharedCopy_Forward_two_axis_copy_reparses_growing_reference_spans_without_touching_literals()
    {
        const string master = "z9+$Z$9+$Z9+Z$9+SUM(Z9:AA10)+\"Z9\"+LOG10(1E10)";
        var bytes = SharedCopyWorkbook(master);
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var cells = plan.Sheets[0].Cells;
        Assert.Equal(master, cells[0].Formula!.Text); // Master text is authored text, including case.
        var follower = cells[1].Formula!;
        const string expected = "AB11+$Z$9+$Z11+AB$9+SUM(AB11:AC12)+\"Z9\"+LOG10(1E10)";
        Assert.Equal(expected, follower.Text);
        Assert.Equal(new[] { "AB11", "$Z$9", "$Z11", "AB$9", "AB11:AC12" },
            follower.References.Select(reference => follower.Text.Substring(reference.Start, reference.Length)));
        Assert.Equal(expected,
            services.GetRequiredService<XlsxFormulaReferences>().Rewrite(follower, reference =>
                reference.First + (reference.IsRange ? ":" + reference.Last : "")));
    }

    [Fact]
    public void SharedCopy_Master_operands_resolve_sheet_identity_and_names_without_copying_name_definitions()
    {
        var bytes = SharedInventoryWorkbook((workbook, part) =>
        {
            part.Worksheet!.Descendants<CellFormula>().Single(formula => formula.Reference is not null).Text =
                "Report!D4+'O''Brien Q4'!D4+Tax+Report!Tax+_xlfn.FUTURE(D4,\"D4\")+1E10";
            var other = workbook.AddNewPart<WorksheetPart>();
            other.Worksheet = new Worksheet(new SheetData(new Row(new Cell { CellReference = "B3", CellValue = new CellValue("7") }) { RowIndex = 3 }));
            workbook.Workbook!.Sheets!.Append(new Sheet { Name = "O'Brien Q4", SheetId = 41, Id = workbook.GetIdOfPart(other) });
            workbook.Workbook.DefinedNames!.Append(
                new DefinedName("'O''Brien Q4'!$B$3:$D$5") { Name = "_xlnm.Print_Area", LocalSheetId = 1 },
                new DefinedName("Report!$B$3") { Name = "Tax" },
                new DefinedName("Report!$C$3") { Name = "Tax", LocalSheetId = 0 });
        });
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var follower = plan.Sheets[0].Cells[0].Formula!;
        Assert.Equal("Report!B3+'O''Brien Q4'!B3+Tax+Report!Tax+_xlfn.FUTURE(B3,\"D4\")+1E10", follower.Text);
        Assert.Equal(new uint?[] { 1, 41, 1, 1, 1 }, follower.References.Select(reference => reference.TargetSheetId));
        Assert.Equal(new int?[] { null, null, 3, 3, null }, follower.References.Select(reference => reference.NameIndex));
        Assert.Equal(new XlsxAddress(3, 3), follower.References[2].First.Address);
        Assert.Equal(new[] { "Report!$B$3", "Report!$C$3" }, plan.DefinedNames.Select(name => name.Formula.Text));
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(3)]
    public void SharedCopy_Repeated_instances_reuse_normalized_source_formulas_without_output_displacement(int count)
    {
        var bytes = SharedInventoryWorkbook((_, part) =>
        {
            foreach (var row in part.Worksheet!.Descendants<Row>())
                row.PrependChild(_fixture.Text("A" + row.RowIndex!.Value, "$l(10).Items"));
        });
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var plan = planner.Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var sheet = plan.Sheets[0];
        using var data = JsonDocument.Parse("{\"Items\":[" + string.Join(',', Enumerable.Repeat("{\"$level\":10}", count)) + "]}");
        var bound = planner.BindSheet(sheet, data.RootElement, context);
        var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
        var positions = geometry.ExpansionGeometry(sheet, planner.ScheduleSheet(sheet, bound, context), geometry.HorizontalSpans(sheet, bound));
        var instances = geometry.EmitCells(sheet, positions).Rows.SelectMany(row => row.Cells)
            .Where(cell => cell.Source.Id.Address == new XlsxAddress(3, 2)).ToArray();
        Assert.Equal(count, instances.Length);
        Assert.Equal(Enumerable.Range(0, count).Select(index => new XlsxAddress(index * 2 + 1, 1)), instances.Select(cell => cell.Output));
        Assert.Equal(count, instances.Select(cell => cell.Id).Distinct().Count());
        Assert.Equal("B3+$D$4+$D3+B$4+SUM(B3:C3)+\"B3\"", sheet.Cells[0].Formula!.Text);
        AssertSharedInventoryFailure(bytes, "requires shared-formula relocation", "B3", errorCount: 3, json: data.RootElement.GetRawText());
    }

    [Theory]
    [InlineData("_xlfn.INDIRECT(\"D4\")", "INDIRECT/OFFSET")]
    [InlineData("SUM(D4", "parentheses")]
    [InlineData("UnknownName+D4", "unresolved")]
    [InlineData("Missing!D4", "unresolved worksheet")]
    [InlineData("[Book]Sheet!D4", "external workbook")]
    public void SharedCopy_Unsafe_master_operands_fail_at_the_master_before_any_result(string formula, string message)
    {
        var bytes = SharedInventoryWorkbook((_, part) =>
            part.Worksheet!.Descendants<CellFormula>().Single(element => element.Reference is not null).Text = formula);
        AssertSharedInventoryFailure(bytes, message, "D4");
    }

    [Theory]
    [InlineData("A1", "D4", "B3", "B3:D4")]
    [InlineData("$A1", "D4", "B3", "B3:D4")]
    [InlineData("A$1", "D4", "B3", "B3:D4")]
    [InlineData("XFD3", "B3", "C3", "B3:C3")]
    [InlineData("B1048576", "B3", "B4", "B3:B4")]
    public void SharedCopy_Unsafe_follower_copy_fails_at_the_follower_with_no_partial_result(string formula,
        string master, string follower, string range)
    {
        var bytes = _fixture.Create((_, part) =>
        {
            var data = part.Worksheet!.GetFirstChild<SheetData>()!; data.RemoveAllChildren();
            foreach (var row in new[] { SharedInventoryCell(master, formula, range, 7), SharedInventoryCell(follower, "ignored", null, 7) }
                .GroupBy(cell => uint.Parse(cell.CellReference!.Value![1..])).OrderBy(group => group.Key))
                data.Append(new Row(row.OrderBy(cell => cell.CellReference!.Value, StringComparer.Ordinal)) { RowIndex = row.Key });
        });
        AssertSharedInventoryFailure(bytes, "inside Excel worksheet dimensions", follower);
    }

    [Fact]
    public void SharedCopy_Mixed_range_endpoint_copy_with_increasing_span_normalizes()
    {
        var bytes = _fixture.Create((workbook, part) =>
        {
            var data = part.Worksheet!.GetFirstChild<SheetData>()!; data.RemoveAllChildren();
            data.Append(new Row(SharedInventoryCell("B3", "SUM($B3:D3)", "B3:F3", 7),
                SharedInventoryCell("F3", "", null, 7)) { RowIndex = 3 });
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$F$5";
        });
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>()
            .Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var follower = plan.Sheets[0].Cells.Single(cell => cell.Source.Id.Address == new XlsxAddress(3, 6));
        Assert.Equal("SUM($B3:H3)", follower.Formula!.Text);
    }

    [Fact]
    public void SharedCopy_Mixed_range_endpoint_crossing_absolute_endpoint_survives_reparse()
    {
        var bytes = _fixture.Create((workbook, part) =>
        {
            var data = part.Worksheet!.GetFirstChild<SheetData>()!; data.RemoveAllChildren();
            data.Append(new Row(SharedInventoryCell("B3", "SUM(A3:$C$3)", "B3:F3", 7),
                SharedInventoryCell("F3", "", null, 7)) { RowIndex = 3 });
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$F$5";
        });
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>()
            .Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var follower = plan.Sheets[0].Cells.Single(cell => cell.Source.Id.Address == new XlsxAddress(3, 6));
        Assert.NotNull(follower.Formula);
        Assert.Equal("SUM(E3:$C$3)", follower.Formula.Text);
        AssertSharedInventoryFailure(bytes, "requires shared-formula relocation", "B3", errorCount: 2);
    }

    private byte[] SharedCopyWorkbook(string formula) => _fixture.Create((_, part) =>
    {
        var data = part.Worksheet!.GetFirstChild<SheetData>()!; data.RemoveAllChildren();
        data.Append(new Row(SharedInventoryCell("B3", formula, "B3:D5", 0)) { RowIndex = 3 },
            new Row(SharedInventoryCell("D5", "ignored follower text", null, 0)) { RowIndex = 5 });
    });
}
