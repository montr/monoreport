using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services;
using MonoReport.Services.Xlsx;
using System.Text.Json;

namespace MonoReport.Tests.Services;

public sealed class Feature012FoundationTests
{
    private readonly XlsxFixture _fixture = new();

    [Theory]
    [InlineData(null)]
    [InlineData("'Report'!$B$3:$D$5,'Report'!$F$3:$G$4")]
    [InlineData("'Report'!$B:$D")]
    [InlineData("'Other'!$B$3:$D$5")]
    [InlineData("'Report'!$D$5:$B$3")]
    [InlineData("'Report'!$B$3:$XFE$5")]
    public void Inventory_Requires_finite_single_rectangular_print_area(string? area)
    {
        var source = _fixture.Create((workbook, _) =>
        {
            var name = workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single();
            if (area is null) name.Remove(); else name.Text = area;
        });
        using var provider = _fixture.Services();
        using var input = new MemoryStream(source);
        var result = provider.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(result.Valid);
        var error = Assert.Single(result.Errors);
        Assert.Equal(DiagnosticCodes.XlsxInvalidPrintArea, error.Code);
        Assert.Contains("PrintArea", error.Message);
        Assert.Contains("Report", error.Location);
    }

    [Fact]
    public void Inventory_Preserves_sheet_order_native_cells_styles_and_feature_inventory()
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            workbook.Workbook!.InsertAt(new WorkbookProperties { Date1904 = true }, 0);
            var row = worksheet.Worksheet!.GetFirstChild<SheetData>()!.GetFirstChild<Row>()!;
            row.Append(new Cell { CellReference = "C3", StyleIndex = 4, CellValue = new CellValue("123.5") });
            row.Append(new Cell { CellReference = "D3", DataType = CellValues.Boolean, CellValue = new CellValue("1") });
            worksheet.Worksheet.Append(new MergeCells(new MergeCell { Reference = "B4:C4" }));
            var second = workbook.AddNewPart<WorksheetPart>();
            second.Worksheet = new Worksheet(new SheetData());
            workbook.Workbook.Sheets!.Append(new Sheet { Name = "A second", SheetId = 2, Id = workbook.GetIdOfPart(second) });
            workbook.Workbook.DefinedNames!.Append(new DefinedName("'A second'!$A$1:$A$1") { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
        });
        using var input = new MemoryStream(source);
        using var document = SpreadsheetDocument.Open(input, false);
        var context = new XlsxOperationContext(new MergeOptions());
        var inventory = new XlsxPackageReader().Read(document, context);
        Assert.Equal(new[] { "Report", "A second" }, inventory.Sheets.Select(sheet => sheet.Name));
        Assert.True(inventory.Date1904);
        Assert.Equal(3, context.SourceCells);
        var sheet = inventory.Sheets[0];
        Assert.Equal("B3:D5", sheet.PrintArea.ToString());
        Assert.Contains("mergeCells", sheet.Features);
        Assert.Equal("{Name}", sheet.Cells[0].Text);
        Assert.Null(sheet.Cells[1].Text);
        Assert.Equal("123.5", sheet.Cells[1].Xml.Elements().Single().Value);
        Assert.Equal("4", (string?)sheet.Cells[1].Xml.Attribute("s"));
        Assert.Equal("b", sheet.Cells[2].NativeType);
        Assert.Equal(new SourceCellId(1, new XlsxAddress(3, 4)), sheet.Cells[2].Id);
        Assert.Equal(source, input.ToArray());
    }

    [Theory]
    [InlineData(1048576L, 16384L, 1L, 1L, true)]
    [InlineData(1L, 1L, 0L, 1L, true)]
    [InlineData(1048576L, 1L, 2L, 1L, false)]
    [InlineData(1L, 16384L, 1L, 2L, false)]
    [InlineData(long.MaxValue, 1L, 2L, 1L, false)]
    public void Inventory_Checks_physical_dimensions_before_allocation(long rows, long columns, long rowCopies, long columnCopies, bool safe)
    {
        var planner = new XlsxGeometryPlanner();
        if (safe) Assert.Equal((checked((int)(rows * rowCopies)), checked((int)(columns * columnCopies))),
            planner.Dimensions(rows, columns, rowCopies, columnCopies));
        else Assert.Equal(DiagnosticCodes.XlsxDimensionLimit,
            Assert.Throws<TemplateException>(() => planner.Dimensions(rows, columns, rowCopies, columnCopies)).Code);
    }

    [Fact]
    public void Scalar_Neutral_scope_preserves_missing_null_case_and_parent_root_paths()
    {
        using var data = JsonDocument.Parse("""{"Title":"root","Items":[{"Title":"parent","Child":{"Name":"leaf","Nil":null}}]}""");
        var root = data.RootElement;
        var parent = root.GetProperty("Items")[0];
        var scope = new DataScope(root, [parent.GetProperty("Child"), parent, root], ["$.Items[0].Child", "$.Items[0]", "$"]);
        var resolver = new DataResolver();
        var parser = new TemplateLanguageParser();
        var paths = new[] { "Name", "name", "Nil", "../Title", "$.Title" }
            .Select(field => parser.Parse("{" + field + "}", "test").Tokens.Single().Expression!.Path!).ToArray();
        Assert.Equal("leaf", resolver.Resolve(paths[0], scope).Value.GetString());
        Assert.Equal(ResolutionState.Missing, resolver.Resolve(paths[1], scope).State);
        Assert.Equal(ResolutionState.Null, resolver.Resolve(paths[2], scope).State);
        Assert.Equal("parent", resolver.Resolve(paths[3], scope).Value.GetString());
        Assert.Equal("root", resolver.Resolve(paths[4], scope).Value.GetString());
        Assert.Equal(new[] { "$.Items[0].Child.Name", "$.Items[0].Child.name", "$.Items[0].Child.Nil", "$.Items[0].Title", "$.Title" },
            paths.Select(path => resolver.DataPath(path, scope)));
    }

    [Theory]
    [InlineData("{Number}", "n", "12.5")]
    [InlineData("{Flag}", "b", "1")]
    [InlineData("{Text}", "inlineStr", "{Name}=1")]
    [InlineData("{Nil}", null, null)]
    [InlineData("{Missing}", null, null)]
    [InlineData("{{{Text}}}", "inlineStr", "{{Name}=1}")]
    [InlineData("Total: {Number}", "inlineStr", "Total: 12.5")]
    [InlineData("{Number number:\"0.00\"}", "inlineStr", "12.50")]
    [InlineData("{Text upper}", "inlineStr", "{NAME}=1")]
    public void Scalar_Parse_once_prepares_native_or_text_without_reinterpreting_insertions(string template, string? type, string? value)
    {
        var (plan, planner) = ScalarPlan(template);
        using var data = JsonDocument.Parse("""{"Number":12.5,"Flag":true,"Text":"{Name}=1","Name":"must not replace","Nil":null}""");
        var context = new XlsxOperationContext(new MergeOptions());
        var result = planner.PrepareValue(plan, new DataScope(data.RootElement, [data.RootElement], ["$"]), context);
        Assert.Equal(new XlsxPreparedValue(type, value), result);
        Assert.Empty(context.Errors);
    }

    [Fact]
    public void Scalar_Missing_policy_differs_from_null_and_objects_are_errors()
    {
        using var data = JsonDocument.Parse("""{"Nil":null,"Object":{},"Array":[]}""");
        var scope = new DataScope(data.RootElement, [data.RootElement], ["$"]);
        foreach (var field in new[] { "Missing", "Nil", "Object", "Array" })
        {
            var (plan, planner) = ScalarPlan("{" + field + "}");
            var context = new XlsxOperationContext(new MergeOptions { MissingValues = MissingValueBehavior.Warning });
            planner.PrepareValue(plan, scope, context);
            Assert.Equal(field == "Missing" ? 1 : 0, context.Warnings.Count);
            Assert.Equal(field is "Object" or "Array" ? 1 : 0, context.Errors.Count);
            if (context.Errors.Count > 0)
            {
                Assert.Equal("$." + field, context.Errors[0].DataPath);
                Assert.Contains("B3", context.Errors[0].Location);
            }
        }
        var (missing, service) = ScalarPlan("{Missing}");
        var strict = new XlsxOperationContext(new MergeOptions { MissingValues = MissingValueBehavior.Error });
        service.PrepareValue(missing, scope, strict);
        Assert.Throws<TemplateException>(strict.ThrowIfErrors);
    }

    [Theory]
    [InlineData("{../Name}", DiagnosticCodes.TemplateInvalidContext)]
    [InlineData("{If:Flag}", DiagnosticCodes.XlsxUnsupportedFeature)]
    [InlineData("{Raw}", DiagnosticCodes.XlsxUnsupportedFeature)]
    [InlineData("{Name number:\"Q\"}", DiagnosticCodes.TemplateInvalidExpression)]
    [InlineData("{Name", DiagnosticCodes.TemplateUnmatchedBrace)]
    public void Scalar_Invalid_or_structural_expressions_fail_before_binding(string text, string code)
    {
        var source = _fixture.Create((_, worksheet) => worksheet.Worksheet!.Descendants<Text>().Single().Text = text);
        using var input = new MemoryStream(source);
        using var services = _fixture.Services();
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(analysis.Valid);
        Assert.Contains(analysis.Errors, error => error.Code == code && error.Location!.Contains("B3", StringComparison.Ordinal));
        Assert.Equal(source, input.ToArray());
    }

    private (XlsxCellPlan Cell, XlsxTemplatePlanner Planner) ScalarPlan(string text)
    {
        var source = _fixture.Create((_, worksheet) => worksheet.Worksheet!.Descendants<Text>().Single().Text = text);
        using var document = SpreadsheetDocument.Open(new MemoryStream(source), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var planner = new XlsxTemplatePlanner(new TemplateLanguageParser(), new DataResolver(), new ValueFormatter(new RubSpellOutFormatter()), new XlsxFormulaReferences());
        var plan = planner.Plan(new XlsxPackageReader().Read(document, context), context);
        return (plan.Sheets.Single().Cells.Single(), planner);
    }
}
