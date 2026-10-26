using System.Collections.Immutable;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed class Feature012StaticTests
{
    private readonly XlsxFixture _fixture = new();

    [Theory]
    [InlineData("12.5", "n", "12.5")]
    [InlineData("true", "b", "1")]
    [InlineData("false", "b", "0")]
    [InlineData("\"00123\"", "inlineStr", "00123")]
    [InlineData("\"=SUM(A1)\"", "inlineStr", "=SUM(A1)")]
    [InlineData("\"{Other}\"", "inlineStr", "{Other}")]
    [InlineData("null", null, "")]
    public void Static_Private_render_retains_native_type_style_and_inert_insertions(string json, string? type, string expected)
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            AddStyles(workbook);
            var cell = worksheet.Worksheet!.Descendants<Cell>().Single();
            cell.StyleIndex = 1;
            cell.InlineString = new InlineString(new Run(new RunProperties(new Bold()), new Text("{Name}")));
        });
        var (output, context) = RenderPrivate(source, "{\"Name\":" + json + "}");
        using var result = SpreadsheetDocument.Open(new MemoryStream(output), false);
        var cell = result.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single();
        Assert.Equal(type, cell.DataType?.InnerText);
        Assert.Equal(expected, type == "inlineStr" ? cell.InlineString!.InnerText : cell.CellValue?.Text ?? string.Empty);
        Assert.Equal(1U, cell.StyleIndex!.Value);
        Assert.Null(cell.CellFormula);
        if (type == "inlineStr") Assert.Null(cell.CellValue); else Assert.Null(cell.InlineString);
        Assert.Equal(2, result.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Count());
        Assert.Equal(1, context.OutputCells);
        AssertValid(result);
    }

    [Theory]
    [InlineData("0.1")]
    [InlineData("123456789012345")]
    [InlineData("1.2300000000000000000e5")]
    [InlineData("10000000000000000")]
    [InlineData("1e308")]
    [InlineData("1e-307")]
    [InlineData("0e10")]
    [InlineData("-0e999999999999999999999")]
    public void Static_Safe_numeric_digits_exponents_and_zero_are_accepted(string number)
    {
        var (output, _) = RenderPrivate(_fixture.Create((_, part) => part.Worksheet!.Descendants<Cell>().Single().InlineString = new InlineString(new Run(new Text("{Name}")))), "{\"Name\":" + number + "}");
        using var result = SpreadsheetDocument.Open(new MemoryStream(output), false);
        var cell = result.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single();
        Assert.Equal(CellValues.Number, cell.DataType!.Value);
        Assert.Equal(number.StartsWith("-0", StringComparison.Ordinal) || number == "0e10" ? "0" : number, cell.CellValue!.Text);
    }

    [Theory]
    [InlineData("10000000000000001")]
    [InlineData("1.00000000000000001")]
    [InlineData("10000000000000000555")]
    [InlineData("1234567890123456789")]
    [InlineData("1e309")]
    [InlineData("1e-400")]
    [InlineData("1e-320")]
    public void Static_Unsafe_numeric_values_fail_with_safe_actionable_diagnostics(string number)
    {
        var error = Assert.Throws<TemplateException>(() => RenderPrivate(_fixture.Create((_, part) => part.Worksheet!.Descendants<Cell>().Single().InlineString = new InlineString(new Run(new Text("{Name}")))), "{\"Name\":" + number + "}"));
        var diagnostic = Assert.Single(error.Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeNumber, diagnostic.Code);
        Assert.Equal("{Name}", diagnostic.Field);
        Assert.Equal("$.Name", diagnostic.DataPath);
        Assert.Contains("JSON string", diagnostic.Message);
        Assert.DoesNotContain(number, JsonSerializer.Serialize(error.Errors));
    }

    [Fact]
    public void Static_Shared_source_strings_are_prepared_per_cell_with_distinct_scopes()
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            workbook.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable(new SharedStringItem(new Text("{Name}")));
            var row = worksheet.Worksheet!.Descendants<Row>().Single();
            row.RemoveAllChildren();
            row.Append(new Cell { CellReference = "B3", DataType = CellValues.SharedString, CellValue = new CellValue("0") },
                new Cell { CellReference = "C3", DataType = CellValues.SharedString, CellValue = new CellValue("0") });
        });
        using var buffer = new MemoryStream(); buffer.Write(source); buffer.Position = 0;
        using var services = _fixture.Services();
        using (var package = SpreadsheetDocument.Open(buffer, true))
        {
            var context = new XlsxOperationContext(new MergeOptions());
            var planner = services.GetRequiredService<XlsxTemplatePlanner>();
            var plan = planner.Plan(new XlsxPackageReader().Read(package, context), context);
            using var data = JsonDocument.Parse("""{"Items":[{"Name":"Alpha"},{"Name":"Beta"}]}""");
            var cells = plan.Sheets[0].Cells;
            var values = cells.Select((cell, index) => KeyValuePair.Create(cell.Source.Id,
                planner.PrepareValue(cell, new DataScope(data.RootElement, [data.RootElement.GetProperty("Items")[index]], ["$.Items[" + index + "]"]), context)))
                .ToImmutableDictionary();
            var outputPlan = services.GetRequiredService<XlsxGeometryPlanner>().Plan(plan, values, context);
            services.GetRequiredService<XlsxDocumentRenderer>().Render(package, outputPlan, context);
        }
        using var result = SpreadsheetDocument.Open(new MemoryStream(buffer.ToArray()), false);
        Assert.Equal(new[] { "Alpha", "Beta" }, result.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Select(cell => cell.InlineString!.InnerText));
        Assert.Null(result.WorkbookPart.SharedStringTablePart);
        AssertValid(result);
    }

    [Theory]
    [InlineData("[]")]
    [InlineData("null")]
    [InlineData("1")]
    public void Static_Nonobject_data_and_invalid_xml_text_reject_without_partial_result(string json)
    {
        var error = Assert.Throws<TemplateException>(() => RenderPrivate(_fixture.Create(), json));
        Assert.Contains("JSON object", error.Errors[0].Message);
        var textError = Assert.Throws<TemplateException>(() => RenderPrivate(_fixture.Create(), "{\"Name\":\"secret\\u0001\"}"));
        Assert.Equal(DiagnosticCodes.GenerationInvalidData, textError.Code);
        Assert.Equal("$.Name", textError.Errors[0].DataPath);
        Assert.DoesNotContain("secret", JsonSerializer.Serialize(textError.Errors));
        Assert.Contains("XML control characters", textError.Errors[0].Message);
    }

    private (byte[] Output, XlsxOperationContext Context) RenderPrivate(byte[] source, string json)
    {
        using var services = _fixture.Services();
        using var data = JsonDocument.Parse(json);
        using var buffer = new MemoryStream(); buffer.Write(source); buffer.Position = 0;
        var context = new XlsxOperationContext(new MergeOptions());
        using (var package = SpreadsheetDocument.Open(buffer, true))
        {
            var planner = services.GetRequiredService<XlsxTemplatePlanner>();
            var plan = planner.Plan(new XlsxPackageReader().Read(package, context), context);
            var values = planner.Bind(plan, data.RootElement, context);
            var outputPlan = services.GetRequiredService<XlsxGeometryPlanner>().Plan(plan, values, context);
            services.GetRequiredService<XlsxDocumentRenderer>().Render(package, outputPlan, context);
        }
        return (buffer.ToArray(), context);
    }

    private static void AddStyles(WorkbookPart workbook)
    {
        workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()),
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
            new Borders(new Border()), new CellStyleFormats(new CellFormat()),
            new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 4, ApplyNumberFormat = true }));
    }

    private static void AssertValid(SpreadsheetDocument document) => Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2010).Validate(document));
}
