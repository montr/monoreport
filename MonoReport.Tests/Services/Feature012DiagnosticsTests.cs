using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed class Feature012DiagnosticsTests
{
    private readonly XlsxFixture _fixture = new();

    [Theory]
    [InlineData(99)]
    [InlineData(100)]
    [InlineData(101)]
    public void Static_Exactly_the_100th_error_stops_resolution_without_a_101st_probe(int count)
    {
        var source = _fixture.Create((_, worksheet) => worksheet.Worksheet!.Descendants<Text>().Single().Text =
            string.Join(' ', Enumerable.Repeat("{Missing}", count)));
        var options = new MergeOptions { MissingValues = MissingValueBehavior.Error };
        using var services = _fixture.Services();
        using var document = SpreadsheetDocument.Open(new MemoryStream(source), false);
        using var data = JsonDocument.Parse("{}");
        var context = new XlsxOperationContext(options);
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var plan = planner.Plan(new XlsxPackageReader().Read(document, context), context);
        var error = Assert.Throws<TemplateException>(() => planner.Bind(plan, data.RootElement, context));
        Assert.Equal(Math.Min(count, 100), context.ResolvedFields);
        Assert.Equal(Math.Min(count, 100), error.Errors.Count);
        Assert.Equal(count >= 100, error.LimitReached);
        Assert.All(error.Errors, item => { Assert.Equal("$.Missing", item.DataPath); Assert.Contains("B3", item.Location); });
        MergeResult? result = null;
        var published = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, "{}", options));
        Assert.Null(result); Assert.Equal(error.Errors, published.Errors); Assert.Equal(error.LimitReached, published.LimitReached);
    }

    [Fact]
    public void Static_Errors_and_warnings_keep_workbook_then_source_cell_order_and_safe_data_paths()
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            var row = worksheet.Worksheet!.Descendants<Row>().Single();
            row.RemoveAllChildren(); row.Append(_fixture.Text("B3", "{First}"), _fixture.Text("C3", "{Second}"));
            var other = workbook.AddNewPart<WorksheetPart>();
            other.Worksheet = new Worksheet(new SheetData(new Row(_fixture.Text("A1", "{Third}")) { RowIndex = 1 }));
            workbook.Workbook!.Sheets!.Append(new Sheet { Name = "A second", SheetId = 2, Id = workbook.GetIdOfPart(other) });
            workbook.Workbook.DefinedNames!.Append(new DefinedName("'A second'!$A$1") { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
        });
        var strict = Assert.Throws<TemplateException>(() => _fixture.Merge(source, "{}", new MergeOptions { MissingValues = MissingValueBehavior.Error }));
        Assert.Equal(new[] { "$.First", "$.Second", "$.Third" }, strict.Errors.Select(item => item.DataPath));
        Assert.Contains("B3", strict.Errors[0].Location); Assert.Contains("C3", strict.Errors[1].Location); Assert.Contains("A second", strict.Errors[2].Location);
        var result = _fixture.Merge(source, "{}", new MergeOptions { MissingValues = MissingValueBehavior.Warning });
        Assert.Equal(new[] { "$.First", "$.Second", "$.Third" }, result.Warnings.Select(item => item.DataPath));
        Assert.All(result.Warnings, item => Assert.Equal(DiagnosticCodes.XlsxMissingValue, item.Code));
    }

    [Fact]
    public async Task Static_Reused_engine_isolates_mixed_word_spreadsheet_successes_failures_and_owned_bytes()
    {
        using var services = _fixture.Services();
        var engine = services.GetRequiredService<MergeEngine>();
        var xlsx = _fixture.Create(); var word = Docx.Create(Docx.P(Docx.Run("{Name}")));
        var outputs = await Task.WhenAll(Enumerable.Range(0, 8).Select(index => Task.Run(() =>
        {
            using var invalid = new MemoryStream(xlsx);
            using var invalidData = JsonDocument.Parse("[]");
            Assert.Throws<TemplateException>(() => engine.Merge(invalid, invalidData));
            var spreadsheet = index % 2 == 0;
            using var input = new MemoryStream(spreadsheet ? xlsx : word);
            using var data = JsonDocument.Parse("{\"Name\":\"" + index + "\"}");
            var result = engine.Merge(input, data);
            Assert.Empty(result.Warnings);
            if (!spreadsheet) Assert.Equal(index.ToString(), Docx.Text(Docx.Read(result.Document)));
            else
            {
                using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
                Assert.Equal(index.ToString(), package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single().InlineString!.InnerText);
            }
            return result.Document;
        })));
        var independent = outputs[2].ToArray(); outputs[0][0] = 0;
        Assert.Equal(independent, outputs[2]);
        using var array = JsonDocument.Parse("[]");
        using var staticWord = new MemoryStream(Docx.Create(Docx.P(Docx.Run("unchanged"))));
        Assert.Equal("unchanged", Docx.Text(Docx.Read(engine.Merge(staticWord, array).Document)));
    }

    [Fact]
    public void Static_Preflight_combines_requirement_conflicts_and_parse_errors_in_source_order()
    {
        var source = _fixture.Create((_, worksheet) =>
        {
            var row = worksheet.Worksheet!.Descendants<Row>().Single(); row.RemoveAllChildren();
            row.Append(_fixture.Text("B3", "{Thing}"), _fixture.Text("C3", "{Thing.Name}"), _fixture.Text("D3", "broken {"));
        });
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.Equal(new[] { DiagnosticCodes.TemplateRequirementConflict, DiagnosticCodes.TemplateRequirementConflict, DiagnosticCodes.TemplateUnmatchedBrace },
            analysis.Errors.Select(issue => issue.Code));
        Assert.Equal(new[] { "B3", "C3", "D3" }, analysis.Errors.Select(issue => issue.Location!.Split(", ")[^1]));
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source));
        Assert.Equal(analysis.Errors.Select(issue => issue.Code), error.Errors.Select(issue => issue.Code));
        Assert.Equal(error.Errors[0].Code, error.Code);
    }
}
