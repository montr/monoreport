using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Cli.Services;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature012CliTests
{
    private readonly XlsxFixture _fixture = new();

    [Fact]
    public void Analyzer_Nested_scope_fixture_matches_core_and_keeps_generation_guard()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures/feature012/analyzer-scopes.template.xlsx");
        var source = File.ReadAllBytes(path);
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analyzer = services.GetRequiredService<TemplateAnalyzer>();
        var expected = analyzer.Analyze(input); Assert.True(expected.Valid);
        var app = new CliApplication(new CliOptionsParser(), services.GetRequiredService<MergeEngine>(), analyzer);
        using var output = new MemoryStream(); using var errors = new StringWriter();
        Assert.Equal(0, app.Run(["--analyze", "--template", path], Stream.Null, output, errors));
        Assert.Empty(errors.ToString());
        using var json = JsonDocument.Parse(output.ToArray());
        Assert.True(JsonElement.DeepEquals(JsonSerializer.SerializeToElement(expected,
            new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase }), json.RootElement));
        Assert.Equal(new[] { "Bids", "Criteria", "Items", "Title" }, json.RootElement.GetProperty("requirements").EnumerateArray().Select(item => item.GetProperty("name").GetString()));
        using var rejected = new MemoryStream();
        Assert.Equal(1, app.Run(["--template", path, "--json", "{\"Items\":[]}"], Stream.Null, rejected, errors));
        Assert.Empty(rejected.ToArray()); Assert.Contains(DiagnosticCodes.XlsxUnsupportedFeature, errors.ToString());
        Assert.Equal(source, File.ReadAllBytes(path));
    }

    [Fact]
    public void Repeated_Existing_cli_switches_render_scoped_native_rows_and_transformed_formulas()
    {
        var path = Path.Combine(Path.GetTempPath(), "feature012-m2-" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.WriteAllBytes(path, _fixture.CreateRepeated());
        try
        {
            using var services = _fixture.Services();
            var app = new CliApplication(new CliOptionsParser(), services.GetRequiredService<MergeEngine>(), services.GetRequiredService<TemplateAnalyzer>());
            using var output = new MemoryStream(); using var errors = new StringWriter();
            Assert.Equal(0, app.Run(["--template", path, "--json", "{\"Items\":[{\"$level\":10,\"Name\":\"one\",\"Price\":12},{\"$level\":10,\"Name\":\"two\",\"Price\":20}]}"], Stream.Null, output, errors));
            Assert.Empty(errors.ToString());
            using var document = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            Assert.Empty(new OpenXmlValidator().Validate(document));
            var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("one", cells["A1"].InlineString!.InnerText); Assert.Equal("20", cells["B2"].CellValue!.Text);
            Assert.Equal("B1*2", cells["C1"].CellFormula!.Text); Assert.Equal("SUM(B1:B2)", cells["C4"].CellFormula!.Text);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Static_Generate_and_analyze_use_existing_switches_even_with_a_misleading_filename()
    {
        var path = Path.Combine(Path.GetTempPath(), "feature012-" + Guid.NewGuid().ToString("N") + ".docx");
        File.WriteAllBytes(path, _fixture.Create());
        try
        {
            using var services = _fixture.Services();
            var app = new CliApplication(new CliOptionsParser(), services.GetRequiredService<MergeEngine>(), services.GetRequiredService<TemplateAnalyzer>());
            using var output = new MemoryStream(); using var errors = new StringWriter();
            Assert.Equal(0, app.Run(["--template", path, "--json", "{\"Name\":\"Ada\"}"], Stream.Null, output, errors));
            Assert.Empty(errors.ToString());
            using var package = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            Assert.Empty(new OpenXmlValidator().Validate(package));
            Assert.Equal("Ada", package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single().InlineString!.InnerText);
            using var analysis = new MemoryStream();
            Assert.Equal(0, app.Run(["--analyze", "--template", path], Stream.Null, analysis, errors));
            using var json = JsonDocument.Parse(analysis.ToArray());
            Assert.True(json.RootElement.GetProperty("valid").GetBoolean());
            Assert.Equal("$.Name", json.RootElement.GetProperty("occurrences")[0].GetProperty("path").GetString());
            Assert.Equal(new[] { "errors", "occurrences", "requirements", "valid", "warnings" }, json.RootElement.EnumerateObject().Select(item => item.Name).Order());
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Static_Failures_leave_stdout_empty_and_existing_output_untouched(bool malformed)
    {
        var path = Path.Combine(Path.GetTempPath(), "feature012-" + Guid.NewGuid().ToString("N") + ".xlsx");
        var destination = path + ".result";
        var source = _fixture.Create();
        if (malformed) source = _fixture.RewritePart(source, "xl/worksheets/sheet1.xml", xml => xml[..^10]);
        File.WriteAllBytes(path, source); File.WriteAllText(destination, "existing output sentinel");
        try
        {
            using var services = _fixture.Services();
            var app = new CliApplication(new CliOptionsParser(), services.GetRequiredService<MergeEngine>(), services.GetRequiredService<TemplateAnalyzer>());
            using var output = new MemoryStream(); using var errors = new StringWriter();
            Assert.Equal(1, app.Run(["--template", path, "--json", "{\"Name\":10000000000000001}", "--output", destination], Stream.Null, output, errors));
            Assert.Empty(output.ToArray()); Assert.Equal("existing output sentinel", File.ReadAllText(destination));
            Assert.Contains(malformed ? DiagnosticCodes.TemplateInvalidXml : DiagnosticCodes.XlsxUnsafeNumber, errors.ToString());
            Assert.DoesNotContain("10000000000000001", errors.ToString());
            if (!malformed) Assert.Contains("JSON string", errors.ToString());
        }
        finally { File.Delete(path); File.Delete(destination); }
    }

    [Fact]
    public void Static_Revision_preserves_string_semantics_and_reports_surrogates_without_output()
    {
        var path = Path.Combine(Path.GetTempPath(), "feature012-revision-" + Guid.NewGuid().ToString("N") + ".xlsx");
        var destination = path + ".result";
        File.WriteAllBytes(path, _fixture.Create()); File.WriteAllText(destination, "existing output sentinel");
        try
        {
            using var services = _fixture.Services();
            var app = new CliApplication(new CliOptionsParser(), services.GetRequiredService<MergeEngine>(), services.GetRequiredService<TemplateAnalyzer>());
            using var output = new MemoryStream(); using var errors = new StringWriter();
            Assert.Equal(0, app.Run(["--template", path, "--json", JsonSerializer.Serialize(new { Name = "a\r_x000D_" })], Stream.Null, output, errors));
            Assert.Equal("a\r_x000D_", _fixture.ReadCellText(output.ToArray(), "A1")); Assert.Empty(errors.ToString());
            using var failureOutput = new MemoryStream();
            Assert.Equal(1, app.Run(["--template", path, "--json", "{\"Name\":\"private\\uD800\"}", "--output", destination], Stream.Null, failureOutput, errors));
            Assert.Empty(failureOutput.ToArray()); Assert.Equal("existing output sentinel", File.ReadAllText(destination));
            Assert.Contains(DiagnosticCodes.GenerationInvalidData, errors.ToString()); Assert.Contains("$.Name", errors.ToString());
            Assert.Contains("paired surrogate", errors.ToString()); Assert.DoesNotContain("private", errors.ToString());
        }
        finally { File.Delete(path); File.Delete(destination); }
    }
}
