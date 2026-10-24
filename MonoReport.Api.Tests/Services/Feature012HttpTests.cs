using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature012HttpTests
{
    private readonly XlsxFixture _fixture = new();

    [Fact]
    public async Task Analyzer_Nested_scope_fixture_matches_core_and_allows_empty_generation()
    {
        var source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures/feature012/analyzer-scopes.template.xlsx"));
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var expected = services.GetRequiredService<TemplateAnalyzer>().Analyze(input); Assert.True(expected.Valid);
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var form = Form(source, null, true);
        using var response = await client.PostAsync("/api/documents/analyze", form); response.EnsureSuccessStatusCode();
        using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        Assert.True(JsonElement.DeepEquals(JsonSerializer.SerializeToElement(expected,
            new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase }), json.RootElement));
        Assert.Equal(new[] { "Bids", "Criteria", "Items", "Title" }, json.RootElement.GetProperty("requirements").EnumerateArray().Select(item => item.GetProperty("name").GetString()));
        using var generation = Form(source, "{\"Items\":[]}", true);
        using var generationResponse = await client.PostAsync("/api/documents/generate", generation);
        generationResponse.EnsureSuccessStatusCode();
        using var generated = SpreadsheetDocument.Open(new MemoryStream(await generationResponse.Content.ReadAsByteArrayAsync()), false);
        Assert.Empty(new OpenXmlValidator().Validate(generated));
        Assert.Equal(source, input.ToArray());
    }

    [Fact]
    public async Task Nested_Public_generation_returns_xlsx_and_rejects_bad_matrix_without_report_headers()
    {
        var fixture = new XlsxNestedFixture(); var source = fixture.Create(shared: true); var original = source.ToArray();
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var form = Form(source, fixture.Data(3), true);
        using var response = await client.PostAsync("/api/documents/generate", form);
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", response.Content.Headers.ContentType!.MediaType);
        Assert.Contains("document.xlsx", response.Content.Headers.ContentDisposition!.ToString());
        Assert.False(response.Headers.Contains("X-MonoReport-Warning-Count"));
        fixture.AssertGenerated(await response.Content.ReadAsByteArrayAsync());
        using var invalid = Form(source, "{\"Bids\":[{}],\"Criteria\":[{},{}],\"Items\":[{\"$level\":10,\"$path\":\"p\"}]}", true);
        using var rejected = await client.PostAsync("/api/documents/generate", invalid);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, rejected.StatusCode);
        Assert.Equal("application/problem+json", rejected.Content.Headers.ContentType!.MediaType);
        Assert.Null(rejected.Content.Headers.ContentDisposition); Assert.False(rejected.Headers.Contains("X-MonoReport-Warning-Count"));
        using var error = JsonDocument.Parse(await rejected.Content.ReadAsStringAsync());
        Assert.Equal(DiagnosticCodes.GenerationInvalidData, error.RootElement.GetProperty("code").GetString());
        var issue = error.RootElement.GetProperty("errors")[0];
        Assert.Equal("$.Items[0]", issue.GetProperty("dataPath").GetString()); Assert.Contains("A9", issue.GetProperty("location").GetString());
        Assert.Equal(original, source);
    }

    [Fact]
    public async Task Activated_deep_branching_multirow_layout_generates_through_http_contract()
    {
        var fixture = new XlsxNestedFixture(); var source = fixture.CreateActivated();
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var form = Form(source, XlsxNestedFixture.ActivatedData(), true);
        using var response = await client.PostAsync("/api/documents/generate", form);
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", response.Content.Headers.ContentType!.MediaType);
        Assert.Contains("document.xlsx", response.Content.Headers.ContentDisposition!.ToString());
        Assert.False(response.Headers.Contains("X-MonoReport-Warning-Count"));
        XlsxNestedFixture.AssertActivated(await response.Content.ReadAsByteArrayAsync());

        using var invalid = Form(source, "{\"Bids\":[{}],\"Criteria\":[{},{}],\"Third\":[{}],\"Items\":[{\"$level\":10,\"$path\":\"p\"}]}", true);
        using var rejected = await client.PostAsync("/api/documents/generate", invalid);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, rejected.StatusCode);
        Assert.Equal("application/problem+json", rejected.Content.Headers.ContentType!.MediaType);
        Assert.Null(rejected.Content.Headers.ContentDisposition);
        using var problem = JsonDocument.Parse(await rejected.Content.ReadAsStringAsync());
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, problem.RootElement.GetProperty("code").GetString());
        Assert.Contains("formula", problem.RootElement.GetProperty("detail").GetString(), StringComparison.OrdinalIgnoreCase);
        Assert.NotEmpty(problem.RootElement.GetProperty("errors")[0].GetProperty("location").GetString()!);
    }

    [Fact]
    public async Task Repeated_Existing_routes_generate_native_rows_and_formulas_and_reject_bad_cardinality_before_output()
    {
        var source = _fixture.CreateRepeated();
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var form = Form(source, "{\"Items\":[{\"$level\":10,\"Name\":\"one\",\"Price\":12},{\"$level\":10,\"Name\":\"two\",\"Price\":20}]}", true);
        using var response = await client.PostAsync("/api/documents/generate", form); response.EnsureSuccessStatusCode();
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", response.Content.Headers.ContentType!.MediaType);
        using var document = SpreadsheetDocument.Open(new MemoryStream(await response.Content.ReadAsByteArrayAsync()), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("one", cells["A1"].InlineString!.InnerText); Assert.Equal("two", cells["A2"].InlineString!.InnerText);
        Assert.Equal("12", cells["B1"].CellValue!.Text); Assert.Equal("B2*2", cells["C2"].CellFormula!.Text);
        Assert.Equal("SUM(B1:B2)", cells["C4"].CellFormula!.Text);
        using var analysisForm = Form(source, null, true); using var analysis = await client.PostAsync("/api/documents/analyze", analysisForm);
        using var json = JsonDocument.Parse(await analysis.Content.ReadAsStringAsync());
        Assert.True(json.RootElement.GetProperty("valid").GetBoolean()); Assert.Equal("Items", json.RootElement.GetProperty("requirements")[0].GetProperty("name").GetString());
        using var invalidForm = Form(source, "{\"Items\":[{\"$level\":\"bad\"}]}", true);
        using var invalid = await client.PostAsync("/api/documents/generate", invalidForm);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, invalid.StatusCode); Assert.Null(invalid.Content.Headers.ContentDisposition);
        using var error = JsonDocument.Parse(await invalid.Content.ReadAsStringAsync());
        Assert.Equal(DiagnosticCodes.GenerationInvalidData, error.RootElement.GetProperty("code").GetString());
        Assert.Equal("$.Items[0]", error.RootElement.GetProperty("errors")[0].GetProperty("dataPath").GetString());
        Assert.Contains("A3", error.RootElement.GetProperty("errors")[0].GetProperty("location").GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Static_Existing_routes_select_actual_format_and_return_the_same_analysis(bool spreadsheet)
    {
        var source = spreadsheet ? _fixture.Create() : Docx.Create(Docx.P(Docx.Run("{Name}")));
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var generation = Form(source, "{\"Name\":\"Ada\"}", spreadsheet);
        using var response = await client.PostAsync("/api/documents/generate", generation);
        response.EnsureSuccessStatusCode();
        Assert.Equal("application/vnd.openxmlformats-officedocument." + (spreadsheet ? "spreadsheetml.sheet" : "wordprocessingml.document"), response.Content.Headers.ContentType!.MediaType);
        Assert.Contains(spreadsheet ? "document.xlsx" : "document.docx", response.Content.Headers.ContentDisposition!.ToString());
        var bytes = await response.Content.ReadAsByteArrayAsync();
        if (spreadsheet)
        {
            using var document = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
            Assert.Empty(new OpenXmlValidator().Validate(document));
            Assert.Equal("Ada", document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single().InlineString!.InnerText);
        }
        else Assert.Equal("Ada", Docx.Text(Docx.Read(bytes)));
        using var analyze = Form(source, null, spreadsheet);
        using var analyzed = await client.PostAsync("/api/documents/analyze", analyze);
        analyzed.EnsureSuccessStatusCode();
        using var json = JsonDocument.Parse(await analyzed.Content.ReadAsStringAsync());
        Assert.Equal("$.Name", json.RootElement.GetProperty("occurrences")[0].GetProperty("path").GetString());
        Assert.True(json.RootElement.GetProperty("valid").GetBoolean());
        Assert.Equal(new[] { "errors", "occurrences", "requirements", "valid", "warnings" }, json.RootElement.EnumerateObject().Select(item => item.Name).Order());
    }

    [Theory]
    [InlineData("xml", DiagnosticCodes.TemplateInvalidXml)]
    [InlineData("part", DiagnosticCodes.XlsxInvalidPackage)]
    [InlineData("precision", DiagnosticCodes.XlsxUnsafeNumber)]
    [InlineData("root", DiagnosticCodes.GenerationInvalidData)]
    [InlineData("text", DiagnosticCodes.GenerationInvalidData)]
    [InlineData("high-surrogate", DiagnosticCodes.GenerationInvalidData)]
    [InlineData("low-surrogate", DiagnosticCodes.GenerationInvalidData)]
    [InlineData("formula", DiagnosticCodes.XlsxUnsafeFormula)]
    public async Task Static_Failures_return_coded_problem_details_before_any_report_headers_or_bytes(string scenario, string code)
    {
        var source = _fixture.Create((_, sheet) =>
        {
            if (scenario != "formula") return;
            var cell = sheet.Worksheet!.Descendants<Cell>().Single(); cell.RemoveAllChildren(); cell.DataType = null; cell.CellFormula = new CellFormula("INDIRECT(\"B3\")");
        });
        if (scenario == "xml") source = _fixture.RewritePart(source, "xl/worksheets/sheet1.xml", xml => xml[..^10]);
        if (scenario == "part") source = _fixture.RewritePart(source, "xl/_rels/workbook.xml.rels", xml => xml.Replace("worksheets/sheet1.xml", "worksheets/missing.xml", StringComparison.Ordinal));
        var data = scenario switch
        {
            "precision" => "{\"Name\":10000000000000001}", "root" => "[]", "text" => "{\"Name\":\"private\\u0001\"}",
            "high-surrogate" => "{\"Name\":\"private\\uD800\"}", "low-surrogate" => "{\"Name\":\"private\\uDC00\"}", _ => "{}"
        };
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var form = Form(source, data, true);
        using var response = await client.PostAsync("/api/documents/generate", form);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, response.StatusCode);
        Assert.Equal("application/problem+json", response.Content.Headers.ContentType!.MediaType);
        Assert.Null(response.Content.Headers.ContentDisposition);
        Assert.False(response.Headers.Contains("X-MonoReport-Warning-Count"));
        var text = await response.Content.ReadAsStringAsync();
        Assert.DoesNotContain("10000000000000001", text); Assert.DoesNotContain("private", text);
        using var json = JsonDocument.Parse(text);
        Assert.Equal(code, json.RootElement.GetProperty("code").GetString());
        Assert.Equal(code, json.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
        Assert.NotEqual("The server could not complete the operation.", json.RootElement.GetProperty("detail").GetString());
        if (scenario == "precision")
        {
            var issue = json.RootElement.GetProperty("errors")[0];
            Assert.Equal("$.Name", issue.GetProperty("dataPath").GetString()); Assert.Contains("JSON string", issue.GetProperty("message").GetString());
        }
        if (scenario is "high-surrogate" or "low-surrogate")
        {
            var issue = json.RootElement.GetProperty("errors")[0];
            Assert.Equal("$.Name", issue.GetProperty("dataPath").GetString()); Assert.Contains("B3", issue.GetProperty("location").GetString());
            Assert.Contains("paired surrogate", issue.GetProperty("message").GetString());
        }
        if (scenario is "xml" or "part" or "formula")
        {
            using var analyze = Form(source, null, true);
            using var analyzed = await client.PostAsync("/api/documents/analyze", analyze);
            analyzed.EnsureSuccessStatusCode();
            using var analysis = JsonDocument.Parse(await analyzed.Content.ReadAsStringAsync());
            Assert.False(analysis.RootElement.GetProperty("valid").GetBoolean());
            Assert.Equal(code, analysis.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
        }
    }

    [Fact]
    public async Task Static_Revision_string_fidelity_and_unsupported_excel_remedy_reach_http_clients()
    {
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var success = Form(_fixture.Create(), JsonSerializer.Serialize(new { Name = "a\r\n_x0041_" }), true);
        using var output = await client.PostAsync("/api/documents/generate", success);
        output.EnsureSuccessStatusCode();
        Assert.Equal("a\r\n_x0041_", _fixture.ReadCellText(await output.Content.ReadAsByteArrayAsync(), "A1"));
        var source = _fixture.Create(kind: DocumentFormat.OpenXml.SpreadsheetDocumentType.Template);
        using var form = Form(source, "{}", true);
        using var response = await client.PostAsync("/api/documents/generate", form);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, response.StatusCode); Assert.Null(response.Content.Headers.ContentDisposition);
        using var failure = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        Assert.Equal(DiagnosticCodes.TemplateInvalidPackage, failure.RootElement.GetProperty("code").GetString());
        Assert.Contains("Word", failure.RootElement.GetProperty("detail").GetString()); Assert.Contains("ordinary .xlsx", failure.RootElement.GetProperty("detail").GetString());
        Assert.Contains("ordinary .xlsx", failure.RootElement.GetProperty("errors")[0].GetProperty("message").GetString());
        using var analyze = Form(source, null, true);
        using var analyzed = await client.PostAsync("/api/documents/analyze", analyze);
        analyzed.EnsureSuccessStatusCode();
        using var analysis = JsonDocument.Parse(await analyzed.Content.ReadAsStringAsync());
        Assert.False(analysis.RootElement.GetProperty("valid").GetBoolean());
        Assert.Contains("ordinary .xlsx", analysis.RootElement.GetProperty("errors")[0].GetProperty("message").GetString());
    }

    private static MultipartFormDataContent Form(byte[] source, string? data, bool spreadsheet)
    {
        var form = new MultipartFormDataContent(); var file = new ByteArrayContent(source);
        file.Headers.ContentType = new MediaTypeHeaderValue(spreadsheet ? "application/msword" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        form.Add(file, "template", spreadsheet ? "misleading.docx" : "misleading.xlsx");
        if (data is not null) form.Add(new StringContent(data), "data");
        return form;
    }
}
