using System.Net;
using System.Text.Json;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using MonoReport.Api.Services;
using MonoReport.Models;

namespace MonoReport.Tests.Services;

public sealed class OpenApiContractTests
{
    [Fact]
    public async Task OpenApi_exposes_the_minimal_analyzer_contract_and_generation_diagnostics()
    {
        using var factory = new WebApplicationFactory<Program>().WithWebHostBuilder(builder => builder.UseEnvironment("Development"));
        using var client = factory.CreateClient();
        using var response = await client.GetAsync("/openapi/v1.json");
        response.EnsureSuccessStatusCode();
        using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        var root = json.RootElement;
        var schemas = root.GetProperty("components").GetProperty("schemas");
        var analysis = schemas.GetProperty("TemplateAnalysisResult");
        Assert.Equal(new[] { "errors", "occurrences", "requirements", "valid", "warnings" },
            analysis.GetProperty("required").EnumerateArray().Select(item => item.GetString()).Order());

        var requirement = schemas.GetProperty("TemplateRequirement");
        Assert.Equal(new[] { "name", "type" }, requirement.GetProperty("required").EnumerateArray().Select(item => item.GetString()).Order());
        Assert.True(requirement.GetProperty("properties").TryGetProperty("fields", out _));
        foreach (var forbidden in new[] { "items", "elementType", "required", "nullable", "format" })
            Assert.False(requirement.GetProperty("properties").TryGetProperty(forbidden, out _));

        var occurrence = schemas.GetProperty("TemplateOccurrence");
        Assert.Equal(new[] { "expression", "kind", "location", "part", "path" },
            occurrence.GetProperty("required").EnumerateArray().Select(item => item.GetString()).Order());

        var generate = root.GetProperty("paths").GetProperty("/api/documents/generate").GetProperty("post");
        var errors = generate.GetProperty("responses").GetProperty("422").GetProperty("content")
            .GetProperty("application/problem+json").GetProperty("schema").GetProperty("properties").GetProperty("errors");
        foreach (var name in new[] { "code", "message", "field", "part", "location", "formatter", "dataPath" })
            Assert.True(errors.GetProperty("items").GetProperty("properties").TryGetProperty(name, out _));
    }

    [Fact]
    public async Task OpenApi_error_code_lists_match_the_endpoint_diagnostic_contract()
    {
        using var factory = new WebApplicationFactory<Program>().WithWebHostBuilder(builder => builder.UseEnvironment("Development"));
        using var client = factory.CreateClient();
        using var response = await client.GetAsync("/openapi/v1.json");
        response.EnsureSuccessStatusCode();
        using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        var generate = json.RootElement.GetProperty("paths").GetProperty("/api/documents/generate")
            .GetProperty("post").GetProperty("responses");
        var analyze = json.RootElement.GetProperty("paths").GetProperty("/api/documents/analyze")
            .GetProperty("post").GetProperty("responses");

        static string[] Codes(JsonElement responses, string status) => responses
            .GetProperty(status).GetProperty("content").GetProperty("application/problem+json")
            .GetProperty("schema").GetProperty("properties").GetProperty("code")
            .GetProperty("description").GetString()!.Split(", ");

        var generationCodes = new[]
        {
            DiagnosticCodes.TemplateInvalidPackage, DiagnosticCodes.TemplateInvalidXml,
            DiagnosticCodes.TemplateUnsupportedStructure, DiagnosticCodes.TemplateInvalidExpression,
            DiagnosticCodes.TemplateUnmatchedBrace, DiagnosticCodes.TemplateInvalidContext,
            DiagnosticCodes.TemplateInvalidStructure, DiagnosticCodes.TemplateRequirementConflict,
            DiagnosticCodes.TemplateTrackedRevision, DiagnosticCodes.GenerationInvalidData,
            DiagnosticCodes.GenerationInvalidFormat, DiagnosticCodes.XlsxInvalidPrintArea,
            DiagnosticCodes.XlsxUnsupportedFeature, DiagnosticCodes.XlsxDimensionLimit,
            DiagnosticCodes.XlsxUnsafeNumber, DiagnosticCodes.XlsxInvalidPackage, DiagnosticCodes.XlsxUnsafeFormula
        };
        Assert.Equal(generationCodes, Codes(generate, "422"));
        var problemFactory = new ApiProblemFactory();
        Assert.All(generationCodes, code => Assert.NotEqual("The server could not complete the operation.", problemFactory.SafeMessage(code)));

        Assert.Equal(new[]
        {
            ApiDiagnosticCodes.RequestMissingInput, ApiDiagnosticCodes.RequestEmptyInput,
            ApiDiagnosticCodes.RequestDuplicateInput, ApiDiagnosticCodes.RequestInvalidInput,
            ApiDiagnosticCodes.RequestInvalidMultipart, ApiDiagnosticCodes.DataInvalidJson,
            ApiDiagnosticCodes.CultureInvalid
        }, Codes(generate, "400"));
        Assert.Equal(new[]
        {
            ApiDiagnosticCodes.RequestMissingInput, ApiDiagnosticCodes.RequestEmptyInput,
            ApiDiagnosticCodes.RequestDuplicateInput, ApiDiagnosticCodes.RequestInvalidInput,
            ApiDiagnosticCodes.RequestInvalidMultipart, ApiDiagnosticCodes.RequestUnexpectedInput
        }, Codes(analyze, "400"));
        Assert.Equal([ApiDiagnosticCodes.RequestUnsupportedMediaType], Codes(generate, "415"));
        Assert.Equal([ApiDiagnosticCodes.ServerInternalError], Codes(generate, "500"));

        using var form = HttpApiTests.Form(Docx.Create(Docx.P(Docx.Run("{Raw}"))), "{}");
        using var failure = await client.PostAsync("/api/documents/generate", form);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, failure.StatusCode);
        using var problem = JsonDocument.Parse(await failure.Content.ReadAsStringAsync());
        var actual = problem.RootElement.GetProperty("code").GetString()!;
        Assert.Contains(actual, generationCodes);
    }

    [Fact]
    public async Task OpenApi_documents_both_binary_formats_on_existing_routes_only()
    {
        using var factory = new WebApplicationFactory<Program>().WithWebHostBuilder(builder => builder.UseEnvironment("Development"));
        using var client = factory.CreateClient();
        using var response = await client.GetAsync("/openapi/v1.json");
        response.EnsureSuccessStatusCode();
        using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        var paths = json.RootElement.GetProperty("paths");
        Assert.Equal(new[] { "/api/documents/analyze", "/api/documents/generate" }, paths.EnumerateObject().Select(item => item.Name).Order());
        var operation = paths.GetProperty("/api/documents/generate").GetProperty("post");
        var content = operation.GetProperty("responses").GetProperty("200").GetProperty("content");
        Assert.Equal(new[] { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
            content.EnumerateObject().Select(item => item.Name).Order());
        Assert.Contains("static", operation.GetProperty("description").GetString());
        Assert.Contains("scalar formulas", operation.GetProperty("description").GetString());
        Assert.Contains(DiagnosticCodes.XlsxUnsafeFormula, operation.GetProperty("responses").GetProperty("422").GetProperty("description").GetString());
    }
}
