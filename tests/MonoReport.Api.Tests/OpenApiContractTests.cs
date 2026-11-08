using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using MonoReport.Api.Services;
using MonoReport.Cli.Tests;
using MonoReport.Core.Diagnostics;
using Xunit;

namespace MonoReport.Api.Tests;

public sealed class OpenApiContractTests
{
    [Fact]
    public async Task Formatter_problem_responses_document_diagnostic_fields_codes_and_validation_errors()
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        using var response = await client.GetAsync("/openapi/v1.json");
        response.EnsureSuccessStatusCode();
        using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        var paths = json.RootElement.GetProperty("paths");
        var execute = paths.GetProperty("/api/formatters/{name}/execute").GetProperty("post").GetProperty("responses");
        var expectedCodes = new[]
        {
            ("400", new[]
            {
                ApiDiagnosticCodes.RequestMissingInput, ApiDiagnosticCodes.RequestInvalidInput,
                ApiDiagnosticCodes.CultureInvalid
            }),
            ("404", new[] {ApiDiagnosticCodes.FormatterNotFound}),
            ("413", new[] {ApiDiagnosticCodes.RequestTooLarge}),
            ("415", new[] {ApiDiagnosticCodes.RequestUnsupportedMediaType}),
            ("422", new[] {ApiDiagnosticCodes.FormatterInvalidValue}),
            ("500", new[] {ApiDiagnosticCodes.ServerInternalError})
        };
        foreach (var (status, codes) in expectedCodes)
        {
            var documented = execute.GetProperty(status);
            var schema = documented.GetProperty("content").GetProperty("application/problem+json").GetProperty("schema");
            var properties = schema.GetProperty("properties");
            foreach (var name in new[] {"type", "title", "status", "detail", "code", "traceId", "field"})
                Assert.True(properties.TryGetProperty(name, out _), $"{status} is missing {name}.");
            foreach (var code in codes)
            {
                Assert.Contains(code, documented.GetProperty("description").GetString());
                Assert.Contains(code, properties.GetProperty("code").GetProperty("description").GetString());
            }

            Assert.Contains(schema.GetProperty("required").EnumerateArray(), item => item.GetString() == "code");
            Assert.Contains(schema.GetProperty("required").EnumerateArray(), item => item.GetString() == "traceId");
        }

        var badRequest = execute.GetProperty("400").GetProperty("content").GetProperty("application/problem+json")
            .GetProperty("schema");
        var errors = badRequest.GetProperty("properties").GetProperty("errors");
        Assert.Equal("array", errors.GetProperty("type").GetString());
        Assert.Equal("string", errors.GetProperty("items").GetProperty("type").GetString());
        Assert.DoesNotContain(badRequest.GetProperty("required").EnumerateArray(), item => item.GetString() == "errors");
        var listFailure = paths.GetProperty("/api/formatters").GetProperty("get").GetProperty("responses")
            .GetProperty("500").GetProperty("content").GetProperty("application/problem+json").GetProperty("schema");
        Assert.Contains(ApiDiagnosticCodes.ServerInternalError,
            listFailure.GetProperty("properties").GetProperty("code").GetProperty("description").GetString());

        using var request = new StringContent("""{"value":1}""", Encoding.UTF8, "application/json");
        using var failure = await client.PostAsync("/api/formatters/number/execute", request);
        Assert.Equal(HttpStatusCode.BadRequest, failure.StatusCode);
        using var problem = JsonDocument.Parse(await failure.Content.ReadAsStringAsync());
        Assert.All(problem.RootElement.EnumerateObject(), property =>
            Assert.True(badRequest.GetProperty("properties").TryGetProperty(property.Name, out _)));
        Assert.All(problem.RootElement.GetProperty("errors").EnumerateArray(), item =>
            Assert.Equal(JsonValueKind.String, item.ValueKind));
    }

    [Fact]
    public async Task OpenApi_exposes_the_minimal_analyzer_contract_and_generation_diagnostics()
    {
        using var factory =
            new WebApplicationFactory<Program>().WithWebHostBuilder(builder => builder.UseEnvironment("Development"));
        using var client = factory.CreateClient();
        using var response = await client.GetAsync("/openapi/v1.json");
        response.EnsureSuccessStatusCode();
        using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        var root = json.RootElement;
        var schemas = root.GetProperty("components").GetProperty("schemas");
        var analysis = schemas.GetProperty("TemplateAnalysisResult");
        Assert.Equal(new[] {"errors", "occurrences", "requirements", "valid", "warnings"},
            analysis.GetProperty("required").EnumerateArray().Select(item => item.GetString()).Order());

        var requirement = schemas.GetProperty("TemplateRequirement");
        Assert.Equal(new[] {"name", "type"},
            requirement.GetProperty("required").EnumerateArray().Select(item => item.GetString()).Order());
        Assert.True(requirement.GetProperty("properties").TryGetProperty("fields", out _));
        foreach (var forbidden in new[] {"items", "elementType", "required", "nullable", "format"})
            Assert.False(requirement.GetProperty("properties").TryGetProperty(forbidden, out _));

        var occurrence = schemas.GetProperty("TemplateOccurrence");
        Assert.Equal(new[] {"expression", "kind", "location", "part", "path"},
            occurrence.GetProperty("required").EnumerateArray().Select(item => item.GetString()).Order());

        var generate = root.GetProperty("paths").GetProperty("/api/documents/render").GetProperty("post");
        var errors = generate.GetProperty("responses").GetProperty("422").GetProperty("content")
            .GetProperty("application/problem+json").GetProperty("schema").GetProperty("properties")
            .GetProperty("errors");
        foreach (var name in new[] {"code", "message", "field", "part", "location", "formatter", "dataPath"})
            Assert.True(errors.GetProperty("items").GetProperty("properties").TryGetProperty(name, out _));
    }

    [Fact]
    public async Task OpenApi_error_code_lists_match_the_endpoint_diagnostic_contract()
    {
        using var factory =
            new WebApplicationFactory<Program>().WithWebHostBuilder(builder => builder.UseEnvironment("Development"));
        using var client = factory.CreateClient();
        using var response = await client.GetAsync("/openapi/v1.json");
        response.EnsureSuccessStatusCode();
        using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        var generate = json.RootElement.GetProperty("paths").GetProperty("/api/documents/render")
            .GetProperty("post").GetProperty("responses");
        var analyze = json.RootElement.GetProperty("paths").GetProperty("/api/documents/analyze")
            .GetProperty("post").GetProperty("responses");

        static string[] Codes(JsonElement responses, string status)
        {
            return responses
                .GetProperty(status).GetProperty("content").GetProperty("application/problem+json")
                .GetProperty("schema").GetProperty("properties").GetProperty("code")
                .GetProperty("description").GetString()!.Split(", ");
        }

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
        Assert.All(generationCodes,
            code => Assert.NotEqual("The server could not complete the operation.", problemFactory.SafeMessage(code)));

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

        using var form = HttpApiTests.Form(DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Raw}"))), "{}");
        using var failure = await client.PostAsync("/api/documents/render", form);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, failure.StatusCode);
        using var problem = JsonDocument.Parse(await failure.Content.ReadAsStringAsync());
        var actual = problem.RootElement.GetProperty("code").GetString()!;
        Assert.Contains(actual, generationCodes);
    }

    [Fact]
    public async Task OpenApi_documents_both_binary_formats_and_formatter_routes()
    {
        using var factory =
            new WebApplicationFactory<Program>().WithWebHostBuilder(builder => builder.UseEnvironment("Development"));
        using var client = factory.CreateClient();
        using var response = await client.GetAsync("/openapi/v1.json");
        response.EnsureSuccessStatusCode();
        using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        var paths = json.RootElement.GetProperty("paths");
        Assert.Equal(new[] {"/api/documents/analyze", "/api/documents/render", "/api/formatters", "/api/formatters/{name}/execute"},
            paths.EnumerateObject().Select(item => item.Name).Order());
        var list = paths.GetProperty("/api/formatters").GetProperty("get");
        Assert.False(list.TryGetProperty("requestBody", out _));
        Assert.True(list.GetProperty("responses").GetProperty("200").GetProperty("content")
            .TryGetProperty("application/json", out _));
        var execute = paths.GetProperty("/api/formatters/{name}/execute").GetProperty("post");
        Assert.Equal(new[] {"application/json"}, execute.GetProperty("requestBody").GetProperty("content")
            .EnumerateObject().Select(item => item.Name));
        Assert.True(execute.GetProperty("responses").TryGetProperty("404", out _));
        var operation = paths.GetProperty("/api/documents/render").GetProperty("post");
        var content = operation.GetProperty("responses").GetProperty("200").GetProperty("content");
        Assert.Equal(
            new[]
            {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "multipart/mixed"
            },
            content.EnumerateObject().Select(item => item.Name).Order());
        Assert.Contains("static", operation.GetProperty("description").GetString());
        Assert.Contains("scalar formulas", operation.GetProperty("description").GetString());
        var description = operation.GetProperty("description").GetString();
        foreach (var phrase in new[]
                 {
                     "branching sibling blocks", "multirow headers", "q=0", "binary wins ties", "fall back to binary",
                     "Vary: Accept"
                 })
            Assert.Contains(phrase, description);
        Assert.DoesNotContain("at most two horizontal levels", description);
        var success = operation.GetProperty("responses").GetProperty("200");
        Assert.True(success.GetProperty("headers").TryGetProperty("Vary", out _));
        Assert.Contains("only when",
            success.GetProperty("headers").GetProperty("X-MonoReport-Warning-Count").GetProperty("description")
                .GetString());
        var diagnostics = content.GetProperty("multipart/mixed").GetProperty("schema").GetProperty("properties")
            .GetProperty("diagnostics");
        Assert.Contains(operation.GetProperty("parameters").EnumerateArray(),
            parameter => parameter.GetProperty("name").GetString() == "Accept" &&
                         parameter.GetProperty("in").GetString() == "header");
        Assert.Equal("application/json",
            content.GetProperty("multipart/mixed").GetProperty("encoding").GetProperty("diagnostics")
                .GetProperty("contentType").GetString());
        Assert.Equal(new[] {"fileName", "format", "warnings"},
            diagnostics.GetProperty("required").EnumerateArray().Select(p => p.GetString()).Order());
        var warnings = diagnostics.GetProperty("properties").GetProperty("warnings");
        Assert.Contains("empty array", warnings.GetProperty("description").GetString());
        Assert.Equal(new[]
            {"code", "dataPath", "field", "fieldCode", "location", "message", "part", "removedCharacterCount"},
            warnings.GetProperty("items").GetProperty("properties").EnumerateObject().Select(p => p.Name).Order());
        Assert.DoesNotContain("code",
            warnings.GetProperty("items").GetProperty("required").EnumerateArray().Select(p => p.GetString()));
        Assert.Contains(DiagnosticCodes.XlsxUnsafeFormula,
            operation.GetProperty("responses").GetProperty("422").GetProperty("description").GetString());
    }
}
