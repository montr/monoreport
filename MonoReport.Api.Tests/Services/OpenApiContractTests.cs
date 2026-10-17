using System.Text.Json;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;

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
}
