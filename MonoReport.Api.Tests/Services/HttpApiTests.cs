using System.Net;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc.Testing;

namespace MonoReport.Tests.Services;

public sealed class HttpApiTests
{
    [Fact]
    public async Task Analyze_returns_ordered_requirements_and_all_occurrences()
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        var template = Docx.Create(
            Docx.P(Docx.Run("{Customer.Name} {Customer.Name}")),
            Docx.P(Docx.Run("{For:Tags}")),
            Docx.P(Docx.Run("{.}")),
            Docx.P(Docx.Run("{EndFor}")));

        using var form = Form(template);
        using var response = await client.PostAsync("/api/documents/analyze", form);

        response.EnsureSuccessStatusCode();
        using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        var root = json.RootElement;
        Assert.True(root.GetProperty("valid").GetBoolean());
        Assert.Equal(new[] { "Customer", "Tags" }, root.GetProperty("requirements").EnumerateArray().Select(item => item.GetProperty("name").GetString()));
        var tags = root.GetProperty("requirements")[1];
        Assert.Equal("array", tags.GetProperty("type").GetString());
        Assert.False(tags.TryGetProperty("fields", out _));
        Assert.Equal(new[] { "$.Customer.Name", "$.Customer.Name", "$.Tags", "$.Tags[]" },
            root.GetProperty("occurrences").EnumerateArray().Select(item => item.GetProperty("path").GetString()));
    }

    [Fact]
    public async Task Generate_renders_unified_syntax_and_returns_typed_data_diagnostics()
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        var template = Docx.Create(Docx.P(Docx.Run("{Amount number:\"0.00\"}")));
        using (var form = Form(template, "{\"Amount\":1234.5}"))
        using (var response = await client.PostAsync("/api/documents/generate", form))
        {
            response.EnsureSuccessStatusCode();
            Assert.Equal("1234.50", Docx.Text(Docx.Read(await response.Content.ReadAsByteArrayAsync())));
        }

        using var invalid = Form(template, "{\"Amount\":true}");
        using var failure = await client.PostAsync("/api/documents/generate", invalid);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, failure.StatusCode);
        using var problem = JsonDocument.Parse(await failure.Content.ReadAsStringAsync());
        Assert.Equal("generation.invalid_format", problem.RootElement.GetProperty("code").GetString());
        var issue = Assert.Single(problem.RootElement.GetProperty("errors").EnumerateArray());
        Assert.Equal("$.Amount", issue.GetProperty("dataPath").GetString());
        Assert.Equal("number", issue.GetProperty("formatter").GetString());
    }

    [Fact]
    public async Task Analyze_and_generate_reject_the_same_invalid_template_before_output()
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        var template = Docx.Create(Docx.P(Docx.Run("{EndFor}")));

        using (var form = Form(template))
        using (var response = await client.PostAsync("/api/documents/analyze", form))
        {
            response.EnsureSuccessStatusCode();
            using var result = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
            Assert.False(result.RootElement.GetProperty("valid").GetBoolean());
            Assert.Equal("template.invalid_structure", result.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
        }

        using var generate = Form(template, "{}");
        using var failure = await client.PostAsync("/api/documents/generate", generate);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, failure.StatusCode);
        using var problem = JsonDocument.Parse(await failure.Content.ReadAsStringAsync());
        Assert.Equal("template.invalid_structure", problem.RootElement.GetProperty("code").GetString());
    }

    [Theory]
    [InlineData("analyze")]
    [InlineData("generate")]
    public async Task Request_contract_rejects_non_multipart_input(string endpoint)
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        using var response = await client.PostAsync($"/api/documents/{endpoint}", new StringContent("{}"));
        Assert.Equal(HttpStatusCode.UnsupportedMediaType, response.StatusCode);
        using var problem = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        Assert.Equal("request.unsupported_media_type", problem.RootElement.GetProperty("code").GetString());
    }

    public static MultipartFormDataContent Form(byte[] template, string? data = null)
    {
        var form = new MultipartFormDataContent();
        var content = new ByteArrayContent(template);
        content.Headers.ContentType = new("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        form.Add(content, "template", "template.docx");
        if (data is not null) form.Add(new StringContent(data), "data");
        return form;
    }
}
