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
            Docx.P(Docx.Run("{/For}")));

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
        var template = Docx.Create(Docx.P(Docx.Run("{Raw}")));

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

    [Fact]
    public async Task Analyze_and_generate_preserve_source_context_without_inventing_a_split_expression()
    {
        using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        var template = Docx.Create(Docx.P(Docx.Run("Invoice reference: {Na"),
            new System.Xml.Linq.XElement(MonoReport.Services.WordXml.WordprocessingNamespace + "r",
                new System.Xml.Linq.XElement(MonoReport.Services.WordXml.WordprocessingNamespace + "tab")), Docx.Run("me}")));
        using var analyze = Form(template);
        using var analyzed = await client.PostAsync("/api/documents/analyze", analyze);
        analyzed.EnsureSuccessStatusCode();
        using var analysis = JsonDocument.Parse(await analyzed.Content.ReadAsStringAsync());
        var errors = analysis.RootElement.GetProperty("errors").EnumerateArray().ToArray();
        Assert.Equal(new[] { "{Na", "}" }, errors.Select(item => item.GetProperty("field").GetString()));
        Assert.All(errors, item =>
        {
            Assert.Contains("Invoice reference:", item.GetProperty("message").GetString());
            Assert.DoesNotContain("{Name}", item.GetProperty("message").GetString());
            Assert.Equal("/word/document.xml, paragraph 1", item.GetProperty("location").GetString());
        });
        using var generate = Form(template, "{}");
        using var generated = await client.PostAsync("/api/documents/generate", generate);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, generated.StatusCode);
        using var problem = JsonDocument.Parse(await generated.Content.ReadAsStringAsync());
        var generationErrors = problem.RootElement.GetProperty("errors").EnumerateArray().ToArray();
        Assert.Equal(errors.Length, generationErrors.Length);
        for (var index = 0; index < errors.Length; index++)
            foreach (var property in new[] { "code", "message", "field", "part", "location" })
                Assert.Equal(errors[index].GetProperty(property).GetString(), generationErrors[index].GetProperty(property).GetString());
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

    [Theory]
    [InlineData("crossing-bookmark")] [InlineData("external-bookmark-link")] [InlineData("nested-bookmark-link")]
    [InlineData("tracked-cell-deletion")] [InlineData("tracked-move-range")] [InlineData("expression-crosses-content-control-exit")]
    public async Task Final_corrections_preserve_diagnostic_parity_and_return_problem_details_without_docx(string name)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "review-reproductions", name);
        var template = File.ReadAllBytes(path + ".template.docx");
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var analyze = Form(template); using var analyzed = await client.PostAsync("/api/documents/analyze", analyze);
        Assert.Equal(HttpStatusCode.OK, analyzed.StatusCode);
        using var analysis = JsonDocument.Parse(await analyzed.Content.ReadAsStringAsync());
        Assert.False(analysis.RootElement.GetProperty("valid").GetBoolean());
        using var generate = Form(template, File.ReadAllText(path + ".json"));
        using var generated = await client.PostAsync("/api/documents/generate", generate);
        Assert.Equal(HttpStatusCode.UnprocessableEntity, generated.StatusCode);
        Assert.Equal("application/problem+json", generated.Content.Headers.ContentType?.MediaType);
        Assert.Null(generated.Content.Headers.ContentDisposition);
        using var problem = JsonDocument.Parse(await generated.Content.ReadAsStringAsync());
        var errors = analysis.RootElement.GetProperty("errors").EnumerateArray().ToArray();
        var generationErrors = problem.RootElement.GetProperty("errors").EnumerateArray().ToArray();
        Assert.Equal(errors.Length, generationErrors.Length);
        for (var i = 0; i < errors.Length; i++)
            foreach (var property in new[] { "code", "message", "part", "location" })
                Assert.Equal(errors[i].GetProperty(property).GetString(), generationErrors[i].GetProperty(property).GetString());
    }

    [Fact]
    public async Task Raw_literal_and_following_expression_generate_through_http()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Raw}")), Docx.P(Docx.Run("{Hidden}")),
            Docx.P(Docx.Run("{/Raw}")), Docx.P(Docx.Run("{Name}")));
        using var factory = new WebApplicationFactory<Program>(); using var client = factory.CreateClient();
        using var form = Form(template, "{\"Name\":\"OK\"}");
        using var response = await client.PostAsync("/api/documents/generate", form);
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var output = await response.Content.ReadAsByteArrayAsync();
        Docx.AssertValid(output, DocumentFormat.OpenXml.FileFormatVersions.Office2010);
        Assert.Equal("{Hidden}OK", Docx.Text(Docx.Read(output)));
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
