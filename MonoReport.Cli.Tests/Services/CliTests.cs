using System.Text;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Cli.Services;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class CliTests
{
    [Fact]
    public void Analyze_writes_the_new_contract_and_does_not_read_json()
    {
        var path = WriteTemplate(Docx.Create(Docx.P(Docx.Run("{For:Tags}")), Docx.P(Docx.Run("{.}")), Docx.P(Docx.Run("{EndFor}"))));
        try
        {
            using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
            var app = new CliApplication(new CliOptionsParser(), provider.GetRequiredService<MergeEngine>(), provider.GetRequiredService<TemplateAnalyzer>());
            using var output = new MemoryStream();
            using var errors = new StringWriter();

            Assert.Equal(0, app.Run(["--analyze", "--template", path], Stream.Null, output, errors));
            Assert.Empty(errors.ToString());
            using var json = JsonDocument.Parse(output.ToArray());
            Assert.Equal(new[] { "errors", "occurrences", "requirements", "valid", "warnings" },
                json.RootElement.EnumerateObject().Select(item => item.Name).Order());
            var requirement = Assert.Single(json.RootElement.GetProperty("requirements").EnumerateArray());
            Assert.Equal("Tags", requirement.GetProperty("name").GetString());
            Assert.False(requirement.TryGetProperty("fields", out _));
            Assert.Equal("$.Tags[]", json.RootElement.GetProperty("occurrences")[1].GetProperty("path").GetString());
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Generate_writes_docx_and_reports_location_bearing_format_errors_without_partial_output()
    {
        var path = WriteTemplate(Docx.Create(Docx.P(Docx.Run("{Amount number:\"0.00\"}"))));
        try
        {
            using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
            var app = new CliApplication(new CliOptionsParser(), provider.GetRequiredService<MergeEngine>(), provider.GetRequiredService<TemplateAnalyzer>());
            using var output = new MemoryStream();
            using var errors = new StringWriter();
            Assert.Equal(1, app.Run(["--template", path, "--json", "{\"Amount\":true}"], Stream.Null, output, errors));
            Assert.Empty(output.ToArray());
            Assert.Contains("generation.invalid_format", errors.ToString());
            Assert.Contains("formatter=number", errors.ToString());
            Assert.Contains("location=/word/document.xml", errors.ToString());
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Generate_supports_inline_file_and_stdin_json_sources()
    {
        var template = WriteTemplate(Docx.Create(Docx.P(Docx.Run("{Name upper}"))));
        var jsonPath = Path.GetTempFileName();
        File.WriteAllText(jsonPath, "{\"Name\":\"Ada\"}");
        try
        {
            foreach (var arguments in new[]
            {
                new[] { "--template", template, "--json", "{\"Name\":\"Ada\"}" },
                new[] { "--template", template, "--json-file", jsonPath },
                new[] { "--template", template, "--stdin" }
            })
            {
                using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
                var app = new CliApplication(new CliOptionsParser(), provider.GetRequiredService<MergeEngine>(), provider.GetRequiredService<TemplateAnalyzer>());
                using var input = new MemoryStream(Encoding.UTF8.GetBytes("{\"Name\":\"Ada\"}"));
                using var output = new MemoryStream();
                using var errors = new StringWriter();
                Assert.Equal(0, app.Run(arguments, input, output, errors));
                Assert.Empty(errors.ToString());
                Assert.Equal("ADA", Docx.Text(Docx.Read(output.ToArray())));
            }
        }
        finally { File.Delete(template); File.Delete(jsonPath); }
    }

    [Fact]
    public void Parser_rejects_generation_options_in_analyze_mode()
    {
        var parser = new CliOptionsParser();
        Assert.Throws<ArgumentException>(() => parser.Parse(["--analyze", "--template", "a.docx", "--json", "{}"]));
        Assert.Throws<ArgumentException>(() => parser.Parse(["--analyze"]));
    }

    [Fact]
    public void Single_template_error_keeps_token_context_and_location_in_analysis_and_generation()
    {
        var path = WriteTemplate(Docx.Create(Docx.P(Docx.Run("Invoice reference: {Name"))));
        try
        {
            using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
            var app = new CliApplication(new CliOptionsParser(), provider.GetRequiredService<MergeEngine>(), provider.GetRequiredService<TemplateAnalyzer>());
            using var output = new MemoryStream();
            using var errors = new StringWriter();
            Assert.Equal(1, app.Run(["--analyze", "--template", path], Stream.Null, output, errors));
            using var analysis = JsonDocument.Parse(output.ToArray());
            var issue = Assert.Single(analysis.RootElement.GetProperty("errors").EnumerateArray());
            var message = issue.GetProperty("message").GetString()!;
            Assert.Equal("{Name", issue.GetProperty("field").GetString());
            Assert.Contains("Invoice reference:", message);
            Assert.Equal("/word/document.xml, paragraph 1", issue.GetProperty("location").GetString());
            using var generation = new MemoryStream();
            Assert.Equal(1, app.Run(["--template", path, "--json", "{}"], Stream.Null, generation, errors));
            Assert.Empty(generation.ToArray());
            Assert.Contains(message, errors.ToString());
            Assert.Contains("template.unmatched_brace", errors.ToString());
            Assert.Contains("field={Name", errors.ToString());
            Assert.Contains("location=/word/document.xml, paragraph 1", errors.ToString());
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("crossing-bookmark")] [InlineData("external-bookmark-link")] [InlineData("nested-bookmark-link")]
    [InlineData("tracked-cell-deletion")] [InlineData("tracked-move-range")] [InlineData("expression-crosses-content-control-exit")]
    public void Final_corrections_preserve_analyze_generate_diagnostics_and_never_write_partial_output(string name)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "review-reproductions", name + ".template.docx");
        var json = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "review-reproductions", name + ".json"));
        var outputPath = Path.Combine(Path.GetTempPath(), "feature010-rejected-" + Guid.NewGuid().ToString("N") + ".docx");
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        var app = new CliApplication(new CliOptionsParser(), provider.GetRequiredService<MergeEngine>(), provider.GetRequiredService<TemplateAnalyzer>());
        using var analysisBytes = new MemoryStream(); using var diagnostics = new StringWriter();
        Assert.Equal(1, app.Run(["--analyze", "--template", path], Stream.Null, analysisBytes, diagnostics));
        using var analysis = JsonDocument.Parse(analysisBytes.ToArray());
        Assert.False(analysis.RootElement.GetProperty("valid").GetBoolean());
        foreach (var outputArguments in new[] { Array.Empty<string>(), new[] { "--output", outputPath } })
        {
            using var output = new MemoryStream(); using var errors = new StringWriter();
            Assert.Equal(1, app.Run(["--template", path, "--json", json, .. outputArguments], Stream.Null, output, errors));
            Assert.Empty(output.ToArray()); Assert.False(File.Exists(outputPath));
            foreach (var issue in analysis.RootElement.GetProperty("errors").EnumerateArray())
            {
                Assert.Contains(issue.GetProperty("code").GetString()!, errors.ToString());
                Assert.Contains(issue.GetProperty("message").GetString()!, errors.ToString());
                Assert.Contains(issue.GetProperty("location").GetString()!, errors.ToString());
            }
        }
    }

    [Fact]
    public void Comment_image_correction_generates_valid_output_through_cli()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "review-reproductions", "comment-with-image.template.docx");
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        var app = new CliApplication(new CliOptionsParser(), provider.GetRequiredService<MergeEngine>(), provider.GetRequiredService<TemplateAnalyzer>());
        using var output = new MemoryStream(); using var errors = new StringWriter();
        Assert.Equal(0, app.Run(["--template", path, "--json", "{\"Items\":[\"A\",\"B\"]}"], Stream.Null, output, errors));
        Assert.Empty(errors.ToString());
        Docx.AssertValid(output.ToArray(), DocumentFormat.OpenXml.FileFormatVersions.Office2010);
        Assert.Equal(2, Docx.Read(output.ToArray()).Descendants(MonoReport.Services.WordXml.WordprocessingNamespace + "commentReference").Count());
    }

    private static string WriteTemplate(byte[] bytes)
    {
        var path = Path.Combine(Path.GetTempPath(), $"monoreport-cli-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, bytes);
        return path;
    }
}
