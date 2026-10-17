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

    private static string WriteTemplate(byte[] bytes)
    {
        var path = Path.Combine(Path.GetTempPath(), $"monoreport-cli-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, bytes);
        return path;
    }
}
