using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Cli.Tests;
using MonoReport.Core.Formatting;
using MonoReport.Tests.Xlsx;
using Xunit;
using CapsFormatter = MonoReport.Core.Formatting.CapsFormatter;
using FirstCapFormatter = MonoReport.Core.Formatting.FirstCapFormatter;

namespace MonoReport.Tests.Core;

public sealed class FormatterPluginPathTests
{
    [Fact]
    public void Custom_keyed_formatter_binds_literals_and_runs_through_docx_and_xlsx_analysis_and_generation()
    {
        var formatter = new ProbeFormatter();
        using var provider = new ServiceCollection().AddMonoReport()
            .AddKeyedSingleton<IFormatter>("probe", formatter).BuildServiceProvider();
        var docx = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{If:Enabled}")),
            DocxFixture.P(DocxFixture.Run("{Name probe \"P\" suffix=\"!\"}")), DocxFixture.P(DocxFixture.Run("{/If}")),
            DocxFixture.P(DocxFixture.Run("{Raw}")), DocxFixture.P(DocxFixture.Run("{Name probe bare}")),
            DocxFixture.P(DocxFixture.Run("{/Raw}")));
        var xlsx = new XlsxFixture().Create((_, worksheet) =>
        {
            var cell = worksheet.Worksheet!.Descendants<Cell>().Single();
            cell.InlineString = new InlineString(
                new Text("{Name probe first=\"P\" suffix=\"!\"}"));
        });
        using var docxInput = new MemoryStream(docx);
        using var xlsxInput = new MemoryStream(xlsx);
        Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(docxInput).Valid);
        var xlsxAnalysis = provider.GetRequiredService<TemplateAnalyzer>().Analyze(xlsxInput);
        Assert.True(xlsxAnalysis.Valid, string.Join("; ", xlsxAnalysis.Errors.Select(error => error.Message)));
        Assert.Equal(0, formatter.FormatCalls);
        var docxText =
            DocxFixture.Text(DocxFixture.Read(Generate(provider, docx, """{"Name":"Ada","Enabled":true}""")));
        Assert.Contains("P:!Ada", docxText);
        Assert.Contains("{Name probe bare}", docxText);
        var xlsxOutput = Generate(provider, xlsx, """{"Name":"Ada"}""");
        using (var package = SpreadsheetDocument.Open(new MemoryStream(xlsxOutput), false))
        {
            var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToArray();
            Assert.Single(cells);
            Assert.Equal("P:!Ada", cells[0].InnerText);
        }

        Assert.Equal(2, formatter.FormatCalls);
    }

    [Fact]
    public void Same_key_replacement_controls_metadata_validation_and_output_together()
    {
        using var provider = new ServiceCollection().AddMonoReport()
            .AddKeyedSingleton<IFormatter>("number", new ReplacementFormatter()).BuildServiceProvider();
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Value number replacement=\"ok\"}")));
        using var input = new MemoryStream(template);
        var analysis = provider.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.True(analysis.Valid, string.Join("; ", analysis.Errors.Select(error => error.Message)));
        Assert.Equal("replaced:7", DocxFixture.Text(DocxFixture.Read(Generate(provider, template, """{"Value":7}"""))));

        var rejected = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Value number \"0.00\"}")));
        using var rejectedInput = new MemoryStream(rejected);
        Assert.False(provider.GetRequiredService<TemplateAnalyzer>().Analyze(rejectedInput).Valid);
    }

    [Fact]
    public void Built_in_keys_are_try_added_and_keep_one_metadata_and_execution_selection()
    {
        var replacement = new ReplacementFormatter();
        var services = new ServiceCollection().AddKeyedSingleton<IFormatter>("number", replacement);
        services.AddMonoReport().AddMonoReport();
        using var provider = services.BuildServiceProvider();
        var selected = provider.GetRequiredKeyedService<IFormatter>("number");
        Assert.Same(replacement, selected);
        Assert.Equal("replacement", Assert.Single(selected.Metadata.Parameters).Name);
        foreach (var key in new[] {"upper", "lower", "caps", "firstcap", "date"})
            Assert.NotNull(provider.GetRequiredKeyedService<IFormatter>(key));
        Assert.Null(provider.GetKeyedService<IFormatter>("spellout"));
        Assert.IsType<UpperFormatter>(provider.GetRequiredKeyedService<IFormatter>("upper"));
        Assert.IsType<LowerFormatter>(provider.GetRequiredKeyedService<IFormatter>("lower"));
        Assert.IsType<CapsFormatter>(provider.GetRequiredKeyedService<IFormatter>("caps"));
        Assert.IsType<FirstCapFormatter>(provider.GetRequiredKeyedService<IFormatter>("firstcap"));
        var dateMetadata = provider.GetRequiredKeyedService<IFormatter>("date").Metadata;
        Assert.Equal("pattern", Assert.Single(dateMetadata.Parameters).Name);
        Assert.Equal("pattern", dateMetadata.PositionalParameter);
    }

    [Fact]
    public void Scoped_formatter_dependency_is_shared_within_and_disposed_with_the_caller_scope()
    {
        var disposed = 0;
        var services = new ServiceCollection().AddMonoReport();
        services.AddKeyedScoped<IFormatter>("scoped", (_, _) => new ScopedFormatter(() => disposed++));
        using var provider = services.BuildServiceProvider();
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Value scoped}")));
        using (var scope = provider.CreateScope())
        {
            using var input = new MemoryStream(template);
            Assert.True(scope.ServiceProvider.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
            Assert.Equal("value",
                DocxFixture.Text(DocxFixture.Read(Generate(scope.ServiceProvider, template, """{"Value":"value"}"""))));
            Assert.Equal(0, disposed);
        }

        Assert.Equal(1, disposed);
    }

    [Fact]
    public async Task Concurrent_formatters_keep_values_and_cultures_isolated()
    {
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        var engine = provider.GetRequiredService<DocumentGenerator>();
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Amount number \"N2\"}")));
        var jobs = new[]
        {
            ("1", "en-US", "1.00"), ("1234.5", "de-DE", "1.234,50"),
            ("42.75", "fr-FR", "42,75"), ("0.125", "en-US", "0.13")
        };
        var output = await Task.WhenAll(jobs.Select(job => Task.Run(() =>
        {
            using var input = new MemoryStream(template);
            using var data = JsonDocument.Parse("{\"Amount\":" + job.Item1 + "}");
            var result = engine.Generate(input, data, new GenerationOptions
            {
                Culture = CultureInfo.GetCultureInfo(job.Item2)
            });
            return DocxFixture.Text(DocxFixture.Read(result.Document));
        })));
        Assert.Equal(jobs.Select(job => job.Item3), output);
    }

    private static byte[] Generate(IServiceProvider provider, byte[] template, string json)
    {
        using var input = new MemoryStream(template);
        using var data = JsonDocument.Parse(json);
        return provider.GetRequiredService<DocumentGenerator>().Generate(input, data).Document;
    }

    private sealed class ProbeFormatter : IFormatter
    {
        internal int FormatCalls { get; private set; }

        public FormatterMetadata Metadata { get; } = new(
        [
            new FormatterParameter("first", typeof(string), true),
            new FormatterParameter("suffix", typeof(string), true)
        ], "first");

        public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
        {
            return parameters.TryGetValue("suffix", out var value) && value.Kind == FormatterOperandKind.Literal &&
                   value.Literal == "!"
                ? FormatterValidationResult.Valid
                : new FormatterValidationResult(["suffix must be !"]);
        }

        public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
            FormatterContext context)
        {
            FormatCalls++;
            return parameters["first"].Literal + ":" + parameters["suffix"].Literal + FormatterScalarText.Read(value);
        }
    }

    private sealed class ReplacementFormatter : IFormatter
    {
        public FormatterMetadata Metadata { get; } = new([new FormatterParameter("replacement", typeof(string), true)]);

        public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
        {
            return parameters["replacement"].Kind == FormatterOperandKind.Literal &&
                   parameters["replacement"].Literal == "ok"
                ? FormatterValidationResult.Valid
                : new FormatterValidationResult(["replacement must be ok"]);
        }

        public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
            FormatterContext context)
        {
            return "replaced:" + value.GetRawText();
        }
    }

    private sealed class ScopedFormatter(Action onDispose) : IFormatter, IDisposable
    {
        public void Dispose()
        {
            onDispose();
        }

        public FormatterMetadata Metadata { get; } = new([]);

        public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
        {
            return FormatterValidationResult.Valid;
        }

        public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
            FormatterContext context)
        {
            return value.GetString()!;
        }
    }
}