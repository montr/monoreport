using System;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Text.Json;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.DependencyInjection;
using MonoReport;
using MonoReport.Cli.Tests;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Templates;
using MonoReport.Formatter.Spellout;
using Xunit;

namespace MonoReport.Tests.Docx;

public sealed class DocxSampleJsonGeneratorTests
{
    [Theory]
    [InlineData(SampleJsonValues.Filled, true)]
    [InlineData(SampleJsonValues.Empty, false)]
    public void If_only_boolean_dependencies_produce_boolean_sample_values(SampleJsonValues values, bool expected)
    {
        using var services = Services();
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{If:Enabled}")),
            DocxFixture.P(DocxFixture.Run("conditional text")),
            DocxFixture.P(DocxFixture.Run("{/If}")));
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {Values = values});

        Assert.Equal(expected, sample.Data.GetProperty("Enabled").GetBoolean());
    }

    [Fact]
    public void Generates_nested_object_and_primitive_samples_that_render()
    {
        using var services = Services();
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{For:Orders}")),
            DocxFixture.P(DocxFixture.Run("{Name}")),
            DocxFixture.P(DocxFixture.Run("{For:Tags}")),
            DocxFixture.P(DocxFixture.Run("{.}")),
            DocxFixture.P(DocxFixture.Run("{/For}")),
            DocxFixture.P(DocxFixture.Run("{/For}")));
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(2)});

        var orders = sample.Data.GetProperty("Orders");
        Assert.Equal(2, orders.GetArrayLength());
        Assert.Equal("Sample Name 1", orders[0].GetProperty("Name").GetString());
        Assert.Equal(new[] {"Sample Tags 1", "Sample Tags 2"},
            orders[0].GetProperty("Tags").EnumerateArray().Select(item => item.GetString()));
        Assert.Equal("Sample Name 2", orders[1].GetProperty("Name").GetString());
        Assert.Equal(new[] {"Sample Tags 1", "Sample Tags 2"},
            orders[1].GetProperty("Tags").EnumerateArray().Select(item => item.GetString()));
        using var data = JsonDocument.Parse(sample.Data.GetRawText());
        using var generationInput = new MemoryStream(template);
        var generated = services.GetRequiredService<DocumentGenerator>().Generate(generationInput, data);
        Assert.Contains("Sample Name 2", DocxFixture.Text(DocxFixture.Read(generated.Document)));
    }

    [Fact]
    public void Syntax_showcase_employees_have_inferred_object_shape()
    {
        using var services = Services();
        var template = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "syntax-showcase.template.docx"));
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input);

        var employees = sample.Data.GetProperty("Employees");
        Assert.Equal(10, employees.GetArrayLength());
        Assert.All(employees.EnumerateArray(), employee => Assert.Equal(JsonValueKind.Object, employee.ValueKind));
        Assert.Equal(Enumerable.Range(1, 10).Select(index => $"Sample Name {index}"),
            employees.EnumerateArray().Select(employee => employee.GetProperty("Name").GetString()));
        Assert.Equal(Enumerable.Range(1, 10).Select(index => $"Sample Role {index}"),
            employees.EnumerateArray().Select(employee => employee.GetProperty("Role").GetString()));
        Assert.DoesNotContain(sample.Warnings, warning => warning.Code == DiagnosticCodes.SampleUnresolvedShape);
    }

    [Fact]
    public void Empty_mode_keeps_structure_and_unknown_arrays_remain_empty_with_source_warnings()
    {
        using var services = Services();
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{For:Orders}")),
            DocxFixture.P(DocxFixture.Run("{If:Enabled}")),
            DocxFixture.P(DocxFixture.Run("{Enabled}")),
            DocxFixture.P(DocxFixture.Run("{/If}")),
            DocxFixture.P(DocxFixture.Run("{For:Unknown}")),
            DocxFixture.P(DocxFixture.Run("literal")),
            DocxFixture.P(DocxFixture.Run("{/For}")),
            DocxFixture.P(DocxFixture.Run("{/For}")));
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(2), Values = SampleJsonValues.Empty});

        Assert.Equal(2, sample.Data.GetProperty("Orders").GetArrayLength());
        Assert.False(sample.Data.GetProperty("Orders")[0].GetProperty("Enabled").GetBoolean());
        Assert.Empty(sample.Data.GetProperty("Orders")[0].GetProperty("Unknown").EnumerateArray());
        var warning = Assert.Single(sample.Warnings);
        Assert.Equal("sample.unresolved_shape", warning.Code);
        Assert.Contains("Unknown", warning.Message);
        Assert.NotNull(warning.Location);
    }

    [Fact]
    public void Unknown_empty_arrays_accept_arbitrary_counts_without_iteration()
    {
        using var services = Services();
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{For:Unknown}")),
            DocxFixture.P(DocxFixture.Run("literal")),
            DocxFixture.P(DocxFixture.Run("{/For}")));
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = BigInteger.Pow(10, 100)});

        Assert.Empty(sample.Data.GetProperty("Unknown").EnumerateArray());
        Assert.Single(sample.Warnings);
    }

    [Fact]
    public void Exact_element_budget_succeeds_and_excess_fails_without_a_candidate()
    {
        using var services = Services();
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{For:Items}")),
            DocxFixture.P(DocxFixture.Run("{.}")),
            DocxFixture.P(DocxFixture.Run("{/For}")));
        using var exactInput = new MemoryStream(template);
        var exact = services.GetRequiredService<SampleJsonGenerator>().Generate(exactInput,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(SampleJsonLimits.MaximumArrayElements)});
        Assert.Equal(SampleJsonLimits.MaximumArrayElements, exact.Data.GetProperty("Items").GetArrayLength());

        using var excessiveInput = new MemoryStream(template);
        var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<SampleJsonGenerator>()
            .Generate(excessiveInput, new SampleJsonOptions {ItemsPerCollection = new BigInteger(SampleJsonLimits.PreflightSentinel)}));
        Assert.Equal("sample.element_limit", error.Code);
        Assert.False(error.LimitReached);
        Assert.DoesNotContain(error.Errors, item => item.DataPath is not null);
    }

    [Fact]
    public void Reads_from_the_current_position_and_leaves_the_caller_stream_open()
    {
        using var services = Services();
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Name}")));
        using var input = new MemoryStream(new byte[] {0, 1, 2}.Concat(template).ToArray());
        input.Position = 3;

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input);

        Assert.Equal("Sample Name 1", sample.Data.GetProperty("Name").GetString());
        Assert.True(input.CanRead);
    }

    [Fact]
    public void Reads_non_seekable_streams_without_retaining_operation_state()
    {
        using var services = Services();
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Name}")));
        var generator = services.GetRequiredService<SampleJsonGenerator>();
        using var firstInput = new NonSeekableReadStream(template);
        using var secondInput = new NonSeekableReadStream(template);
        var first = generator.Generate(firstInput);
        var second = generator.Generate(secondInput);

        Assert.Equal(first.Data.GetRawText(), second.Data.GetRawText());
    }

    [Fact]
    public void Uses_root_and_parent_paths_and_known_formatter_values_in_the_generated_sample()
    {
        using var services = Services();
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{For:Orders}")),
            DocxFixture.P(DocxFixture.Run("{Name}/{../Title}/{$.Title}")),
            DocxFixture.P(DocxFixture.Run("{Amount number \"0.00\"}/{When date \"yyyy-MM-dd\"}")),
            DocxFixture.P(DocxFixture.Run("{/For}")));
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(2)});

        Assert.Equal("Sample Title 1", sample.Data.GetProperty("Title").GetString());
        var first = sample.Data.GetProperty("Orders")[0];
        Assert.Equal("Sample Name 1", first.GetProperty("Name").GetString());
        Assert.Equal(42.5m, first.GetProperty("Amount").GetDecimal());
        Assert.Equal("2024-01-15", first.GetProperty("When").GetString());
        using var data = JsonDocument.Parse(sample.Data.GetRawText());
        using var generationInput = new MemoryStream(template);
        var generated = services.GetRequiredService<DocumentGenerator>().Generate(generationInput, data);
        Assert.Contains("42.50", DocxFixture.Text(DocxFixture.Read(generated.Document)));
    }

    [Fact]
    public void Conflicting_hints_use_first_source_evidence_and_structural_boolean_takes_precedence()
    {
        using var services = Services();
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{Value number \"0.00\"}")),
            DocxFixture.P(DocxFixture.Run("{Value date \"yyyy-MM-dd\"}")),
            DocxFixture.P(DocxFixture.Run("{If:Enabled}")),
            DocxFixture.P(DocxFixture.Run("{Enabled number \"0.00\"}")),
            DocxFixture.P(DocxFixture.Run("{/If}")));
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input);

        Assert.Equal(42.5m, sample.Data.GetProperty("Value").GetDecimal());
        Assert.True(sample.Data.GetProperty("Enabled").GetBoolean());
        Assert.Equal(2, sample.Warnings.Count);
        Assert.All(sample.Warnings, warning => Assert.Equal("sample.semantic_conflict", warning.Code));
    }

    [Fact]
    public void Conflicts_in_one_paragraph_keep_one_warning_per_property()
    {
        using var services = Services();
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run(
            "{First number \"0.0\"}/{First date \"yyyy-MM-dd\"} {Second number \"0.0\"}/{Second date \"yyyy-MM-dd\"}")));
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input);

        var warnings = sample.Warnings.Where(warning => warning.Code == DiagnosticCodes.SampleSemanticConflict).ToArray();
        Assert.Equal(2, warnings.Length);
        Assert.Equal(new[] {"$.First", "$.Second"}, warnings.Select(warning => warning.DataPath));
    }

    [Fact]
    public void Conflicts_in_repeated_collection_items_deduplicate_by_property_and_source()
    {
        using var services = Services();
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{For:Items}")),
            DocxFixture.P(DocxFixture.Run(
                "{Value number \"0.0\"}/{Value date \"yyyy-MM-dd\"} " +
                "{Other number \"0.0\"}/{Other date \"yyyy-MM-dd\"}")),
            DocxFixture.P(DocxFixture.Run("{/For}")));
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(3)});

        var warnings = sample.Warnings.Where(warning => warning.Code == DiagnosticCodes.SampleSemanticConflict).ToArray();
        Assert.Equal(2, warnings.Length);
        Assert.Equal(new[] {"$.Items[0].Value", "$.Items[0].Other"}, warnings.Select(warning => warning.DataPath));
        Assert.All(warnings, warning => Assert.NotNull(warning.Location));
        Assert.Contains("Items", warnings[0].DataPath);
        Assert.Contains("Items", warnings[1].DataPath);
    }

    [Fact]
    public void Semantic_conflicts_follow_story_traversal_order()
    {
        using var services = Services();
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run(new string('x', 200))),
            DocxFixture.P(DocxFixture.Run("{Value number \"0.0\"}")));
        using var package = new MemoryStream();
        package.Write(template);
        package.Position = 0;
        using (var document = WordprocessingDocument.Open(package, true))
        {
            var header = document.MainDocumentPart!.AddNewPart<HeaderPart>();
            header.Header = new Header(new Paragraph(new Run(new Text("{Value date \"yyyy-MM-dd\"}"))));
            header.Header.Save();
        }
        package.Position = 0;

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(package);

        Assert.Equal(42.5m, sample.Data.GetProperty("Value").GetDecimal());
    }

    private static ServiceProvider Services()
    {
        return new ServiceCollection().AddMonoReport().AddMonoReportSpelloutFormatter().BuildServiceProvider();
    }

    private sealed class NonSeekableReadStream(byte[] content) : Stream
    {
        private readonly MemoryStream _source = new(content, false);

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => _source.Read(buffer, offset, count);
        public override int Read(Span<byte> buffer) => _source.Read(buffer);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing) _source.Dispose();
            base.Dispose(disposing);
        }
    }
}
