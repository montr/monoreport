using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Docx;
using MonoReport.Formatter.Spellout;
using Xunit;
using MonoReport.Tests.Docx;

using MonoReport.Cli.Tests;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Linq;

namespace MonoReport.Tests.Integration;

public sealed class DocumentGenerationIntegrationTests
{
    [Fact]
    public void Dependency_injection_composes_instance_services_without_legacy_formatter_registry()
    {
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        var first = provider.GetRequiredService<DocumentGenerator>();
        var second = provider.GetRequiredService<DocumentGenerator>();
        Assert.NotSame(first, second);
        Assert.NotNull(provider.GetRequiredService<TemplateAnalyzer>());
        Assert.Null(Type.GetType("MonoReport.Core.FormatterRegistry, MonoReport"));
        Assert.Null(typeof(GenerationOptions).GetProperty("FieldFormatters"));
    }

    [Fact]
    public async Task Reused_engine_keeps_concurrent_operation_state_isolated()
    {
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        var engine = provider.GetRequiredService<DocumentGenerator>();
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Items}")),
            DocxFixture.P(DocxIdentityTests.Start("permission", 9)), DocxFixture.P(DocxFixture.Run("{Raw}")),
            DocxFixture.P(DocxFixture.Run("{Literal}")), DocxFixture.P(DocxFixture.Run("{/Raw}")),
            DocxFixture.P(DocxFixture.Run("{.}"), DocxIdentityTests.End("permission", 9)),
            DocxFixture.P(DocxFixture.Run("{/For}")));

        var outputs = await Task.WhenAll(Enumerable.Range(0, 20).Select(index => Task.Run(() =>
        {
            using var input = new MemoryStream(template);
            using var data = JsonDocument.Parse($"{{\"Items\":[\"{index}\",\"x\"]}}");
            var result = engine.Generate(input, data).Document;
            DocxFixture.AssertValid(result);
            var xml = DocxFixture.Read(result);
            var w = WordXml.WordprocessingNamespace;
            var permissions = xml.Descendants(w + "permStart").Select(e => (string) e.Attribute(w + "id")!).ToArray();
            Assert.Equal(2, permissions.Distinct().Count());
            Assert.Equal(permissions, xml.Descendants(w + "permEnd").Select(e => (string) e.Attribute(w + "id")!));
            return DocxFixture.Text(xml);
        })));

        Assert.Equal(Enumerable.Range(0, 20).Select(index => "{Literal}" + index + "{Literal}x"), outputs);
    }

    [Fact]
    public void Analyzer_and_generation_accept_readable_non_seekable_inputs_and_leave_them_open()
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Name}")));
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        using var analysisInput = new NonSeekableStream(template);
        Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(analysisInput).Valid);
        Assert.True(analysisInput.CanRead);

        using var generationInput = new NonSeekableStream(template);
        using var data = JsonDocument.Parse("""{"Name":"OK"}""");
        var result = provider.GetRequiredService<DocumentGenerator>().Generate(generationInput, data);
        Assert.Equal("OK", DocxFixture.Text(DocxFixture.Read(result.Document)));
        Assert.True(generationInput.CanRead);
    }

    [Fact]
    public void Authoritative_sample_analyzes_generates_and_matches_checked_in_output()
    {
        var template = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "sample.template.docx"));
        var json = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "sample.json"));
        var checkedIn = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "sample.docx"));
        DocxFixture.AssertValid(template);
        using var input = new MemoryStream(template);
        using var provider = new ServiceCollection().AddMonoReport().AddMonoReportSpelloutFormatter()
            .BuildServiceProvider();
        var analysis = provider.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.True(analysis.Valid);
        Assert.Contains(analysis.Requirements, item => item.Name == "Materials" && item.Type == "array");

        var generated = DocxFixture.Generate(template, json, new GenerationOptions
        {
            Culture = CultureInfo.GetCultureInfo("en-GB")
        }).Document;
        DocxFixture.AssertValid(generated);
        DocxFixture.AssertValid(checkedIn);
        Assert.True(XNode.DeepEquals(DocxFixture.Read(checkedIn), DocxFixture.Read(generated)));
        Assert.DoesNotContain("{For:", DocxFixture.Text(DocxFixture.Read(generated)));
    }

    private sealed class NonSeekableStream(byte[] bytes) : Stream
    {
        private readonly MemoryStream _inner = new(bytes);
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return _inner.Read(buffer, offset, count);
        }

        public override int Read(Span<byte> buffer)
        {
            return _inner.Read(buffer);
        }

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}