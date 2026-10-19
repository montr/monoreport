using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature010IntegrationTests
{
    [Fact]
    public void Dependency_injection_composes_instance_services_without_legacy_formatter_registry()
    {
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        var first = provider.GetRequiredService<MergeEngine>();
        var second = provider.GetRequiredService<MergeEngine>();
        Assert.NotSame(first, second);
        Assert.NotNull(provider.GetRequiredService<TemplateAnalyzer>());
        Assert.Null(Type.GetType("MonoReport.Services.FormatterRegistry, MonoReport"));
        Assert.Null(typeof(MergeOptions).GetProperty("FieldFormatters"));
    }

    [Fact]
    public async Task Reused_engine_keeps_concurrent_operation_state_isolated()
    {
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        var engine = provider.GetRequiredService<MergeEngine>();
        var template = Docx.Create(Docx.P(Docx.Run("{For:Items}")),
            Docx.P(Feature010IdentityTests.Start("permission", 9)), Docx.P(Docx.Run("{Raw}")),
            Docx.P(Docx.Run("{Literal}")), Docx.P(Docx.Run("{/Raw}")),
            Docx.P(Docx.Run("{.}"), Feature010IdentityTests.End("permission", 9)),
            Docx.P(Docx.Run("{/For}")));

        var outputs = await Task.WhenAll(Enumerable.Range(0, 20).Select(index => Task.Run(() =>
        {
            using var input = new MemoryStream(template);
            using var data = JsonDocument.Parse($"{{\"Items\":[\"{index}\",\"x\"]}}");
            var result = engine.Merge(input, data).Document;
            Docx.AssertValid(result);
            var xml = Docx.Read(result);
            var w = MonoReport.Services.WordXml.WordprocessingNamespace;
            var permissions = xml.Descendants(w + "permStart").Select(e => (string)e.Attribute(w + "id")!).ToArray();
            Assert.Equal(2, permissions.Distinct().Count());
            Assert.Equal(permissions, xml.Descendants(w + "permEnd").Select(e => (string)e.Attribute(w + "id")!));
            return Docx.Text(xml);
        })));

        Assert.Equal(Enumerable.Range(0, 20).Select(index => "{Literal}" + index + "{Literal}x"), outputs);
    }

    [Fact]
    public void Analyzer_and_generation_accept_readable_non_seekable_inputs_and_leave_them_open()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Name}")));
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        using var analysisInput = new NonSeekableStream(template);
        Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(analysisInput).Valid);
        Assert.True(analysisInput.CanRead);

        using var generationInput = new NonSeekableStream(template);
        using var data = JsonDocument.Parse("""{"Name":"OK"}""");
        var result = provider.GetRequiredService<MergeEngine>().Merge(generationInput, data);
        Assert.Equal("OK", Docx.Text(Docx.Read(result.Document)));
        Assert.True(generationInput.CanRead);
    }

    [Fact]
    public void Authoritative_sample_analyzes_generates_and_matches_checked_in_output()
    {
        var template = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "sample.template.docx"));
        var json = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "sample.json"));
        var checkedIn = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "sample.docx"));
        Docx.AssertValid(template);
        using var input = new MemoryStream(template);
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        var analysis = provider.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.True(analysis.Valid);
        Assert.Contains(analysis.Requirements, item => item.Name == "Materials" && item.Type == "array");

        var generated = Docx.Merge(template, json, new MergeOptions
        {
            Culture = System.Globalization.CultureInfo.GetCultureInfo("en-GB")
        }).Document;
        Docx.AssertValid(generated);
        Docx.AssertValid(checkedIn);
        Assert.True(System.Xml.Linq.XNode.DeepEquals(Docx.Read(checkedIn), Docx.Read(generated)));
        Assert.DoesNotContain("{For:", Docx.Text(Docx.Read(generated)));
    }

    private sealed class NonSeekableStream(byte[] bytes) : Stream
    {
        private readonly MemoryStream _inner = new(bytes);
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override int Read(Span<byte> buffer) => _inner.Read(buffer);
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}
