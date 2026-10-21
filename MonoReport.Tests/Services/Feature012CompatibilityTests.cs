using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;
using MonoReport.Services;
using Microsoft.Extensions.DependencyInjection;
using System.Text.Json;
using DocumentKind = MonoReport.Models.DocumentFormat;

namespace MonoReport.Tests.Services;

public sealed class Feature012CompatibilityTests
{
    [Theory]
    [InlineData(WordprocessingDocumentType.Document)]
    [InlineData(WordprocessingDocumentType.Template)]
    [InlineData(WordprocessingDocumentType.MacroEnabledDocument)]
    [InlineData(WordprocessingDocumentType.MacroEnabledTemplate)]
    public void WordFamily_Actual_main_part_routes_and_survives(WordprocessingDocumentType kind)
    {
        var source = Docx.Create(Docx.P(Docx.Run("{Name}")));
        using var edited = new MemoryStream();
        edited.Write(source);
        edited.Position = 0;
        string contentType;
        using (var document = WordprocessingDocument.Open(edited, true))
        {
            document.ChangeDocumentType(kind);
            contentType = document.MainDocumentPart!.ContentType;
        }
        source = edited.ToArray();
        using var input = new MemoryStream(source);
        Assert.Equal(DocumentKind.Docx, new PackageFormatDetector().Detect(input));
        Assert.Equal(0, input.Position);
        Assert.True(Feature010IdentityTests.Analyzer().Analyze(input).Valid);
        var result = Docx.Merge(source, "{\"Name\":\"Ada\"}");
        using var output = WordprocessingDocument.Open(new MemoryStream(result.Document), false);
        Assert.Equal(contentType, output.MainDocumentPart!.ContentType);
        Assert.Equal("Ada", output.MainDocumentPart!.Document!.InnerText);
        using var provider = new XlsxFixture().Services();
        using var analysisInput = new MemoryStream(source);
        Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(analysisInput).Valid);
        using var mergeInput = new MemoryStream(source);
        using var data = JsonDocument.Parse("{\"Name\":\"Ada\"}");
        var routed = provider.GetRequiredService<MergeEngine>().Merge(mergeInput, data);
        using var routedDocument = WordprocessingDocument.Open(new MemoryStream(routed.Document), false);
        Assert.Equal(contentType, routedDocument.MainDocumentPart!.ContentType);
        Assert.Equal("Ada", routedDocument.MainDocumentPart.Document!.InnerText);
    }

    [Fact]
    public void WordFamily_Xlsx_is_detected_from_package_without_filename_or_mime()
    {
        using var input = new MemoryStream(new XlsxFixture().Create());
        Assert.Equal(DocumentKind.Xlsx, new PackageFormatDetector().Detect(input));
        Assert.Equal(0, input.Position);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(SpreadsheetDocumentType.Template)]
    [InlineData(SpreadsheetDocumentType.MacroEnabledWorkbook)]
    [InlineData(SpreadsheetDocumentType.MacroEnabledTemplate)]
    public void WordFamily_Unsupported_excel_kinds_reject(SpreadsheetDocumentType kind)
    {
        using var input = new MemoryStream(new XlsxFixture().Create(kind: kind));
        Assert.Equal(DiagnosticCodes.TemplateInvalidPackage,
            Assert.Throws<TemplateException>(() => new PackageFormatDetector().Detect(input)).Code);
        Assert.Equal(0, input.Position);
    }

    [Fact]
    public void WordFamily_Binary_excel_and_malformed_packages_reject()
    {
        var source = new XlsxFixture().Create();
        using var input = new MemoryStream();
        input.Write(source);
        input.Position = 0;
        using (var package = System.IO.Packaging.Package.Open(input, FileMode.Open, FileAccess.ReadWrite))
        {
            var uri = new Uri("/xl/workbook.xml", UriKind.Relative);
            package.DeletePart(uri);
            package.CreatePart(uri, "application/vnd.ms-excel.sheet.binary.macroEnabled.main");
        }
        input.Position = 0;
        Assert.Throws<TemplateException>(() => new PackageFormatDetector().Detect(input));
        using var bad = new MemoryStream([1, 2, 3]);
        Assert.Throws<TemplateException>(() => new PackageFormatDetector().Detect(bad));
    }

    [Fact]
    public void Composition_Registration_preserves_overrides_lifetimes_and_positional_result()
    {
        var resolver = new DataResolver();
        var analyzer = Feature010IdentityTests.Analyzer();
        var engine = Docx.CreateEngine();
        var services = new ServiceCollection().AddSingleton(resolver).AddSingleton(analyzer).AddSingleton(engine);
        services.AddMonoReport().AddMonoReport();
        using var provider = services.BuildServiceProvider();
        Assert.Same(resolver, provider.GetRequiredService<DataResolver>());
        Assert.Same(analyzer, provider.GetRequiredService<TemplateAnalyzer>());
        Assert.Same(engine, provider.GetRequiredService<MergeEngine>());
        using var defaults = new XlsxFixture().Services();
        Assert.NotSame(defaults.GetRequiredService<MergeEngine>(), defaults.GetRequiredService<MergeEngine>());
        Assert.NotSame(defaults.GetRequiredService<TemplateAnalyzer>(), defaults.GetRequiredService<TemplateAnalyzer>());
        var result = new MergeResult([1], []);
        var (bytes, warnings) = result;
        Assert.Same(result.Document, bytes);
        Assert.Same(result.Warnings, warnings);
        Assert.Equal(DocumentKind.Docx, result.Format);
    }

    [Fact]
    public void Composition_Required_dependencies_and_inputs_have_null_guards()
    {
        var reader = new DocxStoryReader();
        var planner = new DocxTemplatePlanner(new TemplateLanguageParser());
        Assert.Throws<ArgumentNullException>(() => new TemplateAnalyzer(null!, planner));
        Assert.Throws<ArgumentNullException>(() => new TemplateAnalyzer(reader, null!));
        Assert.Throws<ArgumentNullException>(() => new MergeEngine(null!, reader, Feature010IdentityTests.Analyzer()));
        var renderer = new DocumentRenderer(new DataResolver(), new ValueFormatter(new RubSpellOutFormatter()));
        Assert.Throws<ArgumentNullException>(() => new MergeEngine(renderer, null!, Feature010IdentityTests.Analyzer()));
        Assert.Throws<ArgumentNullException>(() => new MergeEngine(renderer, reader, null!));
        using var provider = new XlsxFixture().Services();
        var engine = provider.GetRequiredService<MergeEngine>();
        var analyzer = provider.GetRequiredService<TemplateAnalyzer>();
        using var data = JsonDocument.Parse("{}");
        Assert.Throws<ArgumentNullException>(() => engine.Merge(null!, data));
        Assert.Throws<ArgumentNullException>(() => engine.Merge(Stream.Null, null!));
        Assert.Throws<ArgumentNullException>(() => analyzer.Analyze(null!));
        using var closed = new MemoryStream();
        closed.Close();
        Assert.Throws<ArgumentException>(() => engine.Merge(closed, data));
        Assert.Throws<ArgumentException>(() => analyzer.Analyze(closed));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Streams_Current_position_nonseekable_and_caller_ownership(bool nonseekable, bool spreadsheet)
    {
        var bytes = spreadsheet ? new XlsxFixture().Create() : Docx.Create(Docx.P(Docx.Run("{Name}")));
        using var provider = new XlsxFixture().Services();
        using var input = new MemoryStream();
        input.Write([1, 2, 3]); input.Write(bytes); input.Position = 3;
        using var stream = nonseekable ? new ReadOnlyForwardStream(input) : (Stream)new ReadOnlyForwardStream(input, seekable: true);
        Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(stream).Valid);
        Assert.True(stream.CanRead);
        input.Position = 3;
        using var data = JsonDocument.Parse("{\"Name\":\"Ada\"}");
        var result = provider.GetRequiredService<MergeEngine>().Merge(stream, data);
        if (!spreadsheet) Assert.Equal("Ada", Docx.Text(Docx.Read(result.Document)));
        else
        {
            using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
            Assert.Equal("Ada", output.WorkbookPart!.WorksheetParts.Single().Worksheet!.InnerText);
        }
        Assert.True(stream.CanRead);
        Assert.Equal("Ada", data.RootElement.GetProperty("Name").GetString());
        Assert.Equal(bytes, input.ToArray()[3..]);
    }

    private sealed class ReadOnlyForwardStream(Stream inner, bool seekable = false) : Stream
    {
        public override bool CanRead => inner.CanRead;
        public override bool CanSeek => seekable;
        public override bool CanWrite => false;
        public override long Length => seekable ? inner.Length : throw new NotSupportedException();
        public override long Position { get => seekable ? inner.Position : throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => inner.Read(buffer, offset, count);
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
