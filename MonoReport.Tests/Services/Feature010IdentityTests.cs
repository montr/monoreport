using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed partial class Feature010IdentityTests
{
    private static readonly XNamespace W = WordXml.WordprocessingNamespace;
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

    [Theory]
    [InlineData("crossing-bookmark", "template.invalid_structure")]
    [InlineData("external-bookmark-link", "template.invalid_structure")]
    [InlineData("nested-bookmark-link", "template.invalid_structure")]
    [InlineData("tracked-cell-deletion", "template.tracked_revision")]
    [InlineData("tracked-move-range", "template.tracked_revision")]
    [InlineData("expression-crosses-content-control-exit", "template.unmatched_brace")]
    public void Persisted_invalid_layouts_reject_before_output_with_diagnostic_parity(string name, string code)
    {
        var template = Fixture(name, ".template.docx");
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, System.Text.Encoding.UTF8.GetString(Fixture(name, ".json")), code);
        if (name == "crossing-bookmark")
        {
            Assert.Contains("bookmark", analysis.Errors[0].Message);
            Assert.Contains("0", analysis.Errors[0].Message);
            Assert.Contains("paragraph 1", analysis.Errors[0].Location);
            Assert.Contains("paragraph 3", analysis.Errors[0].Message);
        }
        if (name == "tracked-cell-deletion") Assert.Contains("table 1, row 1, cell 1", analysis.Errors[0].Location);
        if (name == "tracked-move-range")
        {
            Assert.Equal(2, analysis.Errors.Count);
            Assert.Contains("moveFromRangeStart", analysis.Errors[0].Location);
            Assert.Contains("moveFromRangeEnd", analysis.Errors[1].Location);
        }
        // Rejection is structural, independent of apparently safe runtime cardinality.
        if (name is "crossing-bookmark" or "external-bookmark-link" or "nested-bookmark-link")
            foreach (var json in new[] { "{}", "{\"Items\":null}", "{\"Items\":[]}", "{\"Items\":[{\"Children\":[]}]}", "{\"Items\":[{\"Children\":[{}]}]}" })
                Reject(template, json, code);
    }

    [Fact]
    public void Persisted_local_comment_images_are_rejected_before_copying()
    {
        var template = Fixture("comment-with-image", ".template.docx");
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        Reject(template, System.Text.Encoding.UTF8.GetString(Fixture("comment-with-image", ".json")),
            "template.invalid_structure");
    }

    internal static void AssertCommentImages(byte[] template, byte[] output, int count)
    {
        using var original = WordprocessingDocument.Open(new MemoryStream(template), false);
        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        var sourcePart = original.MainDocumentPart!.WordprocessingCommentsPart!;
        var part = document.MainDocumentPart!.WordprocessingCommentsPart!;
        using var stream = part.GetStream();
        var comments = XDocument.Load(stream);
        var ids = comments.Descendants(Wp + "docPr").Select(e => (string)e.Attribute("id")!).ToArray();
        Assert.Equal(ids.Length, ids.Distinct().Count());
        Assert.True(ids.Length >= count + 1);
        var xml = Docx.Read(output);
        var references = xml.Descendants(W + "commentReference").Select(e => (string)e.Attribute(W + "id")!).ToArray();
        Assert.Equal(count, references.Length);
        Assert.Equal(count, references.Distinct().Count());
        Assert.Equal(references.Order(), xml.Descendants(W + "commentRangeStart").Select(e => (string)e.Attribute(W + "id")!).Order());
        Assert.Equal(references.Order(), xml.Descendants(W + "commentRangeEnd").Select(e => (string)e.Attribute(W + "id")!).Order());
        using var sourceStream = sourcePart.GetStream();
        var sourceComments = XDocument.Load(sourceStream);
        foreach (var reference in references)
        {
            var definition = comments.Root!.Elements().Single(e => (string?)e.Attribute(W + "id") == reference);
            Assert.Equal(Docx.Text(sourceComments), Docx.Text(definition));
        }
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (var blip in comments.Descendants(a + "blip"))
        {
            var id = (string)blip.Attribute(WordXml.RelationshipsNamespace + "embed")!;
            using var expected = sourcePart.GetPartById(id).GetStream();
            using var actual = part.GetPartById(id).GetStream();
            using var eb = new MemoryStream(); using var ab = new MemoryStream();
            expected.CopyTo(eb); actual.CopyTo(ab);
            Assert.Equal(eb.ToArray(), ab.ToArray());
        }
    }

    internal static byte[] Fixture(string name, string extension) => File.ReadAllBytes(
        Path.Combine(AppContext.BaseDirectory, "Fixtures", "review-reproductions", name + extension));

    internal static TemplateAnalysisResult Reject(byte[] template, string json, string code)
    {
        using var input = new MemoryStream(template);
        var analysis = Analyzer().Analyze(input);
        Assert.False(analysis.Valid);
        Assert.Contains(analysis.Errors, e => e.Code == code);
        input.Position = 0;
        using var data = JsonDocument.Parse(json);
        var error = Assert.Throws<TemplateException>(() => Docx.CreateEngine().Merge(input, data));
        Assert.Equal(analysis.Errors.Take(100).Select(e => (e.Code, e.Part, e.Location, e.Message)),
            error.Errors.Select(e => (e.Code, e.Part, e.Location, e.Message)));
        Assert.Equal(analysis.Errors.Count >= 100, error.LimitReached);
        Assert.True(input.CanRead);
        Assert.Equal(template, input.ToArray());
        return analysis;
    }

    internal static TemplateAnalyzer Analyzer() => new(new DocxStoryReader(), new DocxTemplatePlanner(new TemplateLanguageParser()));
}
