using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed partial class Feature010IdentityTests
{
    [Theory]
    [InlineData("00", false)] [InlineData("00", true)]
    [InlineData("007", false)] [InlineData("007", true)]
    [InlineData("-02", false)] [InlineData("-02", true)]
    public void Noncanonical_local_comment_ids_are_rejected_before_copying(string id, bool point)
    {
        var template = Prepare(VerificationFixture("zero-padded-comment-id"), document =>
        {
            var main = document.MainDocumentPart!;
            var xml = Load(main);
            var attachment = xml.Descendants(W + "body").Single().Elements(W + "p").ElementAt(1);
            foreach (var element in attachment.Descendants().Where(e => e.Attribute(W + "id") is not null))
                element.SetAttributeValue(W + "id", id);
            if (point)
            {
                attachment.Descendants(W + "commentRangeStart").Remove();
                attachment.Descendants(W + "commentRangeEnd").Remove();
            }
            xml.Descendants(W + "body").Single().ReplaceNodes(Docx.P(Docx.Run("{For:Groups}")),
                Docx.P(Docx.Run("{If:Show}")), Docx.P(Docx.Run("{For:Items}")), new XElement(attachment),
                Docx.P(Docx.Run("{/For}")), Docx.P(Docx.Run("{/If}")), Docx.P(Docx.Run("{/For}")));
            Save(main, xml.Root!);
            var comments = Load(main.WordprocessingCommentsPart!);
            comments.Root!.Elements().Single().SetAttributeValue(W + "id", id);
            Save(main.WordprocessingCommentsPart!, comments.Root);
        });
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        foreach (var items in new[] { null, "null", "[]", "[\"A\"]", "[\"A\",\"B\"]" })
        {
            var group = "{\"Show\":true" + (items is null ? "" : ",\"Items\":" + items) + "}";
            var json = "{\"Groups\":[" + group + ", {\"Show\":false,\"Items\":[\"NO\"]}," + group + "]}";
            Reject(template, json, "template.invalid_structure");
        }
    }

    [Fact]
    public void Missing_comment_definition_rejects_in_shared_preflight_even_for_zero_iterations()
    {
        // Deliberately malformed source: rejection must precede any clone or output.
        var template = Prepare(VerificationFixture("zero-padded-comment-id"), document =>
            document.MainDocumentPart!.DeletePart(document.MainDocumentPart.WordprocessingCommentsPart!));
        foreach (var json in new[] { "{}", "{\"Items\":[]}", "{\"Items\":[\"A\"]}" })
            Reject(template, json, "template.invalid_structure");
    }

    [Theory]
    [InlineData("external-relationship-with-anchor")]
    [InlineData("cloned-external-relationship-with-anchor")]
    public void External_hyperlink_fragments_do_not_make_local_bookmarks_supported(string name)
    {
        var template = VerificationFixture(name);
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        foreach (var json in new[] { "{}", "{\"Items\":[]}", "{\"Items\":[\"A\"]}", "{\"Items\":[\"A\",\"B\"]}" })
            Reject(template, json, "template.invalid_structure");
    }

    [Fact]
    public void External_link_does_not_make_a_matching_local_bookmark_supported()
    {
        var template = Prepare(VerificationFixture("external-relationship-with-anchor"), document =>
        {
            var main = document.MainDocumentPart!;
            var xml = Load(main);
            foreach (var text in xml.Descendants(W + "t"))
                text.Value = text.Value switch { "{For:Items}" => "{If:Show}", "{/For}" => "{/If}", "{.}" => "local", _ => text.Value };
            Save(main, xml.Root!);
        });
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        foreach (var show in new[] { false, true })
            Reject(template, JsonSerializer.Serialize(new { Show = show }), "template.invalid_structure");
    }

    [Fact]
    public void External_fragments_in_comment_definitions_do_not_make_local_comment_copies_supported()
    {
        var template = Prepare(VerificationFixture("zero-padded-comment-id"), document =>
        {
            var part = document.MainDocumentPart!.WordprocessingCommentsPart!;
            var relation = part.AddHyperlinkRelationship(new Uri("https://example.org/report.docx"), true);
            var comments = Load(part);
            var link = Link("B4", "external {Hidden}");
            link.SetAttributeValue(WordXml.RelationshipsNamespace + "id", relation.Id);
            comments.Root!.Elements().Single().Add(Docx.P(Start("bookmark", 4), Docx.Run("target"), End("bookmark", 4), link, Link("B4", "internal")));
            Save(part, comments.Root);
        });
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        Reject(template, "{\"Items\":[\"A\",\"B\"]}", "template.invalid_structure");
    }

    [Theory]
    [InlineData("inline", true)] [InlineData("inline", false)]
    [InlineData("block", true)] [InlineData("block", false)]
    [InlineData("row", true)] [InlineData("row", false)]
    [InlineData("cell", true)] [InlineData("cell", false)]
    public void Complex_field_state_crosses_opaque_controls_without_scanning_hidden_expressions(string placement, bool beginsInside)
    {
        var begin = new[] { new XElement(W + "r", Docx.Character("begin")),
            new XElement(W + "r", new XElement(W + "instrText", "MERGEFIELD Opaque")),
            new XElement(W + "r", Docx.Character("separate")) };
        var end = new XElement(W + "r", Docx.Character("end"));
        var hidden = beginsInside ? new[] { Control(begin), Control(Docx.Run("{Hidden}")), Docx.Run("{Hidden}") }
            : new[] { Control(end), Control(Docx.Run("{Hidden}")) };
        var control = placement switch
        {
            "inline" => Control(hidden),
            "block" => Control(Docx.P(hidden)),
            "row" => Control(Docx.Row(hidden)),
            _ => Control(new XElement(W + "tc", Docx.P(hidden)))
        };
        var wrapped = placement switch
        {
            "inline" => Docx.P(control),
            "row" => Docx.Table(control),
            "cell" => Docx.Table(new XElement(W + "tr", control)),
            _ => control
        };
        var cache = Docx.Run("{Name}");
        var template = beginsInside
            ? Docx.Create(wrapped, Docx.P(cache, end, Docx.Run("{Visible}")))
            : Docx.Create(Docx.P(begin, cache), wrapped, Docx.P(Docx.Run("{Visible}")));
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = AnalyzeValid(template);
        Assert.Equal(["{Visible}"], analysis.Occurrences.Select(o => o.Expression));
        Assert.Equal("Visible", Assert.Single(analysis.Requirements).Name);
        var output = Docx.Merge(template, "{\"Name\":\"CORRUPTED\",\"Visible\":\"OK\"}", validationVersion: FileFormatVersions.Office2010).Document;
        var xml = Docx.Read(output);
        Assert.Equal(Docx.Text(Docx.Read(template)).Replace("{Visible}", "OK", StringComparison.Ordinal), Docx.Text(xml));
        Assert.True(XNode.DeepEquals(control, xml.Descendants(W + "sdt").First()));
    }

    [Theory]
    [InlineData("field-begins-in-control", "{Name}OK")]
    [InlineData("field-ends-in-control", "cachedOK")]
    public void Persisted_inline_field_boundaries_keep_caches_literal_and_render_following_text(string name, string expected)
    {
        var template = VerificationFixture(name);
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        Assert.Equal(["{Visible}"], AnalyzeValid(template).Occurrences.Select(o => o.Expression));
        var output = Docx.Merge(template, "{\"Name\":\"CORRUPTED\",\"Visible\":\"OK\"}", validationVersion: FileFormatVersions.Office2010).Document;
        Assert.Equal(expected, Docx.Text(Docx.Read(output)));
        Assert.True(XNode.DeepEquals(Docx.Read(template).Descendants(W + "sdt").Single(), Docx.Read(output).Descendants(W + "sdt").Single()));
    }

    private static byte[] VerificationFixture(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "review-verification", name + ".template.docx"));
    private static XDocument Load(OpenXmlPart part) { using var stream = part.GetStream(); return XDocument.Load(stream); }
    private static MonoReport.Models.TemplateAnalysisResult AnalyzeValid(byte[] template)
    {
        using var input = new MemoryStream(template);
        var analysis = Analyzer().Analyze(input);
        Assert.True(analysis.Valid, string.Join("; ", analysis.Errors.Select(e => e.Message)));
        return analysis;
    }
}
