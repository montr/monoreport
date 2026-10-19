using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature011CompatibilityTests
{
    private static readonly XNamespace W = WordXml.WordprocessingNamespace;
    private static XElement P(string text) => Docx.P(Docx.Run(text));
    private static XElement Start(string family, int id) => Feature010IdentityTests.Start(family, id);
    private static XElement End(string family, int id) => Feature010IdentityTests.End(family, id);

    internal static byte[] Comments(byte[] source, params string[] ids) => Feature010IdentityTests.Prepare(source, document =>
    {
        var main = document.MainDocumentPart!;
        using var stream = main.GetStream(); var xml = XDocument.Load(stream); stream.Close();
        foreach (var holder in xml.Descendants("holder").ToArray()) holder.ReplaceWith(holder.Elements().ToArray());
        Feature010IdentityTests.Save(main, xml.Root!);
        Feature010IdentityTests.Save(main.AddNewPart<WordprocessingCommentsPart>(), new XElement(W + "comments",
            ids.Select(id => new XElement(W + "comment", new XAttribute(W + "id", id), new XAttribute(W + "author", "Author"),
                P("definition " + id)))));
    });

    [Theory]
    [InlineData("bookmark")]
    [InlineData("comment")]
    public void Local_and_marker_annotations_reject_even_in_inactive_or_zero_iteration_operations(string family)
    {
        foreach (var marker in new[] { false, true })
        {
            var source = Comments(Docx.Create(marker ? Docx.P(Start(family, 1), Docx.Run("{For:Items}"), End(family, 1)) : P("{For:Items}"),
                P("{Raw}"), marker ? P("literal") : Docx.P(Start(family, 1), Docx.Run("literal"), End(family, 1)),
                P("{/Raw}"), P("{/For}")), "1");
            foreach (var data in new[] { "{}", "{\"Items\":[]}", "{\"Items\":[1]}", "{\"Items\":[1,2]}" })
                Feature010IdentityTests.Reject(source, data, "template.invalid_structure");
        }
    }

    [Theory]
    [InlineData("bookmark")]
    [InlineData("comment")]
    public void Complete_enclosing_ranges_are_preserved_with_unchanged_identities(string family)
    {
        var source = Comments(Docx.Create(Docx.P(Start(family, 4), Docx.Run("before")),
            P("{For:Items}"), P("{.}"), P("{/For}"), Docx.P(Docx.Run("after"), End(family, 4))), "4");
        using var input = new MemoryStream(source);
        Assert.True(Feature010IdentityTests.Analyzer().Analyze(input).Valid);
        var output = Docx.Merge(source, "{\"Items\":[\"A\",\"B\"]}", validationVersion: FileFormatVersions.Office2010).Document;
        Assert.Equal("beforeABafter", Docx.Text(Docx.Read(output)));
        var name = family == "bookmark" ? "bookmarkStart" : "commentRangeStart";
        var original = Docx.Read(source).Descendants(W + name).Single();
        Assert.True(XNode.DeepEquals(original, Docx.Read(output).Descendants(W + name).Single()));
    }

    [Theory]
    [InlineData("Stable")]
    [InlineData("Missing")]
    public void Internal_links_inside_operations_reject_even_with_stable_or_missing_targets(string target)
    {
        var source = Docx.Create(Docx.P(Feature010IdentityTests.Bookmark(0, "Stable", "stable")),
            P("{If:Show}"), Docx.P(Feature010IdentityTests.Link(target, "link")), P("{/If}"));
        foreach (var data in new[] { "{\"Show\":false}", "{\"Show\":true}" })
            Feature010IdentityTests.Reject(source, data, "template.invalid_structure");
    }

    [Theory]
    [InlineData("bookmark")]
    [InlineData("comment")]
    [InlineData("permission")]
    public void Raw_cannot_split_deleted_and_surviving_range_members(string family)
    {
        var source = Comments(Docx.Create(Docx.P(Start(family, 2), Docx.Run("{Raw}")),
            Docx.P(Docx.Run("body"), End(family, 2)), P("{/Raw}")), "2");
        Feature010IdentityTests.Reject(source, "{}", "template.invalid_structure");
    }

    [Fact]
    public void First_eligible_close_with_unsafe_deletion_is_not_skipped()
    {
        var source = Comments(Docx.Create(P("{Raw}"), Docx.P(Start("bookmark", 0), Docx.Run("body")),
            Docx.P(End("bookmark", 0), Docx.Run("{/Raw}")), P("{/Raw}")), "0");
        var result = Feature010IdentityTests.Reject(source, "{}", "template.invalid_structure");
        Assert.Contains(result.Errors, error => error.Message.Contains("bookmark"));
        Assert.Contains(result.Errors, error => error.Field == "{/Raw}");
    }

    [Fact]
    public void Raw_section_boundary_deletion_and_missing_cell_content_reject()
    {
        var section = new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "sectPr")), Docx.Run("{Raw}"));
        Feature010IdentityTests.Reject(Docx.Create(section, P("body"), P("{/Raw}")), "{}", "template.invalid_structure");
        var cell = new XElement(W + "tc", P("{Raw}"), P("{/Raw}"));
        Feature010IdentityTests.Reject(Docx.Create(Docx.Table(new XElement(W + "tr", cell))), "{}", "template.invalid_structure");
    }

    [Fact]
    public void Recognized_revisions_remain_errors_in_Raw_and_inactive_branches()
    {
        foreach (var wrapping in new[] { "Raw", "If:Show", "For:Items" })
        {
            var close = wrapping == "Raw" ? "{/Raw}" : wrapping.StartsWith("If") ? "{/If}" : "{/For}";
            var revision = new XElement(W + "ins", new XAttribute(W + "id", 0), new XAttribute(W + "author", "A"), Docx.Run("{Bad syntax}"));
            Feature010IdentityTests.Reject(Docx.Create(P("{" + wrapping + "}"), Docx.P(revision), P(close)),
                "{\"Show\":false,\"Items\":[]}", "template.tracked_revision");
        }
    }

    [Theory]
    [InlineData("bookmark", false)]
    [InlineData("comment", false)]
    [InlineData("comment", true)]
    public void Raw_delimiter_annotations_are_deleted_without_dangling_survivors(string family, bool point)
    {
        var annotation = point ? new object[] { new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", "00"))) } :
            [Start(family, 0), End(family, 0)];
        var source = Comments(Docx.Create(Docx.P(annotation, Docx.Run("{Raw}")),
            Docx.P(Start("bookmark", 4), Start("comment", 5), Docx.Run("{Literal}"), End("bookmark", 4), End("comment", 5)),
            P("{/Raw}"), Docx.P(Feature010IdentityTests.Link("B4", "link"))), "0", "5", "99");
        using var input = new MemoryStream(source);
        Assert.True(Feature010IdentityTests.Analyzer().Analyze(input).Valid);
        using var original = WordprocessingDocument.Open(new MemoryStream(source), false);
        var originalPart = original.MainDocumentPart!.WordprocessingCommentsPart!;
        var originalRelationshipId = original.MainDocumentPart.GetIdOfPart(originalPart);
        using var originalStream = originalPart.GetStream();
        var retainedDefinitions = XDocument.Load(originalStream).Root!.Elements()
            .Where(comment => (string?)comment.Attribute(W + "id") is "5" or "99")
            .Select(comment => comment.ToString(SaveOptions.DisableFormatting)).ToArray();
        var output = Docx.Merge(source, "{}", validationVersion: FileFormatVersions.Office2010).Document;
        if (family == "comment" && !point)
        {
            Directory.CreateDirectory("/private/tmp/feature011-evidence/artifacts");
            File.WriteAllBytes("/private/tmp/feature011-evidence/artifacts/raw-annotation-cleanup.docx", output);
        }
        var xml = Docx.Read(output);
        Assert.Equal("{Literal}link", Docx.Text(xml));
        Assert.Single(xml.Descendants(W + "bookmarkStart"));
        Assert.Single(xml.Descendants(W + "commentReference"));
        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        var outputPart = document.MainDocumentPart!.WordprocessingCommentsPart!;
        Assert.Equal(originalRelationshipId, document.MainDocumentPart.GetIdOfPart(outputPart));
        using var stream = outputPart.GetStream();
        var outputComments = XDocument.Load(stream);
        Assert.Equal(retainedDefinitions, outputComments.Root!.Elements()
            .Where(comment => (string?)comment.Attribute(W + "id") is "5" or "99")
            .Select(comment => comment.ToString(SaveOptions.DisableFormatting)));
        var ids = outputComments.Root!.Elements().Select(e => (string?)e.Attribute(W + "id")).ToArray();
        Assert.Equal(family == "comment" ? new[] { "5", "99" } : new[] { "0", "5", "99" }, ids);
    }

    [Fact]
    public void Raw_renders_independently_in_supported_story_and_note_roots()
    {
        var source = Feature010IdentityTests.Prepare(Docx.Create(P("main")), document =>
        {
            var main = document.MainDocumentPart!;
            XElement Content(XName root, string literal) => new(root, P("{Raw}"), P(literal + " {Hidden}"),
                P("{/Raw}"), P("{Name}"));
            Feature010IdentityTests.Save(main.AddNewPart<HeaderPart>(), Content(W + "hdr", "header"));
            Feature010IdentityTests.Save(main.AddNewPart<FooterPart>(), Content(W + "ftr", "footer"));
            Feature010IdentityTests.Save(main.AddNewPart<FootnotesPart>(), new XElement(W + "footnotes",
                new XElement(W + "footnote", new XAttribute(W + "id", 1), Content(W + "customXml", "footnote").Elements())));
            Feature010IdentityTests.Save(main.AddNewPart<EndnotesPart>(), new XElement(W + "endnotes",
                new XElement(W + "endnote", new XAttribute(W + "id", 1), Content(W + "customXml", "endnote").Elements())));
        });
        using var input = new MemoryStream(source);
        var analysis = Feature010IdentityTests.Analyzer().Analyze(input);
        Assert.True(analysis.Valid, string.Join("; ", analysis.Errors.Select(error => error.Message)));
        Assert.Equal("Name", Assert.Single(analysis.Requirements).Name);
        Assert.DoesNotContain(analysis.Occurrences, occurrence => occurrence.Expression == "{Hidden}");
        var output = Docx.Merge(source, "{\"Name\":\"OK\"}", validationVersion: FileFormatVersions.Office2010).Document;
        var stories = Docx.ReadStories(output);
        Assert.Contains(stories.Values, story => Docx.Text(story).Contains("header {Hidden}OK", StringComparison.Ordinal));
        Assert.Contains(stories.Values, story => Docx.Text(story).Contains("footer {Hidden}OK", StringComparison.Ordinal));
        Assert.Contains(stories.Values, story => Docx.Text(story).Contains("footnote {Hidden}OK", StringComparison.Ordinal));
        Assert.Contains(stories.Values, story => Docx.Text(story).Contains("endnote {Hidden}OK", StringComparison.Ordinal));
        Assert.All(stories.Values, story => Assert.DoesNotContain("{Raw}", Docx.Text(story), StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("bookmark")]
    [InlineData("comment")]
    public void Complete_ranges_spanning_standalone_Raw_are_preserved_unchanged(string family)
    {
        var source = Comments(Docx.Create(Docx.P(Start(family, 4), Docx.Run("before")), P("{Raw}"),
            P("literal {Hidden}"), P("{/Raw}"), Docx.P(Docx.Run("after"), End(family, 4)),
            Docx.P(Feature010IdentityTests.Link("B4", "link"))), "4");
        var before = Docx.Read(source);
        var output = Docx.Read(Docx.Merge(source, "{}", validationVersion: FileFormatVersions.Office2010).Document);
        Assert.Equal("beforeliteral {Hidden}afterlink", Docx.Text(output));
        var startName = family == "bookmark" ? "bookmarkStart" : "commentRangeStart";
        var endName = family == "bookmark" ? "bookmarkEnd" : "commentRangeEnd";
        Assert.True(XNode.DeepEquals(before.Descendants(W + startName).Single(), output.Descendants(W + startName).Single()));
        Assert.True(XNode.DeepEquals(before.Descendants(W + endName).Single(), output.Descendants(W + endName).Single()));
        if (family == "bookmark")
            Assert.True(XNode.DeepEquals(before.Descendants(W + "hyperlink").Single(), output.Descendants(W + "hyperlink").Single()));
    }

    [Fact]
    public void Multiple_Raw_deletions_are_cumulative_and_preserve_unrelated_comment_definitions()
    {
        XElement Point(int id) => new(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", id)));
        var source = Comments(Docx.Create(Docx.P(Point(1), Docx.Run("{Raw}")), P("one"), P("{/Raw}"),
            P("between"), Docx.P(Point(2), Docx.Run("{Raw}")), P("two"), P("{/Raw}")), "1", "2", "9");
        var output = Docx.Merge(source, "{}", validationVersion: FileFormatVersions.Office2010).Document;
        Assert.Equal("onebetweentwo", Docx.Text(Docx.Read(output)));
        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        using var stream = document.MainDocumentPart!.WordprocessingCommentsPart!.GetStream();
        Assert.Equal("9", (string?)Assert.Single(XDocument.Load(stream).Root!.Elements()).Attribute(W + "id"));
    }

    [Fact]
    public void Nested_table_requires_a_surviving_final_cell_paragraph_after_Raw_delimiters()
    {
        var nested = Docx.Table(Docx.Row(Docx.Run("nested")));
        var invalid = Docx.Create(Docx.Table(new XElement(W + "tr", new XElement(W + "tc",
            P("{Raw}"), nested, P("{/Raw}")))));
        Feature010IdentityTests.Reject(invalid, "{}", "template.invalid_structure");
        var valid = Docx.Create(Docx.Table(new XElement(W + "tr", new XElement(W + "tc",
            P("{Raw}"), new XElement(nested), P("{/Raw}"), P("tail")))));
        Assert.Equal("nestedtail", Docx.Text(Docx.Read(Docx.Merge(valid, "{}").Document)));
    }

    [Fact]
    public void Stable_and_unresolved_internal_links_and_external_fragments_survive_standalone_Raw()
    {
        var source = Docx.Create(P("{Raw}"), Docx.P(Feature010IdentityTests.Bookmark(3, "Target", "target"),
            Feature010IdentityTests.Link("Target", "stable")), P("{/Raw}"),
            Docx.P(Feature010IdentityTests.Link("Missing", "unresolved")));
        source = Feature010IdentityTests.Prepare(source, document =>
        {
            var main = document.MainDocumentPart!;
            var relationship = main.AddHyperlinkRelationship(new Uri("https://example.org/report.docx"), true);
            using var stream = main.GetStream();
            var xml = XDocument.Load(stream);
            stream.Close();
            xml.Descendants(W + "p").Last().Add(new XElement(W + "hyperlink",
                new XAttribute(WordXml.RelationshipsNamespace + "id", relationship.Id),
                new XAttribute(W + "anchor", "Fragment"), Docx.Run("external")));
            Feature010IdentityTests.Save(main, xml.Root!);
        });
        var output = Docx.Merge(source, "{}", validationVersion: FileFormatVersions.Office2010).Document;
        var xml = Docx.Read(output);
        Assert.Equal(new[] { "Target", "Missing", "Fragment" },
            xml.Descendants(W + "hyperlink").Select(link => (string?)link.Attribute(W + "anchor")));
        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        var relationship = Assert.Single(document.MainDocumentPart!.HyperlinkRelationships);
        Assert.Equal("https://example.org/report.docx", relationship.Uri.ToString());
        Assert.True(relationship.IsExternal);
    }

    [Fact]
    public void Shared_definition_survival_counts_references_in_another_root_even_in_a_malformed_package()
    {
        XElement Reference() => new(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 0)));
        var source = Comments(Docx.Create(Docx.P(Reference(), Docx.Run("{Raw}")), P("literal"), P("{/Raw}")), "0");
        source = Feature010IdentityTests.Prepare(source, document =>
        {
            var header = document.MainDocumentPart!.AddNewPart<HeaderPart>();
            Feature010IdentityTests.Save(header, new XElement(W + "hdr", Docx.P(Reference(), Docx.Run("unchanged"))));
        });
        // Supplemental semantic probe: the header reference cannot resolve the main document's comments part.
        // Conformant sources instead have unique attachment IDs; ordinary definition cleanup is validated separately.
        using var input = new MemoryStream(source);
        using var data = System.Text.Json.JsonDocument.Parse("{}");
        Assert.True(Feature010IdentityTests.Analyzer().Analyze(input).Valid);
        input.Position = 0;
        var output = Docx.CreateEngine().Merge(input, data).Document;
        using var doc = WordprocessingDocument.Open(new MemoryStream(output), false);
        Assert.Single(doc.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements());
        Assert.Empty(Docx.Read(output).Descendants(W + "commentReference"));
        Assert.Single(Docx.ReadStories(output).Values.SelectMany(xml => xml.Descendants(W + "commentReference")));
    }

    [Fact]
    public void Retained_auxiliary_link_to_removed_Raw_bookmark_rejects_without_remapping()
    {
        var source = Docx.Create(Docx.P(Start("bookmark", 0), Docx.Run("{Raw}"), End("bookmark", 0)), P("literal"), P("{/Raw}"));
        source = Feature010IdentityTests.Prepare(source, document => Feature010IdentityTests.Save(
            document.MainDocumentPart!.AddNewPart<HeaderPart>(), new XElement(W + "hdr", Docx.P(Feature010IdentityTests.Link("B0", "outside")))));
        var result = Feature010IdentityTests.Reject(source, "{}", "template.invalid_structure");
        Assert.Contains(result.Errors, error => error.Part!.Contains("header"));
    }

    [Fact]
    public void Raw_row_deletion_preserves_body_merges_but_rejects_a_deleted_required_restart()
    {
        XElement Cell(string text, string? merge = null) => new(W + "tc",
            merge is null ? null : new XElement(W + "tcPr", new XElement(W + "vMerge", new XAttribute(W + "val", merge))), P(text));
        XElement Row(string text, string? merge = null) => new(W + "tr", Cell(text, merge));
        var merged = new[] { Row("{Literal}", "restart"), Row("{{", "continue") };
        var source = Docx.Create(Docx.Table(new[] { Row("{Raw}") }.Concat(merged).Append(Row("{/Raw}")).ToArray()));
        var output = Docx.Read(Docx.Merge(source, "{}").Document);
        Assert.Equal(2, output.Descendants(W + "vMerge").Count());
        Assert.True(merged.Zip(output.Descendants(W + "tr")).All(pair => XNode.DeepEquals(pair.First, pair.Second)));
        var unsafeSource = Docx.Create(Docx.Table(Row("{Raw}", "restart"), Row("literal", "continue"), Row("{/Raw}")));
        Feature010IdentityTests.Reject(unsafeSource, "{}", "template.invalid_structure");
        var harmless = Docx.Create(Docx.Table(Row("before", "restart"), Row("{Raw}", "continue"),
            Row("literal", "continue"), Row("{/Raw}")));
        Assert.Equal("beforeliteral", Docx.Text(Docx.Read(Docx.Merge(harmless, "{}").Document)));
    }

    internal static byte[] SupportedCopyTemplate()
    {
        var source = Feature010IdentityTests.Fixture("comment-with-image", ".template.docx");
        return Feature010IdentityTests.Prepare(source, document =>
        {
            var main = document.MainDocumentPart!;
            var commentsPart = main.WordprocessingCommentsPart!;
            using var input = commentsPart.GetStream(); var comments = XDocument.Load(input); input.Close();
            var drawing = new XElement(comments.Descendants(W + "drawing").Single());
            var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
            var image = commentsPart.GetPartById((string)drawing.Descendants(a + "blip").Single().Attribute(WordXml.RelationshipsNamespace + "embed")!);
            main.AddPart(image);
            drawing.Descendants(a + "blip").Single().SetAttributeValue(WordXml.RelationshipsNamespace + "embed", main.GetIdOfPart(image));
            drawing.Descendants(XName.Get("docPr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")).Single().SetAttributeValue("id", 8);
            comments.Root!.Add(new XElement(W + "comment", new XAttribute(W + "id", 7), new XAttribute(W + "author", "A"), P("delete")));
            Feature010IdentityTests.Save(commentsPart, comments.Root);
            var glossary = main.AddNewPart<GlossaryDocumentPart>();
            glossary.AddPart(image);
            var reserved = new XElement(drawing);
            reserved.Descendants(a + "blip").Single().SetAttributeValue(WordXml.RelationshipsNamespace + "embed", glossary.GetIdOfPart(image));
            reserved.Descendants(XName.Get("docPr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")).Single().SetAttributeValue("id", 3);
            Feature010IdentityTests.Save(glossary, new XElement(W + "glossaryDocument", new XElement(W + "docParts",
                new XElement(W + "docPart", new XElement(W + "docPartPr", new XElement(W + "name", new XAttribute(W + "val", "entry"))),
                    new XElement(W + "docPartBody", Docx.P(new XElement(W + "r", reserved)))))));
            Feature010IdentityTests.Save(main, new XElement(W + "document", new XElement(W + "body",
                Docx.P(new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 7))), Docx.Run("{Raw}")),
                P("{Malformed"), P("{/Raw}"),
                Docx.P(new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 0)))),
                P("{For:Items}"), Docx.P(Start("permission", 9)), P("{Raw}"),
                Docx.P(new XElement(W + "r", drawing)), P("{{ {Literal}"), P("{/Raw}"),
                Docx.P(Docx.Run("{.}"), End("permission", 9)), P("{/For}"))));
        });
    }

    [Fact]
    public void Repeated_Raw_drawings_and_permissions_keep_package_identity_and_media_safety()
    {
        var source = SupportedCopyTemplate();
        var output = Docx.Merge(source, "{\"Items\":[\"A\",\"B\"]}", validationVersion: FileFormatVersions.Office2010).Document;
        Directory.CreateDirectory("/private/tmp/feature011-evidence/artifacts");
        File.WriteAllBytes("/private/tmp/feature011-evidence/artifacts/repeated-raw-drawings-permissions.docx", output);
        var xml = Docx.Read(output);
        Assert.Equal("{Malformed{{ {Literal}A{{ {Literal}B", Docx.Text(xml));
        var starts = xml.Descendants(W + "permStart").Select(e => (string?)e.Attribute(W + "id")).ToArray();
        Assert.Equal(2, starts.Distinct().Count());
        Assert.Equal(starts, xml.Descendants(W + "permEnd").Select(e => (string?)e.Attribute(W + "id")));
        using var original = WordprocessingDocument.Open(new MemoryStream(source), false);
        using var generated = WordprocessingDocument.Open(new MemoryStream(output), false);
        var ids = new DocxStoryReader().Inspect(generated).PackageXml.SelectMany(item => item.Xml.Descendants(
            XName.Get("docPr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")))
            .Select(e => (string?)e.Attribute("id")).ToArray();
        Assert.Equal(ids.Length, ids.Distinct().Count());
        foreach (var part in original.MainDocumentPart!.ImageParts)
        {
            using var before = part.GetStream(); using var after = generated.MainDocumentPart!.GetPartById(original.MainDocumentPart.GetIdOfPart(part)).GetStream();
            using var b = new MemoryStream(); using var a = new MemoryStream(); before.CopyTo(b); after.CopyTo(a);
            Assert.Equal(b.ToArray(), a.ToArray());
        }
        Assert.Single(generated.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements());
    }
}
