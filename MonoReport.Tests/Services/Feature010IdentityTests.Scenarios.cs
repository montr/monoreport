using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed partial class Feature010IdentityTests
{
    [Theory]
    [InlineData("bookmark", false)] [InlineData("permission", false)] [InlineData("comment", false)]
    [InlineData("bookmark", true)] [InlineData("permission", true)] [InlineData("comment", true)]
    public void Ranges_local_to_sibling_roots_and_enclosing_operations_preserve_ownership(string family, bool rows)
    {
        XElement Container(params object[] contents) => rows
            ? new XElement(W + "tr", new XElement(W + "tc", Docx.P(contents))) : Docx.P(contents);
        var content = new[]
        {
            Container(Start(family, 10), Docx.Run("before")),
            Container(Docx.Run("{For:Items}")),
            Container(Start(family, 20), Docx.Run("{.}")),
            Container(Docx.Run("tail"), End(family, 20)),
            Container(Docx.Run("{EndFor}")),
            Container(Docx.Run("after"), End(family, 10))
        };
        var template = WithComments(rows ? Docx.Create(Docx.Table(content)) : Docx.Create(content), 10, 20);
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        using var input = new MemoryStream(template);
        Assert.True(Analyzer().Analyze(input).Valid);
        foreach (var (json, count, text) in new[]
        {
            ("{}", 0, "beforeafter"), ("{\"Items\":null}", 0, "beforeafter"),
            ("{\"Items\":[]}", 0, "beforeafter"), ("{\"Items\":[\"A\"]}", 1, "beforeAtailafter"),
            ("{\"Items\":[\"A\",\"B\"]}", 2, "beforeAtailBtailafter")
        })
        {
            var result = Docx.Merge(template, json, validationVersion: FileFormatVersions.Office2010);
            var xml = Docx.Read(result.Document);
            Assert.Equal(text, Docx.Text(xml));
            AssertRanges(xml, family, count + 1);
            Assert.Single(xml.Descendants(W + StartName(family)), e => (string?)e.Attribute(W + "id") == "10");
        }
    }

    [Theory]
    [InlineData("bookmark")] [InlineData("permission")] [InlineData("comment")]
    public void Conditional_ranges_in_both_branches_and_overlapping_local_pairs_are_supported(string family)
    {
        var template = WithComments(Docx.Create(
            Docx.P(Docx.Run("{If:Show}")),
            Docx.P(Start(family, 1), Start(family, 2), Docx.Run("yes"), End(family, 1), End(family, 2)),
            Docx.P(Docx.Run("{Else}")),
            Docx.P(Start(family, 3), Docx.Run("no"), End(family, 3)),
            Docx.P(Docx.Run("{EndIf}"))), 1, 2, 3);
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        foreach (var show in new[] { true, false })
        {
            var result = Docx.Merge(template, JsonSerializer.Serialize(new { Show = show }), validationVersion: FileFormatVersions.Office2010);
            var xml = Docx.Read(result.Document);
            Assert.Equal(show ? "yes" : "no", Docx.Text(xml));
            AssertRanges(xml, family, show ? 2 : 1);
            if (family == "bookmark") Assert.Equal(show ? new[] { "B1", "B2" } : ["B3"],
                xml.Descendants(W + "bookmarkStart").Select(e => (string)e.Attribute(W + "name")!));
        }
    }

    [Theory]
    [InlineData("bookmark", false)] [InlineData("permission", false)] [InlineData("comment", false)]
    [InlineData("bookmark", true)] [InlineData("permission", true)] [InlineData("comment", true)]
    public void Crossing_ranges_reject_both_directions_nested_scopes_branches_and_rows(string family, bool rows)
    {
        XElement C(params object[] contents) => rows ? new XElement(W + "tr", new XElement(W + "tc", Docx.P(contents))) : Docx.P(contents);
        var variants = new[]
        {
            new[] { C(Start(family, 0)), C(Docx.Run("{For:Items}")), C(Docx.Run("x"), End(family, 0)), C(Docx.Run("{EndFor}")) },
            new[] { C(Docx.Run("{For:Items}")), C(Start(family, 0), Docx.Run("x")), C(Docx.Run("{EndFor}")), C(End(family, 0)) },
            new[] { C(Docx.Run("{For:Items}")), C(Start(family, 0)), C(Docx.Run("{For:Children}")), C(End(family, 0)), C(Docx.Run("{EndFor}")), C(Docx.Run("{EndFor}")) },
            new[] { C(Docx.Run("{If:Show}")), C(Start(family, 0)), C(Docx.Run("{Else}")), C(End(family, 0)), C(Docx.Run("{EndIf}")) },
            new[] { C(Start(family, 0), Docx.Run("{For:Items}")), C(End(family, 0)), C(Docx.Run("{EndFor}")) }
        };
        foreach (var variant in variants)
        {
            var template = WithComments(rows ? Docx.Create(Docx.Table(variant)) : Docx.Create(variant), 0);
            Docx.AssertValid(template, FileFormatVersions.Office2010);
            foreach (var json in new[] { "{}", "{\"Items\":[{\"Children\":[{}]}],\"Show\":true}", "{\"Items\":[],\"Show\":false}" })
                Reject(template, json, "template.invalid_structure");
        }
    }

    [Fact]
    public void Affected_missing_duplicate_or_wrong_order_endpoints_are_rejected_without_repair()
    {
        // Deliberately malformed endpoint probes supplement conformant crossing-range fixtures.
        foreach (var endpoints in new[] { new[] { Start("bookmark", 0) },
                     new[] { Start("bookmark", 0), Start("bookmark", 0), End("bookmark", 0) },
                     new[] { End("bookmark", 0), Start("bookmark", 0) } })
            Reject(Docx.Create(Docx.P(Docx.Run("{For:Items}")), Docx.P(endpoints), Docx.P(Docx.Run("{EndFor}"))), "{}", "template.invalid_structure");
    }

    [Fact]
    public void Range_endpoints_in_separate_notes_are_not_paired()
    {
        var template = Prepare(Docx.Create(Docx.P(Docx.Run("body"))), document =>
        {
            var part = document.MainDocumentPart!.AddNewPart<FootnotesPart>();
            Save(part, new XElement(W + "footnotes",
                new XElement(W + "footnote", new XAttribute(W + "id", 1), Docx.P(Docx.Run("{For:Items}")),
                    Docx.P(Start("bookmark", 0), Docx.Run("x")), Docx.P(Docx.Run("{EndFor}"))),
                new XElement(W + "footnote", new XAttribute(W + "id", 2), Docx.P(End("bookmark", 0)))));
        });
        // Cross-root pairing is invalid even if an OOXML validator accepts the package-level IDs.
        Reject(template, "{}", "template.invalid_structure");
    }

    [Fact]
    public void Scoped_links_bind_forward_backward_local_and_ancestor_targets_to_the_right_iteration()
    {
        var name = new string('B', 38);
        var template = Docx.Create(
            Docx.P(Bookmark(0, "Root", "root")), Docx.P(Bookmark(1, name + "_1", "reserved")),
            Docx.P(Docx.Run("{For:Orders}")),
            Docx.P(Docx.Run("{Name}"), Link(name, "forward")),
            Docx.P(Bookmark(2, name, "{Name}")),
            Docx.P(Docx.Run("{If:Show}")),
            Docx.P(Docx.Run("{Name}"), Link("Conditional", "guarded")),
            Docx.P(Bookmark(3, "Conditional", "{Name}")),
            Docx.P(Docx.Run("{If:Deep}")), Docx.P(Docx.Run("{Name}"), Link("Conditional", "deep")), Docx.P(Docx.Run("{EndIf}")),
            Docx.P(Docx.Run("{EndIf}")),
            Docx.P(Docx.Run("{For:Children}")),
            Docx.P(Docx.Run("{../Name}/{Name}"), Link(name, "ancestor"), Link("Root", "root"), Link("Child", "forward")),
            Docx.P(Bookmark(4, "Child", "{../Name}/{Name}"), Link("Child", "backward")),
            Docx.P(Docx.Run("{EndFor}")), Docx.P(Docx.Run("{EndFor}")));
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        using var input = new MemoryStream(template);
        Assert.True(Analyzer().Analyze(input).Valid);
        var output = Docx.Merge(template,
            """{"Orders":[{"Name":"A","Show":true,"Deep":true,"Children":[{"Name":"a"},{"Name":"b"}]},{"Name":"B","Show":true,"Deep":false,"Children":[{"Name":"c"},{"Name":"d"}]}]}""",
            validationVersion: FileFormatVersions.Office2010).Document;
        var xml = Docx.Read(output);
        var starts = xml.Descendants(W + "bookmarkStart").ToArray();
        var names = starts.Select(e => (string)e.Attribute(W + "name")!).ToArray();
        Assert.Equal(names.Length, names.Distinct().Count());
        Assert.All(names, n => Assert.True(n.Length <= 40));
        AssertRanges(xml, "bookmark", 10);
        foreach (var link in xml.Descendants(W + "hyperlink"))
        {
            var target = starts.Single(e => (string?)e.Attribute(W + "name") == (string?)link.Attribute(W + "anchor"));
            var targetText = Docx.Text(target.Parent!);
            var text = Docx.Text(link);
            var ownerText = string.Concat(link.Parent!.Elements(W + "r").Select(Docx.Text));
            if (text == "root") Assert.Equal("root", targetText);
            else if (text is "ancestor") Assert.Equal(ownerText.Split('/')[0], targetText);
            else if (text is "forward" && ownerText.Contains('/')) Assert.StartsWith(ownerText, targetText);
            else if (text is "backward") Assert.StartsWith(ownerText, targetText);
            else if (text is "forward") Assert.Equal(ownerText, targetText);
            else Assert.Equal(ownerText, targetText);
        }
        Directory.CreateDirectory("/private/tmp/feature010-final-corrections/artifacts");
        File.WriteAllBytes("/private/tmp/feature010-final-corrections/artifacts/scoped-links.docx", output);
    }

    [Fact]
    public void Links_and_targets_disappear_together_and_nonrepeated_names_survive_conditionals()
    {
        var template = Docx.Create(Docx.P(Bookmark(0, "Root", "R")),
            Docx.P(Docx.Run("{If:Show}")), Docx.P(Link("Conditional", "local")),
            Docx.P(Bookmark(1, "Conditional", "C")), Docx.P(Docx.Run("{For:Items}")),
            Docx.P(Link("Conditional", "ancestor"), Link("Root", "root")), Docx.P(Docx.Run("{EndFor}")),
            Docx.P(Docx.Run("{EndIf}")));
        foreach (var json in new[] { "{}", "{\"Show\":false}", "{\"Show\":true,\"Items\":[]}", "{\"Show\":true,\"Items\":[1,2]}" })
        {
            var result = Docx.Merge(template, json, validationVersion: FileFormatVersions.Office2010);
            var xml = Docx.Read(result.Document);
            Assert.All(xml.Descendants(W + "hyperlink"), link => Assert.Single(xml.Descendants(W + "bookmarkStart"),
                e => (string?)e.Attribute(W + "name") == (string?)link.Attribute(W + "anchor")));
            Assert.All(xml.Descendants(W + "bookmarkStart"), b => Assert.Contains((string?)b.Attribute(W + "name"), new[] { "Root", "Conditional" }));
        }
    }

    [Fact]
    public void Ambiguous_or_potentially_removed_targets_reject_during_shared_preflight()
    {
        var variants = new[]
        {
            new[] { Docx.P(Link("B0", "outside")), Docx.P(Docx.Run("{If:Show}")), Docx.P(Bookmark(0,"B0","x")), Docx.P(Docx.Run("{EndIf}")) },
            new[] { Docx.P(Docx.Run("{If:Show}")), Docx.P(Link("B0", "opposite")), Docx.P(Docx.Run("{Else}")), Docx.P(Bookmark(0,"B0","x")), Docx.P(Docx.Run("{EndIf}")) },
            new[] { Docx.P(Docx.Run("{For:Items}")), Docx.P(Link("B0", "sibling")), Docx.P(Docx.Run("{EndFor}")), Docx.P(Docx.Run("{For:Other}")), Docx.P(Bookmark(0,"B0","x")), Docx.P(Docx.Run("{EndFor}")) },
            new[] { Docx.P(Link("B0", "marker")), Docx.P(Start("bookmark",0), Docx.Run("{For:Items}"), End("bookmark",0)), Docx.P(Docx.Run("x")), Docx.P(Docx.Run("{EndFor}")) },
            new[] { Docx.P(Bookmark(0,"B0","one")), Docx.P(Docx.Run("{For:Items}")), Docx.P(Link("B0", "duplicate")), Docx.P(Bookmark(1,"B0","two")), Docx.P(Docx.Run("{EndFor}")) }
        };
        foreach (var variant in variants)
        {
            var template = Docx.Create(variant);
            // Duplicate-name case deliberately probes ambiguous, nonconformant input; other sources are conformant.
            if (variant != variants[^1]) Docx.AssertValid(template, FileFormatVersions.Office2010);
            Reject(template, "{\"Show\":true,\"Items\":[{}],\"Other\":[{}]}", "template.invalid_structure");
        }
    }

    [Theory]
    [InlineData("header")] [InlineData("comments")] [InlineData("glossary")]
    public void Cross_part_anchors_into_operated_targets_are_inspected_without_scanning_auxiliary_expressions(string kind)
    {
        var template = Prepare(Docx.Create(Docx.P(Docx.Run("{For:Items}")), Docx.P(Bookmark(0,"B0","x")),
            Docx.P(Docx.Run("{EndFor}"))), document =>
        {
            var main = document.MainDocumentPart!;
            if (kind == "header") Save(main.AddNewPart<HeaderPart>(), new XElement(W + "hdr", Docx.P(Link("B0", "link"))));
            else if (kind == "comments") Save(main.AddNewPart<WordprocessingCommentsPart>(), new XElement(W + "comments",
                new XElement(W + "comment", new XAttribute(W + "id", 0), new XAttribute(W + "author", "A"), Docx.P(Link("B0", "{Hidden}")))));
            else Save(main.AddNewPart<GlossaryDocumentPart>(), new XElement(W + "glossaryDocument", new XElement(W + "docParts",
                new XElement(W + "docPart", new XElement(W + "docPartPr", new XElement(W + "name", new XAttribute(W + "val", "entry"))),
                    new XElement(W + "docPartBody", Docx.P(Link("B0", "{Hidden}")))))));
        });
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{}", "template.invalid_structure");
        Assert.Contains(kind == "header" ? "header" : kind, analysis.Errors[0].Part);
        Assert.DoesNotContain(analysis.Occurrences, e => e.Expression == "{Hidden}");
    }

    [Fact]
    public void Unaffected_cross_part_targets_unresolved_anchors_external_relationships_and_fields_are_preserved()
    {
        var field = Docx.Field(" REF Target ", "{Cached}");
        var template = Prepare(Docx.Create(Docx.P(Docx.Run("{For:Items}")),
            Docx.P(Link("Target", "cross"), Link("Missing", "unresolved"), field), Docx.P(Docx.Run("{EndFor}"))), document =>
        {
            var main = document.MainDocumentPart!;
            Save(main.AddNewPart<HeaderPart>(), new XElement(W + "hdr", Docx.P(Bookmark(0,"Target","header"))));
            var relation = main.AddHyperlinkRelationship(new Uri("https://example.org/report"), true);
            using var source = main.GetStream(); var xml = XDocument.Load(source);
            xml.Descendants(W + "p").ElementAt(1).Add(new XElement(W + "hyperlink", new XAttribute(WordXml.RelationshipsNamespace + "id", relation.Id), Docx.Run("external")));
            source.Close(); Save(main, xml.Root!);
        });
        var output = Docx.Merge(template, "{\"Items\":[1,2]}", validationVersion: FileFormatVersions.Office2010).Document;
        var xml = Docx.Read(output);
        Assert.Equal(new[] { "Target", "Missing", "Target", "Missing" }, xml.Descendants(W + "hyperlink")
            .Where(e => e.Attribute(W + "anchor") is not null).Select(e => (string)e.Attribute(W + "anchor")!));
        Assert.All(xml.Descendants(W + "fldSimple"), e => Assert.True(XNode.DeepEquals(field, e)));
        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        Assert.Equal("https://example.org/report", Assert.Single(document.MainDocumentPart!.HyperlinkRelationships).Uri.ToString());
    }

    [Theory]
    [InlineData("bookmark")] [InlineData("permission")] [InlineData("comment")]
    public void Local_ranges_across_table_cells_and_complete_marker_attachments_are_supported(string family)
    {
        var template = WithComments(Docx.Create(Docx.Table(
            new XElement(W + "tr", new XElement(W + "tc", Docx.P(Start(family, 1), Docx.Run("{For:Items}"), End(family, 1)))),
            new XElement(W + "tr", new XElement(W + "tc", Docx.P(Start(family, 2), Docx.Run("{.}"))), new XElement(W + "tc", Docx.P(End(family, 2)))),
            Docx.Row(Docx.Run("{EndFor}")))), 1, 2);
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        var output = Docx.Merge(template, "{\"Items\":[\"A\",\"B\"]}", validationVersion: FileFormatVersions.Office2010).Document;
        Assert.Equal("AB", Docx.Text(Docx.Read(output)));
        AssertRanges(Docx.Read(output), family, 2);
    }

    [Fact]
    public void Duplicate_point_attachments_in_one_story_are_not_silently_repaired()
    {
        XElement Reference() => new(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 0)));
        var template = WithComments(Docx.Create(Docx.P(Docx.Run("{For:Items}")), Docx.P(Docx.Run("{.}"), Reference()),
            Docx.P(Docx.Run("second"), Reference()), Docx.P(Docx.Run("{EndFor}"))), 0);
        // The SDK rejects this duplicate-reference source; supplement valid ranged/point fixtures with preflight rejection.
        using var document = WordprocessingDocument.Open(new MemoryStream(template), false);
        Assert.Contains(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Office2010).Validate(document),
            e => e.Description.Contains("unique value", StringComparison.Ordinal));
        Reject(template, "{\"Items\":[\"A\",\"B\"]}", "template.invalid_structure");
    }

    [Fact]
    public void Unreadable_auxiliary_xml_has_shared_location_bearing_rejection()
    {
        var template = Prepare(Docx.Create(Docx.P(Docx.Run("{Name}"))), document =>
        {
            var comments = document.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>();
            using var stream = comments.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream); writer.Write("<broken");
        });
        var analysis = Reject(template, "{\"Name\":\"OK\"}", "template.invalid_xml");
        Assert.Equal("/word/comments.xml", Assert.Single(analysis.Errors).Location);
    }

    internal static XElement Start(string kind, int id) => new(W + StartName(kind), new XAttribute(W + "id", id),
        kind == "bookmark" ? new XAttribute(W + "name", "B" + id) : kind == "permission" ? new XAttribute(W + "edGrp", "everyone") : null);
    internal static XElement End(string kind, int id) => kind == "comment"
        ? new XElement("holder", new XElement(W + "commentRangeEnd", new XAttribute(W + "id", id)),
            new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", id))))
        : new XElement(W + (kind == "bookmark" ? "bookmarkEnd" : "permEnd"), new XAttribute(W + "id", id));
    private static string StartName(string kind) => kind == "bookmark" ? "bookmarkStart" : kind == "permission" ? "permStart" : "commentRangeStart";
    internal static object[] Bookmark(int id, string name, string text) =>
        [new XElement(W + "bookmarkStart", new XAttribute(W + "id", id), new XAttribute(W + "name", name)),
            Docx.Run(text), new XElement(W + "bookmarkEnd", new XAttribute(W + "id", id))];
    internal static XElement Link(string target, string text) => new(W + "hyperlink", new XAttribute(W + "anchor", target), Docx.Run(text));

    private static void AssertRanges(XDocument xml, string kind, int count)
    {
        var starts = xml.Descendants(W + StartName(kind)).Select(e => (string)e.Attribute(W + "id")!).ToArray();
        var ends = xml.Descendants(W + (kind == "bookmark" ? "bookmarkEnd" : kind == "permission" ? "permEnd" : "commentRangeEnd"))
            .Select(e => (string)e.Attribute(W + "id")!).ToArray();
        Assert.Equal(count, starts.Length); Assert.Equal(count, starts.Distinct().Count());
        Assert.Equal(starts.Order(), ends.Order());
        if (kind == "comment") Assert.Equal(starts.Order(), xml.Descendants(W + "commentReference").Select(e => (string)e.Attribute(W + "id")!).Order());
        var ordered = xml.Descendants().ToArray();
        foreach (var id in starts)
            Assert.True(Array.FindIndex(ordered, e => e.Name == W + StartName(kind) && (string?)e.Attribute(W + "id") == id) <
                Array.FindIndex(ordered, e => e.Name == W + (kind == "bookmark" ? "bookmarkEnd" : kind == "permission" ? "permEnd" : "commentRangeEnd") && (string?)e.Attribute(W + "id") == id));
    }

    internal static byte[] Prepare(byte[] template, Action<WordprocessingDocument> action)
    {
        using var stream = new MemoryStream(); stream.Write(template); stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true)) action(document);
        return stream.ToArray();
    }
    internal static void Save(OpenXmlPart part, XElement root)
    {
        using var target = part.GetStream(FileMode.Create, FileAccess.Write);
        new XDocument(root).Save(target);
    }
    private static byte[] WithComments(byte[] template, params int[] ids) => Prepare(template, document =>
    {
        // Unwrap helper-only attachment containers; never write a non-OOXML wrapper into a fixture.
        using var source = document.MainDocumentPart!.GetStream(); var xml = XDocument.Load(source); source.Close();
        foreach (var holder in xml.Descendants("holder").ToArray()) holder.ReplaceWith(holder.Elements().ToArray());
        Save(document.MainDocumentPart, xml.Root!);
        Save(document.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>(), new XElement(W + "comments", ids.Select(id =>
            new XElement(W + "comment", new XAttribute(W + "id", id), new XAttribute(W + "author", "A"), Docx.P(Docx.Run("definition" + id))))));
    });
}
