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
    public void Enclosing_ranges_and_local_permissions_preserve_ownership(string family, bool rows)
    {
        XElement Container(params object[] contents) => rows
            ? new XElement(W + "tr", new XElement(W + "tc", Docx.P(contents))) : Docx.P(contents);
        var content = new[]
        {
            Container(Start(family, 10), Docx.Run("before")),
            Container(Docx.Run("{For:Items}")),
            family == "permission" ? Container(Start(family, 20), Docx.Run("{.}")) : Container(Docx.Run("{.}")),
            family == "permission" ? Container(Docx.Run("tail"), End(family, 20)) : Container(Docx.Run("tail")),
            Container(Docx.Run("{/For}")),
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
            AssertRanges(xml, family, family == "permission" ? count + 1 : 1);
            Assert.Single(xml.Descendants(W + StartName(family)), e => (string?)e.Attribute(W + "id") == "10");
        }
    }

    [Theory]
    [InlineData("bookmark")] [InlineData("permission")] [InlineData("comment")]
    public void Conditional_local_ranges_reject_except_for_retained_permissions(string family)
    {
        var template = WithComments(Docx.Create(
            Docx.P(Docx.Run("{If:Show}")),
            Docx.P(Start(family, 1), Start(family, 2), Docx.Run("yes"), End(family, 1), End(family, 2)),
            Docx.P(Docx.Run("{Else}")),
            Docx.P(Start(family, 3), Docx.Run("no"), End(family, 3)),
            Docx.P(Docx.Run("{/If}"))), 1, 2, 3);
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        if (family != "permission")
        {
            foreach (var show in new[] { true, false })
                Reject(template, JsonSerializer.Serialize(new { Show = show }), "template.invalid_structure");
            return;
        }
        foreach (var show in new[] { true, false })
        {
            var result = Docx.Merge(template, JsonSerializer.Serialize(new { Show = show }), validationVersion: FileFormatVersions.Office2010);
            var xml = Docx.Read(result.Document);
            Assert.Equal(show ? "yes" : "no", Docx.Text(xml));
            AssertRanges(xml, family, show ? 2 : 1);
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
            new[] { C(Start(family, 0)), C(Docx.Run("{For:Items}")), C(Docx.Run("x"), End(family, 0)), C(Docx.Run("{/For}")) },
            new[] { C(Docx.Run("{For:Items}")), C(Start(family, 0), Docx.Run("x")), C(Docx.Run("{/For}")), C(End(family, 0)) },
            new[] { C(Docx.Run("{For:Items}")), C(Start(family, 0)), C(Docx.Run("{For:Children}")), C(End(family, 0)), C(Docx.Run("{/For}")), C(Docx.Run("{/For}")) },
            new[] { C(Docx.Run("{If:Show}")), C(Start(family, 0)), C(Docx.Run("{Else}")), C(End(family, 0)), C(Docx.Run("{/If}")) },
            new[] { C(Start(family, 0), Docx.Run("{For:Items}")), C(End(family, 0)), C(Docx.Run("{/For}")) }
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
            Reject(Docx.Create(Docx.P(Docx.Run("{For:Items}")), Docx.P(endpoints), Docx.P(Docx.Run("{/For}"))), "{}", "template.invalid_structure");
    }

    [Fact]
    public void Range_endpoints_in_separate_notes_are_not_paired()
    {
        var template = Prepare(Docx.Create(Docx.P(Docx.Run("body"))), document =>
        {
            var part = document.MainDocumentPart!.AddNewPart<FootnotesPart>();
            Save(part, new XElement(W + "footnotes",
                new XElement(W + "footnote", new XAttribute(W + "id", 1), Docx.P(Docx.Run("{For:Items}")),
                    Docx.P(Start("bookmark", 0), Docx.Run("x")), Docx.P(Docx.Run("{/For}"))),
                new XElement(W + "footnote", new XAttribute(W + "id", 2), Docx.P(End("bookmark", 0)))));
        });
        // Cross-root pairing is invalid even if an OOXML validator accepts the package-level IDs.
        Reject(template, "{}", "template.invalid_structure");
    }

    [Fact]
    public void Scoped_links_and_local_targets_are_rejected_without_remapping()
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
            Docx.P(Docx.Run("{If:Deep}")), Docx.P(Docx.Run("{Name}"), Link("Conditional", "deep")), Docx.P(Docx.Run("{/If}")),
            Docx.P(Docx.Run("{/If}")),
            Docx.P(Docx.Run("{For:Children}")),
            Docx.P(Docx.Run("{../Name}/{Name}"), Link(name, "ancestor"), Link("Root", "root"), Link("Child", "forward")),
            Docx.P(Bookmark(4, "Child", "{../Name}/{Name}"), Link("Child", "backward")),
            Docx.P(Docx.Run("{/For}")), Docx.P(Docx.Run("{/For}")));
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        Reject(template,
            """{"Orders":[{"Name":"A","Show":true,"Deep":true,"Children":[{"Name":"a"},{"Name":"b"}]}]}""",
            "template.invalid_structure");
    }

    [Fact]
    public void Links_and_targets_inside_conditionals_are_rejected_without_remapping()
    {
        var template = Docx.Create(Docx.P(Bookmark(0, "Root", "R")),
            Docx.P(Docx.Run("{If:Show}")), Docx.P(Link("Conditional", "local")),
            Docx.P(Bookmark(1, "Conditional", "C")), Docx.P(Docx.Run("{For:Items}")),
            Docx.P(Link("Conditional", "ancestor"), Link("Root", "root")), Docx.P(Docx.Run("{/For}")),
            Docx.P(Docx.Run("{/If}")));
        foreach (var json in new[] { "{}", "{\"Show\":false}", "{\"Show\":true,\"Items\":[]}", "{\"Show\":true,\"Items\":[1,2]}" })
            Reject(template, json, "template.invalid_structure");
    }

    [Fact]
    public void Ambiguous_or_potentially_removed_targets_reject_during_shared_preflight()
    {
        var variants = new[]
        {
            new[] { Docx.P(Link("B0", "outside")), Docx.P(Docx.Run("{If:Show}")), Docx.P(Bookmark(0,"B0","x")), Docx.P(Docx.Run("{/If}")) },
            new[] { Docx.P(Docx.Run("{If:Show}")), Docx.P(Link("B0", "opposite")), Docx.P(Docx.Run("{Else}")), Docx.P(Bookmark(0,"B0","x")), Docx.P(Docx.Run("{/If}")) },
            new[] { Docx.P(Docx.Run("{For:Items}")), Docx.P(Link("B0", "sibling")), Docx.P(Docx.Run("{/For}")), Docx.P(Docx.Run("{For:Other}")), Docx.P(Bookmark(0,"B0","x")), Docx.P(Docx.Run("{/For}")) },
            new[] { Docx.P(Link("B0", "marker")), Docx.P(Start("bookmark",0), Docx.Run("{For:Items}"), End("bookmark",0)), Docx.P(Docx.Run("x")), Docx.P(Docx.Run("{/For}")) },
            new[] { Docx.P(Bookmark(0,"B0","one")), Docx.P(Docx.Run("{For:Items}")), Docx.P(Link("B0", "duplicate")), Docx.P(Bookmark(1,"B0","two")), Docx.P(Docx.Run("{/For}")) }
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
            Docx.P(Docx.Run("{/For}"))), document =>
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
        Assert.Contains(analysis.Errors, error => error.Part!.Contains(kind == "header" ? "header" : kind));
        Assert.DoesNotContain(analysis.Occurrences, e => e.Expression == "{Hidden}");
    }

    [Fact]
    public void Internal_links_inside_operations_reject_even_with_cross_part_targets_and_opaque_fields()
    {
        var field = Docx.Field(" REF Target ", "{Cached}");
        var template = Prepare(Docx.Create(Docx.P(Docx.Run("{For:Items}")),
            Docx.P(Link("Target", "cross"), Link("Missing", "unresolved"), field), Docx.P(Docx.Run("{/For}"))), document =>
        {
            var main = document.MainDocumentPart!;
            Save(main.AddNewPart<HeaderPart>(), new XElement(W + "hdr", Docx.P(Bookmark(0,"Target","header"))));
            var relation = main.AddHyperlinkRelationship(new Uri("https://example.org/report"), true);
            using var source = main.GetStream(); var xml = XDocument.Load(source);
            xml.Descendants(W + "p").ElementAt(1).Add(new XElement(W + "hyperlink", new XAttribute(WordXml.RelationshipsNamespace + "id", relation.Id), Docx.Run("external")));
            source.Close(); Save(main, xml.Root!);
        });
        Reject(template, "{\"Items\":[1,2]}", "template.invalid_structure");
    }

    [Theory]
    [InlineData("bookmark")] [InlineData("permission")] [InlineData("comment")]
    public void Local_table_ranges_reject_except_for_retained_permissions(string family)
    {
        var template = WithComments(Docx.Create(Docx.Table(
            new XElement(W + "tr", new XElement(W + "tc", Docx.P(Start(family, 1), Docx.Run("{For:Items}"), End(family, 1)))),
            new XElement(W + "tr", new XElement(W + "tc", Docx.P(Start(family, 2), Docx.Run("{.}"))), new XElement(W + "tc", Docx.P(End(family, 2)))),
            Docx.Row(Docx.Run("{/For}")))), 1, 2);
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        if (family != "permission")
        {
            Reject(template, "{\"Items\":[\"A\",\"B\"]}", "template.invalid_structure");
            return;
        }
        var output = Docx.Merge(template, "{\"Items\":[\"A\",\"B\"]}", validationVersion: FileFormatVersions.Office2010).Document;
        Assert.Equal("AB", Docx.Text(Docx.Read(output)));
        AssertRanges(Docx.Read(output), family, 2);
    }

    [Fact]
    public void Duplicate_point_attachments_in_one_story_are_not_silently_repaired()
    {
        XElement Reference() => new(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 0)));
        var template = WithComments(Docx.Create(Docx.P(Docx.Run("{For:Items}")), Docx.P(Docx.Run("{.}"), Reference()),
            Docx.P(Docx.Run("second"), Reference()), Docx.P(Docx.Run("{/For}"))), 0);
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
