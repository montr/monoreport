using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MonoReport.Cli.Tests;
using MonoReport.Docx;
using Xunit;

namespace MonoReport.Tests.Docx;

public sealed partial class DocxIdentityTests
{
    [Theory]
    [InlineData("bookmark", false)]
    [InlineData("permission", false)]
    [InlineData("comment", false)]
    [InlineData("bookmark", true)]
    [InlineData("permission", true)]
    [InlineData("comment", true)]
    public void Enclosing_ranges_and_local_permissions_preserve_ownership(string family, bool rows)
    {
        XElement Container(params object[] contents)
        {
            return rows
                ? new XElement(W + "tr", new XElement(W + "tc", DocxFixture.P(contents)))
                : DocxFixture.P(contents);
        }

        var content = new[]
        {
            Container(Start(family, 10), DocxFixture.Run("before")),
            Container(DocxFixture.Run("{For:Items}")),
            family == "permission"
                ? Container(Start(family, 20), DocxFixture.Run("{.}"))
                : Container(DocxFixture.Run("{.}")),
            family == "permission"
                ? Container(DocxFixture.Run("tail"), End(family, 20))
                : Container(DocxFixture.Run("tail")),
            Container(DocxFixture.Run("{/For}")),
            Container(DocxFixture.Run("after"), End(family, 10))
        };
        var template = WithComments(rows ? DocxFixture.Create(DocxFixture.Table(content)) : DocxFixture.Create(content),
            10, 20);
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        using var input = new MemoryStream(template);
        Assert.True(Analyzer().Analyze(input).Valid);
        foreach (var (json, count, text) in new[]
                 {
                     ("{}", 0, "beforeafter"), ("{\"Items\":null}", 0, "beforeafter"),
                     ("{\"Items\":[]}", 0, "beforeafter"), ("{\"Items\":[\"A\"]}", 1, "beforeAtailafter"),
                     ("{\"Items\":[\"A\",\"B\"]}", 2, "beforeAtailBtailafter")
                 })
        {
            var result = DocxFixture.Generate(template, json, validationVersion: FileFormatVersions.Office2010);
            var xml = DocxFixture.Read(result.Document);
            Assert.Equal(text, DocxFixture.Text(xml));
            AssertRanges(xml, family, family == "permission" ? count + 1 : 1);
            Assert.Single(xml.Descendants(W + StartName(family)), e => (string?) e.Attribute(W + "id") == "10");
        }
    }

    [Theory]
    [InlineData("bookmark")]
    [InlineData("permission")]
    [InlineData("comment")]
    public void Conditional_local_ranges_reject_except_for_retained_permissions(string family)
    {
        var template = WithComments(DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{If:Show}")),
            DocxFixture.P(Start(family, 1), Start(family, 2), DocxFixture.Run("yes"), End(family, 1), End(family, 2)),
            DocxFixture.P(DocxFixture.Run("{Else}")),
            DocxFixture.P(Start(family, 3), DocxFixture.Run("no"), End(family, 3)),
            DocxFixture.P(DocxFixture.Run("{/If}"))), 1, 2, 3);
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        if (family != "permission")
        {
            foreach (var show in new[] {true, false})
                Reject(template, JsonSerializer.Serialize(new {Show = show}), "template.invalid_structure");
            return;
        }

        foreach (var show in new[] {true, false})
        {
            var result = DocxFixture.Generate(template, JsonSerializer.Serialize(new {Show = show}),
                validationVersion: FileFormatVersions.Office2010);
            var xml = DocxFixture.Read(result.Document);
            Assert.Equal(show ? "yes" : "no", DocxFixture.Text(xml));
            AssertRanges(xml, family, show ? 2 : 1);
        }
    }

    [Theory]
    [InlineData("bookmark", false)]
    [InlineData("permission", false)]
    [InlineData("comment", false)]
    [InlineData("bookmark", true)]
    [InlineData("permission", true)]
    [InlineData("comment", true)]
    public void Crossing_ranges_reject_both_directions_nested_scopes_branches_and_rows(string family, bool rows)
    {
        XElement C(params object[] contents)
        {
            return rows
                ? new XElement(W + "tr", new XElement(W + "tc", DocxFixture.P(contents)))
                : DocxFixture.P(contents);
        }

        var variants = new[]
        {
            new[]
            {
                C(Start(family, 0)), C(DocxFixture.Run("{For:Items}")), C(DocxFixture.Run("x"), End(family, 0)),
                C(DocxFixture.Run("{/For}"))
            },
            new[]
            {
                C(DocxFixture.Run("{For:Items}")), C(Start(family, 0), DocxFixture.Run("x")),
                C(DocxFixture.Run("{/For}")), C(End(family, 0))
            },
            new[]
            {
                C(DocxFixture.Run("{For:Items}")), C(Start(family, 0)), C(DocxFixture.Run("{For:Children}")),
                C(End(family, 0)),
                C(DocxFixture.Run("{/For}")), C(DocxFixture.Run("{/For}"))
            },
            new[]
            {
                C(DocxFixture.Run("{If:Show}")), C(Start(family, 0)), C(DocxFixture.Run("{Else}")), C(End(family, 0)),
                C(DocxFixture.Run("{/If}"))
            },
            new[] {C(Start(family, 0), DocxFixture.Run("{For:Items}")), C(End(family, 0)), C(DocxFixture.Run("{/For}"))}
        };
        foreach (var variant in variants)
        {
            var template =
                WithComments(rows ? DocxFixture.Create(DocxFixture.Table(variant)) : DocxFixture.Create(variant), 0);
            DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
            foreach (var json in new[]
                         {"{}", "{\"Items\":[{\"Children\":[{}]}],\"Show\":true}", "{\"Items\":[],\"Show\":false}"})
                Reject(template, json, "template.invalid_structure");
        }
    }

    [Fact]
    public void Affected_missing_duplicate_or_wrong_order_endpoints_are_rejected_without_repair()
    {
        // Deliberately malformed endpoint probes supplement conformant crossing-range fixtures.
        foreach (var endpoints in new[]
                 {
                     new[] {Start("bookmark", 0)},
                     new[] {Start("bookmark", 0), Start("bookmark", 0), End("bookmark", 0)},
                     new[] {End("bookmark", 0), Start("bookmark", 0)}
                 })
            Reject(
                DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Items}")), DocxFixture.P(endpoints),
                    DocxFixture.P(DocxFixture.Run("{/For}"))), "{}",
                "template.invalid_structure");
    }

    [Fact]
    public void Range_endpoints_in_separate_notes_are_not_paired()
    {
        var template = Prepare(DocxFixture.Create(DocxFixture.P(DocxFixture.Run("body"))), document =>
        {
            var part = document.MainDocumentPart!.AddNewPart<FootnotesPart>();
            Save(part, new XElement(W + "footnotes",
                new XElement(W + "footnote", new XAttribute(W + "id", 1), DocxFixture.P(DocxFixture.Run("{For:Items}")),
                    DocxFixture.P(Start("bookmark", 0), DocxFixture.Run("x")),
                    DocxFixture.P(DocxFixture.Run("{/For}"))),
                new XElement(W + "footnote", new XAttribute(W + "id", 2), DocxFixture.P(End("bookmark", 0)))));
        });
        // Cross-root pairing is invalid even if an OOXML validator accepts the package-level IDs.
        Reject(template, "{}", "template.invalid_structure");
    }

    [Fact]
    public void Scoped_links_and_local_targets_are_rejected_without_remapping()
    {
        var name = new string('B', 38);
        var template = DocxFixture.Create(
            DocxFixture.P(Bookmark(0, "Root", "root")), DocxFixture.P(Bookmark(1, name + "_1", "reserved")),
            DocxFixture.P(DocxFixture.Run("{For:Orders}")),
            DocxFixture.P(DocxFixture.Run("{Name}"), Link(name, "forward")),
            DocxFixture.P(Bookmark(2, name, "{Name}")),
            DocxFixture.P(DocxFixture.Run("{If:Show}")),
            DocxFixture.P(DocxFixture.Run("{Name}"), Link("Conditional", "guarded")),
            DocxFixture.P(Bookmark(3, "Conditional", "{Name}")),
            DocxFixture.P(DocxFixture.Run("{If:Deep}")),
            DocxFixture.P(DocxFixture.Run("{Name}"), Link("Conditional", "deep")),
            DocxFixture.P(DocxFixture.Run("{/If}")),
            DocxFixture.P(DocxFixture.Run("{/If}")),
            DocxFixture.P(DocxFixture.Run("{For:Children}")),
            DocxFixture.P(DocxFixture.Run("{../Name}/{Name}"), Link(name, "ancestor"), Link("Root", "root"),
                Link("Child", "forward")),
            DocxFixture.P(Bookmark(4, "Child", "{../Name}/{Name}"), Link("Child", "backward")),
            DocxFixture.P(DocxFixture.Run("{/For}")), DocxFixture.P(DocxFixture.Run("{/For}")));
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        Reject(template,
            """{"Orders":[{"Name":"A","Show":true,"Deep":true,"Children":[{"Name":"a"},{"Name":"b"}]}]}""",
            "template.invalid_structure");
    }

    [Fact]
    public void Links_and_targets_inside_conditionals_are_rejected_without_remapping()
    {
        var template = DocxFixture.Create(DocxFixture.P(Bookmark(0, "Root", "R")),
            DocxFixture.P(DocxFixture.Run("{If:Show}")), DocxFixture.P(Link("Conditional", "local")),
            DocxFixture.P(Bookmark(1, "Conditional", "C")), DocxFixture.P(DocxFixture.Run("{For:Items}")),
            DocxFixture.P(Link("Conditional", "ancestor"), Link("Root", "root")),
            DocxFixture.P(DocxFixture.Run("{/For}")),
            DocxFixture.P(DocxFixture.Run("{/If}")));
        foreach (var json in new[]
                     {"{}", "{\"Show\":false}", "{\"Show\":true,\"Items\":[]}", "{\"Show\":true,\"Items\":[1,2]}"})
            Reject(template, json, "template.invalid_structure");
    }

    [Fact]
    public void Ambiguous_or_potentially_removed_targets_reject_during_shared_preflight()
    {
        var variants = new[]
        {
            new[]
            {
                DocxFixture.P(Link("B0", "outside")), DocxFixture.P(DocxFixture.Run("{If:Show}")),
                DocxFixture.P(Bookmark(0, "B0", "x")),
                DocxFixture.P(DocxFixture.Run("{/If}"))
            },
            new[]
            {
                DocxFixture.P(DocxFixture.Run("{If:Show}")), DocxFixture.P(Link("B0", "opposite")),
                DocxFixture.P(DocxFixture.Run("{Else}")),
                DocxFixture.P(Bookmark(0, "B0", "x")), DocxFixture.P(DocxFixture.Run("{/If}"))
            },
            new[]
            {
                DocxFixture.P(DocxFixture.Run("{For:Items}")), DocxFixture.P(Link("B0", "sibling")),
                DocxFixture.P(DocxFixture.Run("{/For}")),
                DocxFixture.P(DocxFixture.Run("{For:Other}")), DocxFixture.P(Bookmark(0, "B0", "x")),
                DocxFixture.P(DocxFixture.Run("{/For}"))
            },
            new[]
            {
                DocxFixture.P(Link("B0", "marker")),
                DocxFixture.P(Start("bookmark", 0), DocxFixture.Run("{For:Items}"), End("bookmark", 0)),
                DocxFixture.P(DocxFixture.Run("x")), DocxFixture.P(DocxFixture.Run("{/For}"))
            },
            new[]
            {
                DocxFixture.P(Bookmark(0, "B0", "one")), DocxFixture.P(DocxFixture.Run("{For:Items}")),
                DocxFixture.P(Link("B0", "duplicate")),
                DocxFixture.P(Bookmark(1, "B0", "two")), DocxFixture.P(DocxFixture.Run("{/For}"))
            }
        };
        foreach (var variant in variants)
        {
            var template = DocxFixture.Create(variant);
            // Duplicate-name case deliberately probes ambiguous, nonconformant input; other sources are conformant.
            if (variant != variants[^1]) DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
            Reject(template, "{\"Show\":true,\"Items\":[{}],\"Other\":[{}]}", "template.invalid_structure");
        }
    }

    [Theory]
    [InlineData("header")]
    [InlineData("comments")]
    [InlineData("glossary")]
    public void Cross_part_anchors_into_operated_targets_are_inspected_without_scanning_auxiliary_expressions(
        string kind)
    {
        var template = Prepare(DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Items}")),
            DocxFixture.P(Bookmark(0, "B0", "x")),
            DocxFixture.P(DocxFixture.Run("{/For}"))), document =>
        {
            var main = document.MainDocumentPart!;
            if (kind == "header")
                Save(main.AddNewPart<HeaderPart>(), new XElement(W + "hdr", DocxFixture.P(Link("B0", "link"))));
            else if (kind == "comments")
                Save(main.AddNewPart<WordprocessingCommentsPart>(), new XElement(W + "comments",
                    new XElement(W + "comment", new XAttribute(W + "id", 0), new XAttribute(W + "author", "A"),
                        DocxFixture.P(Link("B0", "{Hidden}")))));
            else
                Save(main.AddNewPart<GlossaryDocumentPart>(), new XElement(W + "glossaryDocument", new XElement(
                    W + "docParts",
                    new XElement(W + "docPart",
                        new XElement(W + "docPartPr", new XElement(W + "name", new XAttribute(W + "val", "entry"))),
                        new XElement(W + "docPartBody", DocxFixture.P(Link("B0", "{Hidden}")))))));
        });
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{}", "template.invalid_structure");
        Assert.Contains(analysis.Errors, error => error.Part!.Contains(kind == "header" ? "header" : kind));
        Assert.DoesNotContain(analysis.Occurrences, e => e.Expression == "{Hidden}");
    }

    [Fact]
    public void Internal_links_inside_operations_reject_even_with_cross_part_targets_and_opaque_fields()
    {
        var field = DocxFixture.Field(" REF Target ", "{Cached}");
        var template = Prepare(DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Items}")),
                DocxFixture.P(Link("Target", "cross"), Link("Missing", "unresolved"), field),
                DocxFixture.P(DocxFixture.Run("{/For}"))),
            document =>
            {
                var main = document.MainDocumentPart!;
                Save(main.AddNewPart<HeaderPart>(),
                    new XElement(W + "hdr", DocxFixture.P(Bookmark(0, "Target", "header"))));
                var relation = main.AddHyperlinkRelationship(new Uri("https://example.org/report"), true);
                using var source = main.GetStream();
                var xml = XDocument.Load(source);
                xml.Descendants(W + "p").ElementAt(1).Add(new XElement(W + "hyperlink",
                    new XAttribute(WordXml.RelationshipsNamespace + "id", relation.Id), DocxFixture.Run("external")));
                source.Close();
                Save(main, xml.Root!);
            });
        Reject(template, "{\"Items\":[1,2]}", "template.invalid_structure");
    }

    [Theory]
    [InlineData("bookmark")]
    [InlineData("permission")]
    [InlineData("comment")]
    public void Local_table_ranges_reject_except_for_retained_permissions(string family)
    {
        var template = WithComments(DocxFixture.Create(DocxFixture.Table(
            new XElement(W + "tr",
                new XElement(W + "tc",
                    DocxFixture.P(Start(family, 1), DocxFixture.Run("{For:Items}"), End(family, 1)))),
            new XElement(W + "tr", new XElement(W + "tc", DocxFixture.P(Start(family, 2), DocxFixture.Run("{.}"))),
                new XElement(W + "tc", DocxFixture.P(End(family, 2)))),
            DocxFixture.Row(DocxFixture.Run("{/For}")))), 1, 2);
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        if (family != "permission")
        {
            Reject(template, "{\"Items\":[\"A\",\"B\"]}", "template.invalid_structure");
            return;
        }

        var output = DocxFixture.Generate(template, "{\"Items\":[\"A\",\"B\"]}",
                validationVersion: FileFormatVersions.Office2010)
            .Document;
        Assert.Equal("AB", DocxFixture.Text(DocxFixture.Read(output)));
        AssertRanges(DocxFixture.Read(output), family, 2);
    }

    [Fact]
    public void Duplicate_point_attachments_in_one_story_are_not_silently_repaired()
    {
        XElement Reference()
        {
            return new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 0)));
        }

        var template = WithComments(DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Items}")),
            DocxFixture.P(DocxFixture.Run("{.}"), Reference()),
            DocxFixture.P(DocxFixture.Run("second"), Reference()), DocxFixture.P(DocxFixture.Run("{/For}"))), 0);
        // The SDK rejects this duplicate-reference source; supplement valid ranged/point fixtures with preflight rejection.
        using var document = WordprocessingDocument.Open(new MemoryStream(template), false);
        Assert.Contains(new OpenXmlValidator(FileFormatVersions.Office2010).Validate(document),
            e => e.Description.Contains("unique value", StringComparison.Ordinal));
        Reject(template, "{\"Items\":[\"A\",\"B\"]}", "template.invalid_structure");
    }

    [Fact]
    public void Unreadable_auxiliary_xml_has_shared_location_bearing_rejection()
    {
        var template = Prepare(DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Name}"))), document =>
        {
            var comments = document.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>();
            using var stream = comments.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream);
            writer.Write("<broken");
        });
        var analysis = Reject(template, "{\"Name\":\"OK\"}", "template.invalid_xml");
        Assert.Equal("/word/comments.xml", Assert.Single(analysis.Errors).Location);
    }

    internal static XElement Start(string kind, int id)
    {
        return new XElement(W + StartName(kind), new XAttribute(W + "id", id),
            kind == "bookmark" ? new XAttribute(W + "name", "B" + id) :
            kind == "permission" ? new XAttribute(W + "edGrp", "everyone") : null);
    }

    internal static XElement End(string kind, int id)
    {
        return kind == "comment"
            ? new XElement("holder", new XElement(W + "commentRangeEnd", new XAttribute(W + "id", id)),
                new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", id))))
            : new XElement(W + (kind == "bookmark" ? "bookmarkEnd" : "permEnd"), new XAttribute(W + "id", id));
    }

    private static string StartName(string kind)
    {
        return kind == "bookmark" ? "bookmarkStart" : kind == "permission" ? "permStart" : "commentRangeStart";
    }

    internal static object[] Bookmark(int id, string name, string text)
    {
        return
        [
            new XElement(W + "bookmarkStart", new XAttribute(W + "id", id), new XAttribute(W + "name", name)),
            DocxFixture.Run(text), new XElement(W + "bookmarkEnd", new XAttribute(W + "id", id))
        ];
    }

    internal static XElement Link(string target, string text)
    {
        return new XElement(W + "hyperlink", new XAttribute(W + "anchor", target), DocxFixture.Run(text));
    }

    private static void AssertRanges(XDocument xml, string kind, int count)
    {
        var starts = xml.Descendants(W + StartName(kind)).Select(e => (string) e.Attribute(W + "id")!).ToArray();
        var ends = xml
            .Descendants(
                W + (kind == "bookmark" ? "bookmarkEnd" : kind == "permission" ? "permEnd" : "commentRangeEnd"))
            .Select(e => (string) e.Attribute(W + "id")!).ToArray();
        Assert.Equal(count, starts.Length);
        Assert.Equal(count, starts.Distinct().Count());
        Assert.Equal(starts.Order(), ends.Order());
        if (kind == "comment")
            Assert.Equal(starts.Order(),
                xml.Descendants(W + "commentReference").Select(e => (string) e.Attribute(W + "id")!).Order());
        var ordered = xml.Descendants().ToArray();
        foreach (var id in starts)
            Assert.True(Array.FindIndex(ordered,
                            e => e.Name == W + StartName(kind) && (string?) e.Attribute(W + "id") == id) <
                        Array.FindIndex(ordered,
                            e => e.Name == W + (kind == "bookmark" ? "bookmarkEnd" :
                                     kind == "permission" ? "permEnd" : "commentRangeEnd") &&
                                 (string?) e.Attribute(W + "id") == id));
    }

    internal static byte[] Prepare(byte[] template, Action<WordprocessingDocument> action)
    {
        using var stream = new MemoryStream();
        stream.Write(template);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            action(document);
        }

        return stream.ToArray();
    }

    internal static void Save(OpenXmlPart part, XElement root)
    {
        using var target = part.GetStream(FileMode.Create, FileAccess.Write);
        new XDocument(root).Save(target);
    }

    private static byte[] WithComments(byte[] template, params int[] ids)
    {
        return Prepare(template, document =>
        {
            // Unwrap helper-only attachment containers; never write a non-OOXML wrapper into a fixture.
            using var source = document.MainDocumentPart!.GetStream();
            var xml = XDocument.Load(source);
            source.Close();
            foreach (var holder in xml.Descendants("holder").ToArray()) holder.ReplaceWith(holder.Elements().ToArray());
            Save(document.MainDocumentPart, xml.Root!);
            Save(document.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>(), new XElement(W + "comments",
                ids.Select(id =>
                    new XElement(W + "comment", new XAttribute(W + "id", id), new XAttribute(W + "author", "A"),
                        DocxFixture.P(DocxFixture.Run("definition" + id))))));
        });
    }
}