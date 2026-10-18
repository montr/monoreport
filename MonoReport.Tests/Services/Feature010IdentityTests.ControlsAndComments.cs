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
    [InlineData("{Na", "me}", "")]
    [InlineData("{Na", "", "me}")]
    [InlineData("{Na", "", "")]
    [InlineData("", "{Na", "me}")]
    public void Content_control_entry_exit_and_empty_controls_never_join_expression_fragments(string before, string hidden, string after)
    {
        var control = Control(hidden.Length == 0 ? [] : [Docx.Run(hidden)]);
        var template = Docx.Create(Docx.P(Docx.Run(before), control, Control([]), Docx.Run(after)));
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{\"Name\":\"CROSSED\"}", "template.unmatched_brace");
        Assert.DoesNotContain(analysis.Occurrences, e => e.Expression == "{Name}");
    }

    [Theory]
    [InlineData("inline")] [InlineData("block")] [InlineData("row")] [InlineData("cell")]
    public void Controls_are_opaque_and_preserved_when_an_outer_operation_moves_them_intact(string placement)
    {
        var inner = Control(Docx.Run("{Hidden}"));
        var control = placement switch
        {
            "inline" => Control(Docx.Run("{For:Hidden}"), inner, Control([]), Control(Docx.Run("{EndFor}"))),
            "block" => Control(Docx.P(Docx.Run("{For:Hidden}")), Docx.P(Control(Docx.Run("{Name}"))), Docx.P(Docx.Run("{EndFor}"))),
            "row" => Control(Docx.Row(Docx.Run("{For:Hidden}")), Docx.Row(Docx.Run("{Name}")), Docx.Row(Docx.Run("{EndFor}"))),
            _ => Control(new XElement(W + "tc", Docx.P(Docx.Run("{For:Hidden}")), Docx.P(Docx.Run("{Name}")), Docx.P(Docx.Run("{EndFor}"))))
        };
        control.Element(W + "sdtPr")!.Add(new XElement(W + "dataBinding", new XAttribute(W + "xpath", "/root/value"),
            new XAttribute(W + "storeItemID", "{01234567-89AB-CDEF-0123-456789ABCDEF}")), new XElement(W + "lock", new XAttribute(W + "val", "sdtLocked")));
        XElement body = placement switch
        {
            "inline" => Docx.P(control, Docx.Run("{Name}")),
            "row" => Docx.Table(control),
            "cell" => Docx.Table(new XElement(W + "tr", control)),
            _ => control
        };
        var template = Docx.Create(Docx.P(Docx.Run("{For:Items}")), body, Docx.P(Docx.Run("{EndFor}")), Docx.P(Docx.Run("{Visible}")));
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        using var input = new MemoryStream(template);
        var analysis = Analyzer().Analyze(input);
        Assert.True(analysis.Valid);
        Assert.DoesNotContain(analysis.Occurrences, e => e.Expression.Contains("Hidden") || (placement != "inline" && e.Expression == "{Name}"));
        var output = Docx.Merge(template, "{\"Items\":[{\"Name\":\"A\"},{\"Name\":\"B\"}],\"Visible\":\"OK\"}", validationVersion: FileFormatVersions.Office2010).Document;
        var xml = Docx.Read(output);
        var copies = xml.Descendants(W + "sdt").Where(e => !e.Ancestors(W + "sdt").Any()).ToArray();
        Assert.Equal(2, copies.Length);
        Assert.All(copies, copy => Assert.True(XNode.DeepEquals(control, copy)));
        Assert.EndsWith("OK", Docx.Text(xml));
        if (placement == "inline") Assert.Contains("A", Docx.Text(xml));
    }

    [Fact]
    public void Control_relationships_and_hidden_revision_preflight_remain_independent_of_expression_opacity()
    {
        var template = Prepare(Docx.Create(Docx.P(Control(Docx.Run("{Hidden}")), Docx.Run("{Visible}"))), document =>
        {
            var main = document.MainDocumentPart!;
            var relationship = main.AddHyperlinkRelationship(new Uri("https://example.org/control"), true);
            using var input = main.GetStream(); var xml = XDocument.Load(input); input.Close();
            xml.Descendants(W + "sdtContent").Single().Add(new XElement(W + "hyperlink", new XAttribute(WordXml.RelationshipsNamespace + "id", relationship.Id), Docx.Run("link")));
            Save(main, xml.Root!);
        });
        var output = Docx.Merge(template, "{\"Visible\":\"OK\"}", validationVersion: FileFormatVersions.Office2010).Document;
        Assert.True(XNode.DeepEquals(Docx.Read(template).Descendants(W + "sdt").Single(), Docx.Read(output).Descendants(W + "sdt").Single()));
        using (var document = WordprocessingDocument.Open(new MemoryStream(output), false))
            Assert.Equal("https://example.org/control", Assert.Single(document.MainDocumentPart!.HyperlinkRelationships).Uri.ToString());
        var revised = Docx.Create(Docx.P(Control(new XElement(W + "ins", new XAttribute(W + "id", 1), new XAttribute(W + "author", "A"), Docx.Run("{Hidden}")))));
        Docx.AssertValid(revised, FileFormatVersions.Office2010);
        Reject(revised, "{}", "template.tracked_revision");
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void Nested_comment_copies_normalize_all_drawings_and_local_definition_ranges(bool point)
    {
        var source = Fixture("comment-with-image", ".template.docx");
        var template = Prepare(source, document =>
        {
            var main = document.MainDocumentPart!;
            using var input = main.GetStream(); var xml = XDocument.Load(input); input.Close();
            var body = xml.Descendants(W + "body").Single();
            var attachment = new XElement(body.Elements(W + "p").ElementAt(1));
            attachment.Descendants(W + "t").First().Value = "{Name}";
            if (point)
            {
                attachment.Descendants(W + "commentRangeStart").Remove();
                attachment.Descendants(W + "commentRangeEnd").Remove();
            }
            body.ReplaceNodes(Docx.P(Docx.Run("{For:Items}")), Docx.P(Docx.Run("{If:Show}")),
                Docx.P(Docx.Run("{For:Children}")), attachment, Docx.P(Docx.Run("{EndFor}")),
                Docx.P(Docx.Run("{EndIf}")), Docx.P(Docx.Run("{EndFor}")));
            Save(main, xml.Root!);
            var commentsPart = main.WordprocessingCommentsPart!;
            using var commentsInput = commentsPart.GetStream(); var comments = XDocument.Load(commentsInput); commentsInput.Close();
            var definition = comments.Root!.Elements().Single();
            var drawingRun = new XElement(definition.Descendants(W + "drawing").First().Parent!);
            drawingRun.Descendants(Wp + "docPr").Single().SetAttributeValue("id", 2);
            definition.Add(Docx.P(Start("bookmark", 4), Start("permission", 5), Docx.Run("opaque {Text}"),
                End("permission", 5), End("bookmark", 4), Link("B4", "local"), drawingRun));
            Save(commentsPart, comments.Root);
            // Reserve identities in untouched main/header/footer/glossary XML as well as definitions.
            var drawing = new XElement(definition.Descendants(W + "drawing").First());
            var imagePart = commentsPart.GetPartById((string)drawing.Descendants(XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main") + "blip").Single().Attribute(WordXml.RelationshipsNamespace + "embed")!);
            var header = main.AddNewPart<HeaderPart>(); var footer = main.AddNewPart<FooterPart>();
            var glossary = main.AddNewPart<GlossaryDocumentPart>();
            foreach (var (part, id) in new (OpenXmlPart, uint)[] { (main, 3), (header, 4), (footer, 5), (glossary, 6) })
            {
                var copy = new XElement(drawing); copy.Descendants(Wp + "docPr").Single().SetAttributeValue("id", id);
                var relation = part.AddPart(imagePart);
                copy.Descendants(XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main") + "blip").Single()
                    .SetAttributeValue(WordXml.RelationshipsNamespace + "embed", part.GetIdOfPart(relation));
                var paragraph = Docx.P(new XElement(W + "r", copy));
                if (part == main) { xml.Descendants(W + "body").Single().AddFirst(paragraph); Save(main, xml.Root!); }
                else if (part == header) Save(part, new XElement(W + "hdr", paragraph));
                else if (part == footer) Save(part, new XElement(W + "ftr", paragraph));
                else Save(part, new XElement(W + "glossaryDocument", new XElement(W + "docParts", new XElement(W + "docPart",
                    new XElement(W + "docPartPr", new XElement(W + "name", new XAttribute(W + "val", "entry"))), new XElement(W + "docPartBody", paragraph)))));
            }
        });
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        foreach (var json in new[] { "{\"Items\":[]}", """{"Items":[{"Show":true,"Children":[{"Name":"A"},{"Name":"B"}]},{"Show":false,"Children":[{"Name":"NO"}]}]}""" })
        {
            var output = Docx.Merge(template, json, validationVersion: FileFormatVersions.Office2010).Document;
            using var generated = WordprocessingDocument.Open(new MemoryStream(output), false);
            var package = PackageXml(generated);
            var ids = package.SelectMany(xml => xml.Descendants(Wp + "docPr")).Select(e => (string)e.Attribute("id")!).ToArray();
            Assert.Equal(ids.Length, ids.Distinct().Count());
            var names = package.SelectMany(xml => xml.Descendants(W + "bookmarkStart")).Select(e => (string)e.Attribute(W + "name")!).ToArray();
            Assert.Equal(names.Length, names.Distinct().Count());
            foreach (var comment in package.SelectMany(xml => xml.Descendants(W + "comment")))
            {
                Assert.Single(comment.Descendants(W + "bookmarkStart"));
                Assert.Single(comment.Descendants(W + "bookmarkEnd"));
                Assert.Single(comment.Descendants(W + "permStart"));
                Assert.Single(comment.Descendants(W + "permEnd"));
                Assert.Equal((string?)comment.Descendants(W + "bookmarkStart").Single().Attribute(W + "name"),
                    (string?)comment.Descendants(W + "hyperlink").Single().Attribute(W + "anchor"));
                Assert.Equal((string?)comment.Descendants(W + "bookmarkStart").Single().Attribute(W + "id"),
                    (string?)comment.Descendants(W + "bookmarkEnd").Single().Attribute(W + "id"));
                Assert.Equal((string?)comment.Descendants(W + "permStart").Single().Attribute(W + "id"),
                    (string?)comment.Descendants(W + "permEnd").Single().Attribute(W + "id"));
            }
            var references = Docx.Read(output).Descendants(W + "commentReference").Select(e => (string)e.Attribute(W + "id")!).ToArray();
            Assert.Equal(json.Contains("Name") ? 2 : 0, references.Length);
            Assert.Equal(references.Length, references.Distinct().Count());
            foreach (var reference in references) Assert.Single(package.SelectMany(xml => xml.Descendants(W + "comment")), e => (string?)e.Attribute(W + "id") == reference);
            Assert.DoesNotContain("NO", Docx.Text(Docx.Read(output)));
        }
    }

    [Fact]
    public async Task Reused_engine_isolates_real_comment_drawing_allocations_sequentially_and_concurrently()
    {
        var template = Fixture("comment-with-image", ".template.docx");
        var engine = Docx.CreateEngine();
        byte[] Run()
        {
            using var input = new MemoryStream(template); using var data = JsonDocument.Parse("{\"Items\":[\"A\",\"B\"]}");
            var output = engine.Merge(input, data).Document;
            Docx.AssertValid(output, FileFormatVersions.Office2010);
            AssertCommentImages(template, output, 2);
            Assert.Equal(template, input.ToArray()); Assert.True(input.CanRead);
            return output;
        }
        Run(); Run();
        await Task.WhenAll(Enumerable.Range(0, 8).Select(_ => Task.Run(Run)));
        Directory.CreateDirectory("/private/tmp/feature010-final-corrections/artifacts");
        File.WriteAllBytes("/private/tmp/feature010-final-corrections/artifacts/comment-images.docx", Run());
    }

    private static XElement Control(params object[] content) => new(W + "sdt", new XElement(W + "sdtPr",
        new XElement(W + "alias", new XAttribute(W + "val", "opaque")), new XElement(W + "tag", new XAttribute(W + "val", "keep"))),
        new XElement(W + "sdtContent", content));

    private static XDocument[] PackageXml(WordprocessingDocument document)
    {
        var result = new List<XDocument>(); var queue = new Queue<OpenXmlPart>(document.Parts.Select(p => p.OpenXmlPart)); var seen = new HashSet<OpenXmlPart>();
        while (queue.TryDequeue(out var part))
        {
            if (!seen.Add(part)) continue;
            foreach (var p in part.Parts) queue.Enqueue(p.OpenXmlPart);
            if (!part.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase)) continue;
            using var input = part.GetStream(); result.Add(XDocument.Load(input));
        }
        return result.ToArray();
    }
}
