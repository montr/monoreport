using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Cli.Tests;
using MonoReport.Docx;
using Xunit;

namespace MonoReport.Tests.Docx;

public sealed partial class DocxIdentityTests
{
    [Theory]
    [InlineData("{Na", "me}", "")]
    [InlineData("{Na", "", "me}")]
    [InlineData("{Na", "", "")]
    [InlineData("", "{Na", "me}")]
    public void Content_control_entry_exit_and_empty_controls_never_join_expression_fragments(string before,
        string hidden, string after)
    {
        var control = Control(hidden.Length == 0 ? [] : [DocxFixture.Run(hidden)]);
        var template =
            DocxFixture.Create(DocxFixture.P(DocxFixture.Run(before), control, Control(), DocxFixture.Run(after)));
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{\"Name\":\"CROSSED\"}", "template.unmatched_brace");
        Assert.DoesNotContain(analysis.Occurrences, e => e.Expression == "{Name}");
    }

    [Theory]
    [InlineData("inline")]
    [InlineData("block")]
    [InlineData("row")]
    [InlineData("cell")]
    public void Controls_are_opaque_and_preserved_when_an_outer_operation_moves_them_intact(string placement)
    {
        var inner = Control(DocxFixture.Run("{Hidden}"));
        var control = placement switch
        {
            "inline" => Control(DocxFixture.Run("{For:Hidden}"), inner, Control(), Control(DocxFixture.Run("{/For}"))),
            "block" => Control(DocxFixture.P(DocxFixture.Run("{For:Hidden}")),
                DocxFixture.P(Control(DocxFixture.Run("{Name}"))),
                DocxFixture.P(DocxFixture.Run("{/For}"))),
            "row" => Control(DocxFixture.Row(DocxFixture.Run("{For:Hidden}")),
                DocxFixture.Row(DocxFixture.Run("{Name}")),
                DocxFixture.Row(DocxFixture.Run("{/For}"))),
            _ => Control(new XElement(W + "tc", DocxFixture.P(DocxFixture.Run("{For:Hidden}")),
                DocxFixture.P(DocxFixture.Run("{Name}")),
                DocxFixture.P(DocxFixture.Run("{/For}"))))
        };
        control.Element(W + "sdtPr")!.Add(new XElement(W + "dataBinding", new XAttribute(W + "xpath", "/root/value"),
                new XAttribute(W + "storeItemID", "{01234567-89AB-CDEF-0123-456789ABCDEF}")),
            new XElement(W + "lock", new XAttribute(W + "val", "sdtLocked")));
        var body = placement switch
        {
            "inline" => DocxFixture.P(control, DocxFixture.Run("{Name}")),
            "row" => DocxFixture.Table(control),
            "cell" => DocxFixture.Table(new XElement(W + "tr", control)),
            _ => control
        };
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Items}")), body,
            DocxFixture.P(DocxFixture.Run("{/For}")),
            DocxFixture.P(DocxFixture.Run("{Visible}")));
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        using var input = new MemoryStream(template);
        var analysis = Analyzer().Analyze(input);
        Assert.True(analysis.Valid);
        Assert.DoesNotContain(analysis.Occurrences,
            e => e.Expression.Contains("Hidden") || (placement != "inline" && e.Expression == "{Name}"));
        var output = DocxFixture.Generate(template,
            "{\"Items\":[{\"Name\":\"A\"},{\"Name\":\"B\"}],\"Visible\":\"OK\"}",
            validationVersion: FileFormatVersions.Office2010).Document;
        var xml = DocxFixture.Read(output);
        var copies = xml.Descendants(W + "sdt").Where(e => !e.Ancestors(W + "sdt").Any()).ToArray();
        Assert.Equal(2, copies.Length);
        Assert.All(copies, copy => Assert.True(XNode.DeepEquals(control, copy)));
        Assert.EndsWith("OK", DocxFixture.Text(xml));
        if (placement == "inline") Assert.Contains("A", DocxFixture.Text(xml));
    }

    [Fact]
    public void Control_relationships_and_hidden_revision_preflight_remain_independent_of_expression_opacity()
    {
        var template = Prepare(
            DocxFixture.Create(DocxFixture.P(Control(DocxFixture.Run("{Hidden}")), DocxFixture.Run("{Visible}"))),
            document =>
            {
                var main = document.MainDocumentPart!;
                var relationship = main.AddHyperlinkRelationship(new Uri("https://example.org/control"), true);
                using var input = main.GetStream();
                var xml = XDocument.Load(input);
                input.Close();
                xml.Descendants(W + "sdtContent").Single().Add(new XElement(W + "hyperlink",
                    new XAttribute(WordXml.RelationshipsNamespace + "id", relationship.Id), DocxFixture.Run("link")));
                Save(main, xml.Root!);
            });
        var output = DocxFixture
            .Generate(template, "{\"Visible\":\"OK\"}", validationVersion: FileFormatVersions.Office2010)
            .Document;
        Assert.True(XNode.DeepEquals(DocxFixture.Read(template).Descendants(W + "sdt").Single(),
            DocxFixture.Read(output).Descendants(W + "sdt").Single()));
        using (var document = WordprocessingDocument.Open(new MemoryStream(output), false))
        {
            Assert.Equal("https://example.org/control",
                Assert.Single(document.MainDocumentPart!.HyperlinkRelationships).Uri.ToString());
        }

        var revised = DocxFixture.Create(DocxFixture.P(Control(new XElement(W + "ins", new XAttribute(W + "id", 1),
            new XAttribute(W + "author", "A"), DocxFixture.Run("{Hidden}")))));
        DocxFixture.AssertValid(revised, FileFormatVersions.Office2010);
        Reject(revised, "{}", "template.tracked_revision");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Nested_local_comment_copies_are_rejected_before_identity_allocation(bool point)
    {
        var source = Fixture("comment-with-image", ".template.docx");
        var template = Prepare(source, document =>
        {
            var main = document.MainDocumentPart!;
            using var input = main.GetStream();
            var xml = XDocument.Load(input);
            input.Close();
            var body = xml.Descendants(W + "body").Single();
            var attachment = new XElement(body.Elements(W + "p").ElementAt(1));
            attachment.Descendants(W + "t").First().Value = "{Name}";
            if (point)
            {
                attachment.Descendants(W + "commentRangeStart").Remove();
                attachment.Descendants(W + "commentRangeEnd").Remove();
            }

            body.ReplaceNodes(DocxFixture.P(DocxFixture.Run("{For:Items}")),
                DocxFixture.P(DocxFixture.Run("{If:Show}")),
                DocxFixture.P(DocxFixture.Run("{For:Children}")), attachment, DocxFixture.P(DocxFixture.Run("{/For}")),
                DocxFixture.P(DocxFixture.Run("{/If}")), DocxFixture.P(DocxFixture.Run("{/For}")));
            Save(main, xml.Root!);
            var commentsPart = main.WordprocessingCommentsPart!;
            using var commentsInput = commentsPart.GetStream();
            var comments = XDocument.Load(commentsInput);
            commentsInput.Close();
            var definition = comments.Root!.Elements().Single();
            var drawingRun = new XElement(definition.Descendants(W + "drawing").First().Parent!);
            drawingRun.Descendants(Wp + "docPr").Single().SetAttributeValue("id", 2);
            definition.Add(DocxFixture.P(Start("bookmark", 4), Start("permission", 5), DocxFixture.Run("opaque {Text}"),
                End("permission", 5), End("bookmark", 4), Link("B4", "local"), drawingRun));
            Save(commentsPart, comments.Root);
            // Reserve identities in untouched main/header/footer/glossary XML as well as definitions.
            var drawing = new XElement(definition.Descendants(W + "drawing").First());
            var imagePart = commentsPart.GetPartById((string) drawing
                .Descendants(XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main") + "blip").Single()
                .Attribute(WordXml.RelationshipsNamespace + "embed")!);
            var header = main.AddNewPart<HeaderPart>();
            var footer = main.AddNewPart<FooterPart>();
            var glossary = main.AddNewPart<GlossaryDocumentPart>();
            foreach (var (part, id) in new (OpenXmlPart, uint)[] {(main, 3), (header, 4), (footer, 5), (glossary, 6)})
            {
                var copy = new XElement(drawing);
                copy.Descendants(Wp + "docPr").Single().SetAttributeValue("id", id);
                var relation = part.AddPart(imagePart);
                copy.Descendants(XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main") + "blip")
                    .Single()
                    .SetAttributeValue(WordXml.RelationshipsNamespace + "embed", part.GetIdOfPart(relation));
                var paragraph = DocxFixture.P(new XElement(W + "r", copy));
                if (part == main)
                {
                    xml.Descendants(W + "body").Single().AddFirst(paragraph);
                    Save(main, xml.Root!);
                }
                else if (part == header)
                {
                    Save(part, new XElement(W + "hdr", paragraph));
                }
                else if (part == footer)
                {
                    Save(part, new XElement(W + "ftr", paragraph));
                }
                else
                {
                    Save(part, new XElement(W + "glossaryDocument", new XElement(W + "docParts",
                        new XElement(W + "docPart",
                            new XElement(W + "docPartPr", new XElement(W + "name", new XAttribute(W + "val", "entry"))),
                            new XElement(W + "docPartBody", paragraph)))));
                }
            }
        });
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        foreach (var json in new[]
                 {
                     "{\"Items\":[]}",
                     """{"Items":[{"Show":true,"Children":[{"Name":"A"},{"Name":"B"}]},{"Show":false,"Children":[{"Name":"NO"}]}]}"""
                 })
            Reject(template, json, "template.invalid_structure");
    }

    [Fact]
    public async Task Reused_engine_isolates_Raw_cleanup_drawing_and_permission_allocations()
    {
        var template = DocxCompatibilityTests.SupportedCopyTemplate();
        var engine = DocxFixture.CreateGenerator();

        byte[] Run()
        {
            using var input = new MemoryStream(template);
            using var data = JsonDocument.Parse("{\"Items\":[\"A\",\"B\"]}");
            var output = engine.Generate(input, data).Document;
            DocxFixture.AssertValid(output, FileFormatVersions.Office2010);
            using var generated = WordprocessingDocument.Open(new MemoryStream(output), false);
            var ids = PackageXml(generated).SelectMany(xml => xml.Descendants(Wp + "docPr"))
                .Select(element => (string) element.Attribute("id")!).ToArray();
            Assert.Equal(ids.Length, ids.Distinct().Count());
            var main = DocxFixture.Read(output);
            var permissions = main.Descendants(W + "permStart").Select(element => (string) element.Attribute(W + "id")!)
                .ToArray();
            Assert.Equal(2, permissions.Distinct().Count());
            Assert.Equal(permissions,
                main.Descendants(W + "permEnd").Select(element => (string) element.Attribute(W + "id")!));
            Assert.Single(generated.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements());
            Assert.Equal(template, input.ToArray());
            Assert.True(input.CanRead);
            return output;
        }

        Run();
        Run();
        await Task.WhenAll(Enumerable.Range(0, 8).Select(_ => Task.Run(Run)));
        Run();
    }

    private static XElement Control(params object[] content)
    {
        return new XElement(W + "sdt", new XElement(W + "sdtPr",
                new XElement(W + "alias", new XAttribute(W + "val", "opaque")),
                new XElement(W + "tag", new XAttribute(W + "val", "keep"))),
            new XElement(W + "sdtContent", content));
    }

    private static XDocument[] PackageXml(WordprocessingDocument document)
    {
        var result = new List<XDocument>();
        var queue = new Queue<OpenXmlPart>(document.Parts.Select(p => p.OpenXmlPart));
        var seen = new HashSet<OpenXmlPart>();
        while (queue.TryDequeue(out var part))
        {
            if (!seen.Add(part)) continue;
            foreach (var p in part.Parts) queue.Enqueue(p.OpenXmlPart);
            if (!part.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase)) continue;
            using var input = part.GetStream();
            result.Add(XDocument.Load(input));
        }

        return result.ToArray();
    }
}