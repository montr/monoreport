using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Cli.Tests;
using MonoReport.Core.Templates;
using MonoReport.Docx;
using Xunit;

namespace MonoReport.Tests.Docx;

public sealed class DocxCompatibilityTests
{
    private static readonly XNamespace W = WordXml.WordprocessingNamespace;

    private static XElement P(string text)
    {
        return DocxFixture.P(DocxFixture.Run(text));
    }

    private static XElement Start(string family, int id)
    {
        return DocxIdentityTests.Start(family, id);
    }

    private static XElement End(string family, int id)
    {
        return DocxIdentityTests.End(family, id);
    }

    private static XDocument MergeCell(XElement cell, string json)
    {
        var source = DocxFixture.Create(DocxFixture.Table(new XElement(W + "tr", cell)));
        using var input = new MemoryStream(source);
        var analysis = DocxIdentityTests.Analyzer().Analyze(input);
        Assert.True(analysis.Valid, string.Join("; ", analysis.Errors.Select(error => error.Message)));
        return DocxFixture.Read(DocxFixture.Generate(source, json, validationVersion: FileFormatVersions.Office2010)
            .Document);
    }

    private static void AssertCompatibleCells(XContainer xml)
    {
        foreach (var cell in xml.Descendants(W + "tc"))
        {
            var blocks = cell.Elements().Where(element => element.Name != W + "tcPr").ToArray();
            Assert.NotEmpty(blocks);
            Assert.Equal(W + "p", blocks[^1].Name);
        }
    }

    internal static byte[] Comments(byte[] source, params string[] ids)
    {
        return DocxIdentityTests.Prepare(source, document =>
        {
            var main = document.MainDocumentPart!;
            using var stream = main.GetStream();
            var xml = XDocument.Load(stream);
            stream.Close();
            foreach (var holder in xml.Descendants("holder").ToArray()) holder.ReplaceWith(holder.Elements().ToArray());
            DocxIdentityTests.Save(main, xml.Root!);
            DocxIdentityTests.Save(main.AddNewPart<WordprocessingCommentsPart>(), new XElement(W + "comments",
                ids.Select(id => new XElement(W + "comment", new XAttribute(W + "id", id),
                    new XAttribute(W + "author", "Author"),
                    P("definition " + id)))));
        });
    }

    [Theory]
    [InlineData("bookmark")]
    [InlineData("comment")]
    public void Local_and_marker_annotations_reject_even_in_inactive_or_zero_iteration_operations(string family)
    {
        foreach (var marker in new[] {false, true})
        {
            var source = Comments(DocxFixture.Create(
                marker
                    ? DocxFixture.P(Start(family, 1), DocxFixture.Run("{For:Items}"), End(family, 1))
                    : P("{For:Items}"),
                P("{Raw}"),
                marker ? P("literal") : DocxFixture.P(Start(family, 1), DocxFixture.Run("literal"), End(family, 1)),
                P("{/Raw}"), P("{/For}")), "1");
            foreach (var data in new[] {"{}", "{\"Items\":[]}", "{\"Items\":[1]}", "{\"Items\":[1,2]}"})
                DocxIdentityTests.Reject(source, data, "template.invalid_structure");
        }
    }

    [Theory]
    [InlineData("bookmark")]
    [InlineData("comment")]
    public void Complete_enclosing_ranges_are_preserved_with_unchanged_identities(string family)
    {
        var source = Comments(DocxFixture.Create(DocxFixture.P(Start(family, 4), DocxFixture.Run("before")),
            P("{For:Items}"), P("{.}"), P("{/For}"), DocxFixture.P(DocxFixture.Run("after"), End(family, 4))), "4");
        using var input = new MemoryStream(source);
        Assert.True(DocxIdentityTests.Analyzer().Analyze(input).Valid);
        var output = DocxFixture
            .Generate(source, "{\"Items\":[\"A\",\"B\"]}", validationVersion: FileFormatVersions.Office2010)
            .Document;
        Assert.Equal("beforeABafter", DocxFixture.Text(DocxFixture.Read(output)));
        var name = family == "bookmark" ? "bookmarkStart" : "commentRangeStart";
        var original = DocxFixture.Read(source).Descendants(W + name).Single();
        Assert.True(XNode.DeepEquals(original, DocxFixture.Read(output).Descendants(W + name).Single()));
    }

    [Theory]
    [InlineData("Stable")]
    [InlineData("Missing")]
    public void Internal_links_inside_operations_reject_even_with_stable_or_missing_targets(string target)
    {
        var source = DocxFixture.Create(DocxFixture.P(DocxIdentityTests.Bookmark(0, "Stable", "stable")),
            P("{If:Show}"), DocxFixture.P(DocxIdentityTests.Link(target, "link")), P("{/If}"));
        foreach (var data in new[] {"{\"Show\":false}", "{\"Show\":true}"})
            DocxIdentityTests.Reject(source, data, "template.invalid_structure");
    }

    [Theory]
    [InlineData("bookmark")]
    [InlineData("comment")]
    [InlineData("permission")]
    public void Raw_cannot_split_deleted_and_surviving_range_members(string family)
    {
        var source = Comments(DocxFixture.Create(DocxFixture.P(Start(family, 2), DocxFixture.Run("{Raw}")),
            DocxFixture.P(DocxFixture.Run("body"), End(family, 2)), P("{/Raw}")), "2");
        DocxIdentityTests.Reject(source, "{}", "template.invalid_structure");
    }

    [Fact]
    public void First_eligible_close_with_unsafe_deletion_is_not_skipped()
    {
        var source = Comments(DocxFixture.Create(P("{Raw}"),
            DocxFixture.P(Start("bookmark", 0), DocxFixture.Run("body")),
            DocxFixture.P(End("bookmark", 0), DocxFixture.Run("{/Raw}")), P("{/Raw}")), "0");
        var result = DocxIdentityTests.Reject(source, "{}", "template.invalid_structure");
        Assert.Contains(result.Errors, error => error.Message.Contains("bookmark"));
        Assert.Contains(result.Errors, error => error.Field == "{/Raw}");
    }

    [Fact]
    public void Raw_section_boundary_deletion_rejects()
    {
        var section = new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "sectPr")),
            DocxFixture.Run("{Raw}"));
        var source = DocxFixture.Create(section, P("body"), P("{/Raw}"));
        DocxFixture.AssertValid(source);
        var analysis = DocxIdentityTests.Reject(source, "{}", "template.invalid_structure");
        var error = Assert.Single(analysis.Errors);
        Assert.Contains("Removing Raw delimiters would remove or damage a section boundary.", error.Message);
        Assert.Contains("Move the delimiters to ordinary paragraphs/rows outside that structure in Word",
            error.Message);
        Assert.Contains("preserve section boundaries", error.Message);
        Assert.Equal("{Raw}", error.Field);
        Assert.Equal("/word/document.xml", error.Part);
        Assert.Contains("paragraph 1", error.Location);

        using var input = new MemoryStream(source);
        using var data = JsonDocument.Parse("{}");
        GenerationResult? result = null;
        Assert.Throws<TemplateException>(() => result = DocxFixture.CreateGenerator().Generate(input, data));
        Assert.Null(result);
        Assert.Equal(source, input.ToArray());
    }

    [Theory]
    [InlineData("Raw", false)]
    [InlineData("If:Show", false)]
    [InlineData("Raw", true)]
    public void Surviving_cell_bookmark_markers_do_not_hide_empty_or_table_ending_blocks(string control,
        bool nestedTable)
    {
        var properties = new XElement(W + "tcPr", new XElement(W + "tcW", new XAttribute(W + "w", 2400)));
        var start = Start("bookmark", 7);
        var end = End("bookmark", 7);
        var table = DocxFixture.Table(DocxFixture.Row(DocxFixture.Run("literal {Hidden}")));
        var source = DocxFixture.Create(DocxFixture.Table(new XElement(W + "tr", new XElement(W + "tc",
            properties, start, P("{" + control + "}"), nestedTable ? table : null,
            P(control == "Raw" ? "{/Raw}" : "{/If}"), end))), P("unrelated"));
        DocxFixture.AssertValid(source, FileFormatVersions.Office2010);
        using var input = new MemoryStream(source);
        var analysis = DocxIdentityTests.Analyzer().Analyze(input);
        Assert.True(analysis.Valid, string.Join("; ", analysis.Errors.Select(error => error.Message)));
        Assert.DoesNotContain(analysis.Requirements, requirement => requirement.Name == "Hidden");

        var output = DocxFixture.Read(DocxFixture.Generate(source, "{\"Show\":false}",
            validationVersion: FileFormatVersions.Office2010).Document);
        var cell = output.Descendants(W + "tc").First();
        var children = cell.Elements().ToArray();
        Assert.Equal(nestedTable
                ? new[] {W + "tcPr", W + "bookmarkStart", W + "tbl", W + "bookmarkEnd", W + "p"}
                : new[] {W + "tcPr", W + "bookmarkStart", W + "bookmarkEnd", W + "p"},
            children.Select(element => element.Name));
        Assert.True(XNode.DeepEquals(properties, children[0]));
        Assert.True(XNode.DeepEquals(start, children[1]));
        Assert.True(XNode.DeepEquals(end, children[^2]));
        var paragraph = Assert.Single(cell.Elements(W + "p"));
        Assert.Empty(paragraph.Attributes());
        Assert.Empty(paragraph.Nodes());
        if (nestedTable) Assert.True(XNode.DeepEquals(table, children[2]));
        Assert.True(XNode.DeepEquals(P("unrelated"), output.Root!.Element(W + "body")!.Elements().Last()));
    }

    [Fact]
    public void Surviving_cell_bookmark_after_a_formatted_paragraph_does_not_add_another_paragraph()
    {
        var retained = DocxFixture.P(new XElement(W + "pPr", new XElement(W + "keepNext")),
            DocxFixture.Run("retained"));
        var expected = new XElement(W + "tc", Start("bookmark", 7), retained, End("bookmark", 7));
        var output = MergeCell(new XElement(W + "tc", Start("bookmark", 7),
            P("{Raw}"), retained, P("{/Raw}"), End("bookmark", 7)), "{}");

        Assert.True(XNode.DeepEquals(expected, output.Descendants(W + "tc").Single()));
    }

    [Fact]
    public void Raw_only_cell_gets_one_empty_compatible_paragraph()
    {
        var properties = new XElement(W + "tcPr", new XElement(W + "tcW", new XAttribute(W + "w", 2400)));
        var output = MergeCell(new XElement(W + "tc", properties, P("{Raw}"), P("{/Raw}")), "{}");
        var cell = output.Descendants(W + "tc").Single();

        Assert.True(XNode.DeepEquals(properties, cell.Element(W + "tcPr")));
        Assert.Empty(cell.Elements(W + "p").Single().Elements());
        AssertCompatibleCells(output);
    }

    [Theory]
    [InlineData("{For:Items}", "{/For}", "{\"Items\":[]}")]
    [InlineData("{If:Show}", "{/If}", "{\"Show\":false}")]
    public void Empty_For_and_If_cell_results_get_one_compatible_paragraph(string open, string close, string json)
    {
        var output = MergeCell(new XElement(W + "tc", P(open), P(close)), json);

        var paragraph = Assert.Single(output.Descendants(W + "tc").Single().Elements(W + "p"));
        Assert.Empty(paragraph.Elements());
        AssertCompatibleCells(output);
    }

    [Fact]
    public void Mixed_For_If_and_Raw_that_remove_all_cell_content_leave_one_paragraph()
    {
        var output = MergeCell(new XElement(W + "tc", P("{For:Items}"), P("{If:Show}"),
            P("{Raw}"), P("{/Raw}"), P("{/If}"), P("{/For}")), "{\"Items\":[{\"Show\":true}]}");

        Assert.Empty(Assert.Single(output.Descendants(W + "tc").Single().Elements(W + "p")).Elements());
        AssertCompatibleCells(output);
    }

    [Fact]
    public void Cell_with_a_normal_trailing_paragraph_is_not_changed_or_extended()
    {
        var retained = DocxFixture.P(new XElement(W + "pPr", new XElement(W + "keepNext")),
            new XElement(W + "r", DocxFixture.Properties(), new XElement(W + "t", "retained")));
        var output = MergeCell(new XElement(W + "tc", P("{If:Show}"), P("discarded"),
            P("{/If}"), retained), "{\"Show\":false}");
        var paragraphs = output.Descendants(W + "tc").Single().Elements(W + "p").ToArray();

        Assert.Single(paragraphs);
        Assert.True(XNode.DeepEquals(retained, paragraphs[0]));
        AssertCompatibleCells(output);
    }

    [Fact]
    public void Recognized_revisions_remain_errors_in_Raw_and_inactive_branches()
    {
        foreach (var wrapping in new[] {"Raw", "If:Show", "For:Items"})
        {
            var close = wrapping == "Raw" ? "{/Raw}" : wrapping.StartsWith("If") ? "{/If}" : "{/For}";
            var revision = new XElement(W + "ins", new XAttribute(W + "id", 0), new XAttribute(W + "author", "A"),
                DocxFixture.Run("{Bad syntax}"));
            DocxIdentityTests.Reject(DocxFixture.Create(P("{" + wrapping + "}"), DocxFixture.P(revision), P(close)),
                "{\"Show\":false,\"Items\":[]}", "template.tracked_revision");
        }
    }

    [Theory]
    [InlineData("bookmark", false)]
    [InlineData("comment", false)]
    [InlineData("comment", true)]
    public void Raw_delimiter_annotations_are_deleted_without_dangling_survivors(string family, bool point)
    {
        var annotation = point
            ? new object[] {new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", "00")))}
            : [Start(family, 0), End(family, 0)];
        var source = Comments(DocxFixture.Create(DocxFixture.P(annotation, DocxFixture.Run("{Raw}")),
            DocxFixture.P(Start("bookmark", 4), Start("comment", 5), DocxFixture.Run("{Literal}"), End("bookmark", 4),
                End("comment", 5)),
            P("{/Raw}"), DocxFixture.P(DocxIdentityTests.Link("B4", "link"))), "0", "5", "99");
        using var input = new MemoryStream(source);
        Assert.True(DocxIdentityTests.Analyzer().Analyze(input).Valid);
        using var original = WordprocessingDocument.Open(new MemoryStream(source), false);
        var originalPart = original.MainDocumentPart!.WordprocessingCommentsPart!;
        var originalRelationshipId = original.MainDocumentPart.GetIdOfPart(originalPart);
        using var originalStream = originalPart.GetStream();
        var retainedDefinitions = XDocument.Load(originalStream).Root!.Elements()
            .Where(comment => (string?) comment.Attribute(W + "id") is "5" or "99")
            .Select(comment => comment.ToString(SaveOptions.DisableFormatting)).ToArray();
        var output = DocxFixture.Generate(source, "{}", validationVersion: FileFormatVersions.Office2010).Document;
        if (family == "comment" && !point)
        {
        }

        var xml = DocxFixture.Read(output);
        Assert.Equal("{Literal}link", DocxFixture.Text(xml));
        Assert.Single(xml.Descendants(W + "bookmarkStart"));
        Assert.Single(xml.Descendants(W + "commentReference"));
        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        var outputPart = document.MainDocumentPart!.WordprocessingCommentsPart!;
        Assert.Equal(originalRelationshipId, document.MainDocumentPart.GetIdOfPart(outputPart));
        using var stream = outputPart.GetStream();
        var outputComments = XDocument.Load(stream);
        Assert.Equal(retainedDefinitions, outputComments.Root!.Elements()
            .Where(comment => (string?) comment.Attribute(W + "id") is "5" or "99")
            .Select(comment => comment.ToString(SaveOptions.DisableFormatting)));
        var ids = outputComments.Root!.Elements().Select(e => (string?) e.Attribute(W + "id")).ToArray();
        Assert.Equal(family == "comment" ? new[] {"5", "99"} : new[] {"0", "5", "99"}, ids);
    }

    [Fact]
    public void Raw_renders_independently_in_supported_story_and_note_roots()
    {
        var source = DocxIdentityTests.Prepare(DocxFixture.Create(P("main")), document =>
        {
            var main = document.MainDocumentPart!;

            XElement Content(XName root, string literal)
            {
                return new XElement(root, P("{Raw}"), P(literal + " {Hidden}"),
                    P("{/Raw}"), P("{Name}"));
            }

            DocxIdentityTests.Save(main.AddNewPart<HeaderPart>(), Content(W + "hdr", "header"));
            DocxIdentityTests.Save(main.AddNewPart<FooterPart>(), Content(W + "ftr", "footer"));
            DocxIdentityTests.Save(main.AddNewPart<FootnotesPart>(), new XElement(W + "footnotes",
                new XElement(W + "footnote", new XAttribute(W + "id", 1),
                    Content(W + "customXml", "footnote").Elements())));
            DocxIdentityTests.Save(main.AddNewPart<EndnotesPart>(), new XElement(W + "endnotes",
                new XElement(W + "endnote", new XAttribute(W + "id", 1),
                    Content(W + "customXml", "endnote").Elements())));
        });
        using var input = new MemoryStream(source);
        var analysis = DocxIdentityTests.Analyzer().Analyze(input);
        Assert.True(analysis.Valid, string.Join("; ", analysis.Errors.Select(error => error.Message)));
        Assert.Equal("Name", Assert.Single(analysis.Requirements).Name);
        Assert.DoesNotContain(analysis.Occurrences, occurrence => occurrence.Expression == "{Hidden}");
        var output = DocxFixture.Generate(source, "{\"Name\":\"OK\"}", validationVersion: FileFormatVersions.Office2010)
            .Document;
        var stories = DocxFixture.ReadStories(output);
        Assert.Contains(stories.Values,
            story => DocxFixture.Text(story).Contains("header {Hidden}OK", StringComparison.Ordinal));
        Assert.Contains(stories.Values,
            story => DocxFixture.Text(story).Contains("footer {Hidden}OK", StringComparison.Ordinal));
        Assert.Contains(stories.Values,
            story => DocxFixture.Text(story).Contains("footnote {Hidden}OK", StringComparison.Ordinal));
        Assert.Contains(stories.Values,
            story => DocxFixture.Text(story).Contains("endnote {Hidden}OK", StringComparison.Ordinal));
        Assert.All(stories.Values,
            story => Assert.DoesNotContain("{Raw}", DocxFixture.Text(story), StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("bookmark")]
    [InlineData("comment")]
    public void Complete_ranges_spanning_standalone_Raw_are_preserved_unchanged(string family)
    {
        var source = Comments(DocxFixture.Create(DocxFixture.P(Start(family, 4), DocxFixture.Run("before")), P("{Raw}"),
            P("literal {Hidden}"), P("{/Raw}"), DocxFixture.P(DocxFixture.Run("after"), End(family, 4)),
            DocxFixture.P(DocxIdentityTests.Link("B4", "link"))), "4");
        var before = DocxFixture.Read(source);
        var output = DocxFixture.Read(DocxFixture
            .Generate(source, "{}", validationVersion: FileFormatVersions.Office2010).Document);
        Assert.Equal("beforeliteral {Hidden}afterlink", DocxFixture.Text(output));
        var startName = family == "bookmark" ? "bookmarkStart" : "commentRangeStart";
        var endName = family == "bookmark" ? "bookmarkEnd" : "commentRangeEnd";
        Assert.True(XNode.DeepEquals(before.Descendants(W + startName).Single(),
            output.Descendants(W + startName).Single()));
        Assert.True(
            XNode.DeepEquals(before.Descendants(W + endName).Single(), output.Descendants(W + endName).Single()));
        if (family == "bookmark")
            Assert.True(XNode.DeepEquals(before.Descendants(W + "hyperlink").Single(),
                output.Descendants(W + "hyperlink").Single()));
    }

    [Fact]
    public void Multiple_Raw_deletions_are_cumulative_and_preserve_unrelated_comment_definitions()
    {
        XElement Point(int id)
        {
            return new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", id)));
        }

        var source = Comments(DocxFixture.Create(DocxFixture.P(Point(1), DocxFixture.Run("{Raw}")), P("one"),
            P("{/Raw}"),
            P("between"), DocxFixture.P(Point(2), DocxFixture.Run("{Raw}")), P("two"), P("{/Raw}")), "1", "2", "9");
        var output = DocxFixture.Generate(source, "{}", validationVersion: FileFormatVersions.Office2010).Document;
        Assert.Equal("onebetweentwo", DocxFixture.Text(DocxFixture.Read(output)));
        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        using var stream = document.MainDocumentPart!.WordprocessingCommentsPart!.GetStream();
        Assert.Equal("9", (string?) Assert.Single(XDocument.Load(stream).Root!.Elements()).Attribute(W + "id"));
    }

    [Fact]
    public void Nested_tables_get_a_trailing_cell_paragraph_after_Raw_and_repeated_For_rendering()
    {
        var nested = DocxFixture.Table(DocxFixture.Row(DocxFixture.Run("nested")));
        foreach (var (cell, json, expectedTables) in new[]
                 {
                     (new XElement(W + "tc", P("{Raw}"), new XElement(nested), P("{/Raw}")), "{}", 1),
                     (new XElement(W + "tc", P("{For:Items}"), new XElement(nested), P("{/For}")),
                         "{\"Items\":[1,2]}", 2)
                 })
        {
            var output = MergeCell(cell, json);
            var outerCell = output.Descendants(W + "tc").First();
            var blocks = outerCell.Elements().ToArray();

            Assert.Equal(expectedTables, blocks.Count(element => element.Name == W + "tbl"));
            Assert.Equal(W + "p", blocks[^1].Name);
            Assert.Empty(blocks[^1].Elements());
            AssertCompatibleCells(output);
        }
    }

    [Fact]
    public void Repeated_nested_tables_and_nested_empty_blocks_leave_compatible_cells()
    {
        var nestedCell = new XElement(W + "tc", P("{If:Show}"), P("{Raw}"), P("{/Raw}"), P("{/If}"));
        var repeatedTable = DocxFixture.Table(new XElement(W + "tr", nestedCell));
        var source = DocxFixture.Create(P("{For:Items}"), repeatedTable, P("{/For}"));
        using var input = new MemoryStream(source);
        Assert.True(DocxIdentityTests.Analyzer().Analyze(input).Valid);

        var output = DocxFixture.Read(DocxFixture.Generate(source,
                "{\"Items\":[{\"Show\":true},{\"Show\":false}]}", validationVersion: FileFormatVersions.Office2010)
            .Document);

        Assert.Equal(2, output.Descendants(W + "tbl").Count());
        Assert.All(output.Descendants(W + "tc"), cell =>
        {
            var paragraph = Assert.Single(cell.Elements(W + "p"));
            Assert.Empty(paragraph.Elements());
        });
        AssertCompatibleCells(output);
    }

    [Fact]
    public void Stable_and_unresolved_internal_links_and_external_fragments_survive_standalone_Raw()
    {
        var source = DocxFixture.Create(P("{Raw}"), DocxFixture.P(DocxIdentityTests.Bookmark(3, "Target", "target"),
                DocxIdentityTests.Link("Target", "stable")), P("{/Raw}"),
            DocxFixture.P(DocxIdentityTests.Link("Missing", "unresolved")));
        source = DocxIdentityTests.Prepare(source, document =>
        {
            var main = document.MainDocumentPart!;
            var relationship = main.AddHyperlinkRelationship(new Uri("https://example.org/report.docx"), true);
            using var stream = main.GetStream();
            var xml = XDocument.Load(stream);
            stream.Close();
            xml.Descendants(W + "p").Last().Add(new XElement(W + "hyperlink",
                new XAttribute(WordXml.RelationshipsNamespace + "id", relationship.Id),
                new XAttribute(W + "anchor", "Fragment"), DocxFixture.Run("external")));
            DocxIdentityTests.Save(main, xml.Root!);
        });
        var output = DocxFixture.Generate(source, "{}", validationVersion: FileFormatVersions.Office2010).Document;
        var xml = DocxFixture.Read(output);
        Assert.Equal(new[] {"Target", "Missing", "Fragment"},
            xml.Descendants(W + "hyperlink").Select(link => (string?) link.Attribute(W + "anchor")));
        using var document = WordprocessingDocument.Open(new MemoryStream(output), false);
        var relationship = Assert.Single(document.MainDocumentPart!.HyperlinkRelationships);
        Assert.Equal("https://example.org/report.docx", relationship.Uri.ToString());
        Assert.True(relationship.IsExternal);
    }

    [Fact]
    public void Shared_definition_survival_counts_references_in_another_root_even_in_a_malformed_package()
    {
        XElement Reference()
        {
            return new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 0)));
        }

        var source =
            Comments(
                DocxFixture.Create(DocxFixture.P(Reference(), DocxFixture.Run("{Raw}")), P("literal"), P("{/Raw}")),
                "0");
        source = DocxIdentityTests.Prepare(source, document =>
        {
            var header = document.MainDocumentPart!.AddNewPart<HeaderPart>();
            DocxIdentityTests.Save(header,
                new XElement(W + "hdr", DocxFixture.P(Reference(), DocxFixture.Run("unchanged"))));
        });
        // Supplemental semantic probe: the header reference cannot resolve the main document's comments part.
        // Conformant sources instead have unique attachment IDs; ordinary definition cleanup is validated separately.
        using var input = new MemoryStream(source);
        using var data = JsonDocument.Parse("{}");
        Assert.True(DocxIdentityTests.Analyzer().Analyze(input).Valid);
        input.Position = 0;
        var output = DocxFixture.CreateGenerator().Generate(input, data).Document;
        using var doc = WordprocessingDocument.Open(new MemoryStream(output), false);
        Assert.Single(doc.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements());
        Assert.Empty(DocxFixture.Read(output).Descendants(W + "commentReference"));
        Assert.Single(DocxFixture.ReadStories(output).Values
            .SelectMany(xml => xml.Descendants(W + "commentReference")));
    }

    [Fact]
    public void Retained_auxiliary_link_to_removed_Raw_bookmark_rejects_without_remapping()
    {
        var source = DocxFixture.Create(
            DocxFixture.P(Start("bookmark", 0), DocxFixture.Run("{Raw}"), End("bookmark", 0)), P("literal"),
            P("{/Raw}"));
        source = DocxIdentityTests.Prepare(source, document => DocxIdentityTests.Save(
            document.MainDocumentPart!.AddNewPart<HeaderPart>(),
            new XElement(W + "hdr", DocxFixture.P(DocxIdentityTests.Link("B0", "outside")))));
        var result = DocxIdentityTests.Reject(source, "{}", "template.invalid_structure");
        Assert.Contains(result.Errors, error => error.Part!.Contains("header"));
    }

    [Fact]
    public void Raw_row_deletion_preserves_body_merges_but_rejects_a_deleted_required_restart()
    {
        XElement Cell(string text, string? merge = null)
        {
            return new XElement(W + "tc",
                merge is null
                    ? null
                    : new XElement(W + "tcPr", new XElement(W + "vMerge", new XAttribute(W + "val", merge))), P(text));
        }

        XElement Row(string text, string? merge = null)
        {
            return new XElement(W + "tr", Cell(text, merge));
        }

        var merged = new[] {Row("{Literal}", "restart"), Row("{{", "continue")};
        var source =
            DocxFixture.Create(DocxFixture.Table(new[] {Row("{Raw}")}.Concat(merged).Append(Row("{/Raw}")).ToArray()));
        var output = DocxFixture.Read(DocxFixture.Generate(source, "{}").Document);
        Assert.Equal(2, output.Descendants(W + "vMerge").Count());
        Assert.True(merged.Zip(output.Descendants(W + "tr")).All(pair => XNode.DeepEquals(pair.First, pair.Second)));
        var unsafeSource =
            DocxFixture.Create(DocxFixture.Table(Row("{Raw}", "restart"), Row("literal", "continue"), Row("{/Raw}")));
        DocxIdentityTests.Reject(unsafeSource, "{}", "template.invalid_structure");
        var harmless = DocxFixture.Create(DocxFixture.Table(Row("before", "restart"), Row("{Raw}", "continue"),
            Row("literal", "continue"), Row("{/Raw}")));
        Assert.Equal("beforeliteral",
            DocxFixture.Text(DocxFixture.Read(DocxFixture.Generate(harmless, "{}").Document)));
    }

    internal static byte[] SupportedCopyTemplate()
    {
        var source = DocxIdentityTests.Fixture("comment-with-image", ".template.docx");
        return DocxIdentityTests.Prepare(source, document =>
        {
            var main = document.MainDocumentPart!;
            var commentsPart = main.WordprocessingCommentsPart!;
            using var input = commentsPart.GetStream();
            var comments = XDocument.Load(input);
            input.Close();
            var drawing = new XElement(comments.Descendants(W + "drawing").Single());
            var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
            var image = commentsPart.GetPartById((string) drawing.Descendants(a + "blip").Single()
                .Attribute(WordXml.RelationshipsNamespace + "embed")!);
            main.AddPart(image);
            drawing.Descendants(a + "blip").Single()
                .SetAttributeValue(WordXml.RelationshipsNamespace + "embed", main.GetIdOfPart(image));
            drawing.Descendants(XName.Get("docPr",
                    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")).Single()
                .SetAttributeValue("id", 8);
            comments.Root!.Add(new XElement(W + "comment", new XAttribute(W + "id", 7),
                new XAttribute(W + "author", "A"), P("delete")));
            DocxIdentityTests.Save(commentsPart, comments.Root);
            var glossary = main.AddNewPart<GlossaryDocumentPart>();
            glossary.AddPart(image);
            var reserved = new XElement(drawing);
            reserved.Descendants(a + "blip").Single().SetAttributeValue(WordXml.RelationshipsNamespace + "embed",
                glossary.GetIdOfPart(image));
            reserved.Descendants(XName.Get("docPr",
                    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")).Single()
                .SetAttributeValue("id", 3);
            DocxIdentityTests.Save(glossary, new XElement(W + "glossaryDocument", new XElement(W + "docParts",
                new XElement(W + "docPart",
                    new XElement(W + "docPartPr", new XElement(W + "name", new XAttribute(W + "val", "entry"))),
                    new XElement(W + "docPartBody", DocxFixture.P(new XElement(W + "r", reserved)))))));
            DocxIdentityTests.Save(main, new XElement(W + "document", new XElement(W + "body",
                DocxFixture.P(new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 7))),
                    DocxFixture.Run("{Raw}")),
                P("{Malformed"), P("{/Raw}"),
                DocxFixture.P(new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 0)))),
                P("{For:Items}"), DocxFixture.P(Start("permission", 9)), P("{Raw}"),
                DocxFixture.P(new XElement(W + "r", drawing)), P("{{ {Literal}"), P("{/Raw}"),
                DocxFixture.P(DocxFixture.Run("{.}"), End("permission", 9)), P("{/For}"))));
        });
    }

    [Fact]
    public void Repeated_Raw_drawings_and_permissions_keep_package_identity_and_media_safety()
    {
        var source = SupportedCopyTemplate();
        var output = DocxFixture
            .Generate(source, "{\"Items\":[\"A\",\"B\"]}", validationVersion: FileFormatVersions.Office2010)
            .Document;
        var xml = DocxFixture.Read(output);
        Assert.Equal("{Malformed{{ {Literal}A{{ {Literal}B", DocxFixture.Text(xml));
        var starts = xml.Descendants(W + "permStart").Select(e => (string?) e.Attribute(W + "id")).ToArray();
        Assert.Equal(2, starts.Distinct().Count());
        Assert.Equal(starts, xml.Descendants(W + "permEnd").Select(e => (string?) e.Attribute(W + "id")));
        using var original = WordprocessingDocument.Open(new MemoryStream(source), false);
        using var generated = WordprocessingDocument.Open(new MemoryStream(output), false);
        var ids = new DocxStoryReader().Inspect(generated).PackageXml.SelectMany(item => item.Xml.Descendants(
                XName.Get("docPr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")))
            .Select(e => (string?) e.Attribute("id")).ToArray();
        Assert.Equal(ids.Length, ids.Distinct().Count());
        foreach (var part in original.MainDocumentPart!.ImageParts)
        {
            using var before = part.GetStream();
            using var after = generated.MainDocumentPart!.GetPartById(original.MainDocumentPart.GetIdOfPart(part))
                .GetStream();
            using var b = new MemoryStream();
            using var a = new MemoryStream();
            before.CopyTo(b);
            after.CopyTo(a);
            Assert.Equal(b.ToArray(), a.ToArray());
        }

        Assert.Single(generated.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements());
    }
}