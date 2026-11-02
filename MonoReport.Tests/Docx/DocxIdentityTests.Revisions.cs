using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

using MonoReport.Cli.Tests;
using System;
using System.IO;
using System.Linq;

namespace MonoReport.Tests.Docx;

public sealed partial class DocxIdentityTests
{
    [Theory]
    [InlineData("ins")]
    [InlineData("del")]
    [InlineData("moveFrom")]
    [InlineData("moveTo")]
    [InlineData("cellIns")]
    [InlineData("cellDel")]
    [InlineData("cellMerge")]
    [InlineData("rPrChange")]
    [InlineData("pPrChange")]
    [InlineData("sectPrChange")]
    [InlineData("tblPrChange")]
    [InlineData("tblGridChange")]
    [InlineData("trPrChange")]
    [InlineData("tcPrChange")]
    [InlineData("tblPrExChange")]
    [InlineData("numberingChange")]
    [InlineData("delText")]
    [InlineData("delInstrText")]
    [InlineData("conflictIns")]
    [InlineData("conflictDel")]
    public void Revision_content_cell_and_property_families_reject_valid_sources(string name)
    {
        var template = RevisionFixture(name);
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{}", "template.tracked_revision");
        Assert.Contains(analysis.Errors,
            e => e.Location!.EndsWith("tracked revision " + name, StringComparison.Ordinal));
        Assert.All(analysis.Errors, e => Assert.Equal("template.tracked_revision", e.Code));
        if (name.StartsWith("cell") || name == "tcPrChange")
            Assert.Contains("table 1, row 1, cell 1", analysis.Errors[0].Location);
        if (name is "trPrChange" or "tblPrExChange") Assert.Contains("table 1, row 1", analysis.Errors[0].Location);
        if (name == "sectPrChange") Assert.Contains("document properties", analysis.Errors[0].Location);
    }

    [Theory]
    [InlineData("moveFromRange")]
    [InlineData("moveToRange")]
    [InlineData("customXmlInsRange")]
    [InlineData("customXmlDelRange")]
    [InlineData("customXmlMoveFromRange")]
    [InlineData("customXmlMoveToRange")]
    [InlineData("customXmlConflictInsRange")]
    [InlineData("customXmlConflictDelRange")]
    public void Both_ends_of_each_revision_range_are_rejected_in_source_order(string family)
    {
        var ns = family.Contains("Conflict")
            ? XNamespace.Get("http://schemas.microsoft.com/office/word/2010/wordml")
            : W;
        var start = Revision(ns + (family + "Start"));
        if (family.StartsWith("move")) start.Add(new XAttribute(W + "name", "Move"));
        var end = new XElement(ns + (family + "End"), new XAttribute(W + "id", 1));
        var template = DocxFixture.Create(DocxFixture.P(start, DocxFixture.Run("{Name}"), end));
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{}", "template.tracked_revision");
        Assert.Equal(new[] {family + "Start", family + "End"},
            analysis.Errors.Select(e => e.Location!.Split("tracked revision ")[1]));
        Assert.All(analysis.Errors, e => Assert.Contains("paragraph 1", e.Location));
        // An isolated end is intentionally not a valid paired revision range; detection must still work.
        var isolated = Reject(DocxFixture.Create(DocxFixture.P(new XElement(end))), "{}", "template.tracked_revision");
        Assert.EndsWith("tracked revision " + family + "End", Assert.Single(isolated.Errors).Location);
    }

    [Fact]
    public void Revision_ranges_in_all_stories_keep_stable_part_and_paragraph_order_and_error_cap()
    {
        XElement Paragraph()
        {
            return DocxFixture.P(Revision(W + "moveToRangeStart", new XAttribute(W + "name", "Moved")),
                DocxFixture.Run("x"), new XElement(W + "moveToRangeEnd", new XAttribute(W + "id", 1)));
        }

        var template = Prepare(DocxFixture.Create(Paragraph()), document =>
        {
            var main = document.MainDocumentPart!;
            Save(main.AddNewPart<HeaderPart>(), new XElement(W + "hdr", Paragraph()));
            Save(main.AddNewPart<FooterPart>(), new XElement(W + "ftr", Paragraph()));
            Save(main.AddNewPart<FootnotesPart>(),
                new XElement(W + "footnotes", new XElement(W + "footnote", new XAttribute(W + "id", 1), Paragraph())));
            Save(main.AddNewPart<EndnotesPart>(),
                new XElement(W + "endnotes", new XElement(W + "endnote", new XAttribute(W + "id", 1), Paragraph())));
        });
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{}", "template.tracked_revision");
        Assert.Equal(
            new[]
            {
                "/word/document.xml", "/word/header1.xml", "/word/footer1.xml", "/word/footnotes.xml",
                "/word/endnotes.xml"
            },
            analysis.Errors.Select(e => e.Part).Distinct());
        Assert.Equal(10, analysis.Errors.Count);
        var many = DocxFixture.Create(Enumerable.Range(1, 105)
            .Select(id => DocxFixture.P(Revision(W + "ins", DocxFixture.Run("text"), id)))
            .ToArray());
        DocxFixture.AssertValid(many, FileFormatVersions.Office2010);
        Assert.Equal(105, Reject(many, "{}", "template.tracked_revision").Errors.Count);
    }

    [Fact]
    public void Deleted_payload_detection_does_not_depend_on_a_wrapper()
    {
        // Deliberately nonconformant isolated payload probes supplement valid deletion wrappers above.
        foreach (var name in new[] {"delText", "delInstrText"})
            Assert.Single(Reject(
                DocxFixture.Create(DocxFixture.P(new XElement(W + "r", new XElement(W + name, "old")))), "{}",
                "template.tracked_revision").Errors);
    }

    [Fact]
    public void Nonrevision_metadata_and_track_future_edits_are_not_tracked_content()
    {
        var template = Prepare(WithComments(DocxFixture.Create(DocxFixture.P(
                    new XAttribute(W + "rsidR", "00AABBCC"), Start("bookmark", 0), Start("permission", 1),
                    Start("comment", 2),
                    new XElement(W + "proofErr", new XAttribute(W + "type", "spellStart")), DocxFixture.Run("{Name}"),
                    new XElement(W + "proofErr", new XAttribute(W + "type", "spellEnd")), End("comment", 2),
                    End("permission", 1), End("bookmark", 0)),
                DocxFixture.Table(new XElement(W + "tr",
                    new XElement(W + "tc",
                        new XElement(W + "tcPr", new XElement(W + "vMerge", new XAttribute(W + "val", "restart"))),
                        DocxFixture.P(DocxFixture.Run("merged")))))), 2),
            document => Save(document.MainDocumentPart!.AddNewPart<DocumentSettingsPart>(),
                new XElement(W + "settings", new XElement(W + "trackRevisions"))));
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        using var input = new MemoryStream(template);
        Assert.True(Analyzer().Analyze(input).Valid);
        Assert.Equal("OKmerged",
            DocxFixture.Text(DocxFixture.Read(DocxFixture
                .Generate(template, "{\"Name\":\"OK\"}", validationVersion: FileFormatVersions.Office2010).Document)));
    }

    [Theory]
    [InlineData("paragraph-mark")]
    [InlineData("row")]
    [InlineData("custom-xml")]
    public void Revision_detection_covers_property_marks_and_custom_xml_wrappers(string placement)
    {
        var template = placement switch
        {
            "paragraph-mark" => DocxFixture.Create(
                DocxFixture.P(new XElement(W + "pPr", new XElement(W + "rPr", Revision(W + "ins"))),
                    DocxFixture.Run("x"))),
            "row" => DocxFixture.Create(DocxFixture.Table(new XElement(W + "tr",
                new XElement(W + "trPr", Revision(W + "del")),
                new XElement(W + "tc", DocxFixture.P(DocxFixture.Run("x")))))),
            _ => DocxFixture.Create(new XElement(W + "customXml", new XAttribute(W + "element", "custom"),
                DocxFixture.P(
                    Revision(W + "customXmlInsRangeStart"), DocxFixture.Run("x"),
                    new XElement(W + "customXmlInsRangeEnd", new XAttribute(W + "id", 1)))))
        };
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{}", "template.tracked_revision");
        Assert.Equal(placement == "custom-xml" ? 2 : 1, analysis.Errors.Count);
    }

    [Fact]
    public void Revision_like_content_inside_opaque_pictures_retains_the_existing_exclusion()
    {
        XNamespace v = "urn:schemas-microsoft-com:vml";
        var picture = new XElement(W + "pict", new XElement(v + "shape", new XElement(v + "textbox",
            new XElement(W + "txbxContent", DocxFixture.P(Revision(W + "ins", DocxFixture.Run("{Hidden}")))))));
        var template = DocxFixture.Create(DocxFixture.P(new XElement(W + "r", picture), DocxFixture.Run("{Name}")));
        DocxFixture.AssertValid(template, FileFormatVersions.Office2010);
        using var input = new MemoryStream(template);
        Assert.True(Analyzer().Analyze(input).Valid);
        var output = DocxFixture
            .Generate(template, "{\"Name\":\"OK\"}", validationVersion: FileFormatVersions.Office2010)
            .Document;
        Assert.True(XNode.DeepEquals(DocxFixture.Read(template).Descendants(W + "pict").Single(),
            DocxFixture.Read(output).Descendants(W + "pict").Single()));
        Assert.EndsWith("OK", DocxFixture.Text(DocxFixture.Read(output)));
    }

    private static XElement Revision(XName name, object? content = null, int id = 1)
    {
        return new XElement(name,
            new XAttribute(W + "id", id), new XAttribute(W + "author", "Reviewer"),
            new XAttribute(W + "date", "2026-10-18T00:00:00Z"), content);
    }

    private static byte[] RevisionFixture(string name)
    {
        XNamespace w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        if (name is "ins" or "del" or "moveFrom" or "moveTo")
            return DocxFixture.Create(DocxFixture.P(Revision(W + name, DocxFixture.Run("x"))));
        if (name is "conflictIns" or "conflictDel")
            return DocxFixture.Create(DocxFixture.P(Revision(w14 + name, DocxFixture.Run("x"))));
        if (name is "delText" or "delInstrText")
            return DocxFixture.Create(DocxFixture.P(Revision(W + "del",
                new XElement(W + "r", new XElement(W + name, "old")))));
        if (name == "rPrChange")
            return DocxFixture.Create(DocxFixture.P(new XElement(W + "r",
                new XElement(W + "rPr", Revision(W + name, new XElement(W + "rPr"))), new XElement(W + "t", "x"))));
        if (name == "pPrChange")
            return DocxFixture.Create(DocxFixture.P(
                new XElement(W + "pPr", Revision(W + name, new XElement(W + "pPr"))),
                DocxFixture.Run("x")));
        if (name == "numberingChange")
            return DocxFixture.Create(DocxFixture.P(
                new XElement(W + "pPr",
                    new XElement(W + "numPr", Revision(W + name, new XAttribute(W + "original", "1")))),
                DocxFixture.Run("x")));
        if (name == "sectPrChange")
            return DocxFixture.Create(DocxFixture.P(DocxFixture.Run("x")),
                new XElement(W + "sectPr", Revision(W + name, new XElement(W + "sectPr"))));
        var table = DocxFixture.Table(DocxFixture.Row(DocxFixture.Run("x")));
        if (name == "tblPrChange") table.Element(W + "tblPr")!.Add(Revision(W + name, new XElement(W + "tblPr")));
        else if (name == "tblGridChange")
            table.Element(W + "tblGrid")!.Add(new XElement(W + name, new XAttribute(W + "id", 1),
                new XElement(W + "tblGrid")));
        else if (name == "tblPrExChange")
            table.Element(W + "tr")!.AddFirst(new XElement(W + "tblPrEx",
                Revision(W + name, new XElement(W + "tblPrEx"))));
        else if (name == "trPrChange")
            table.Element(W + "tr")!.AddFirst(new XElement(W + "trPr", Revision(W + name, new XElement(W + "trPr"))));
        else
            table.Descendants(W + "tc").First().AddFirst(new XElement(W + "tcPr",
                Revision(W + name, name == "tcPrChange" ? new XElement(W + "tcPr") : null)));
        return DocxFixture.Create(table);
    }
}