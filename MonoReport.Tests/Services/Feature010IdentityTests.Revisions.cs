using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed partial class Feature010IdentityTests
{
    [Theory]
    [InlineData("ins")] [InlineData("del")] [InlineData("moveFrom")] [InlineData("moveTo")]
    [InlineData("cellIns")] [InlineData("cellDel")] [InlineData("cellMerge")]
    [InlineData("rPrChange")] [InlineData("pPrChange")] [InlineData("sectPrChange")]
    [InlineData("tblPrChange")] [InlineData("tblGridChange")] [InlineData("trPrChange")]
    [InlineData("tcPrChange")] [InlineData("tblPrExChange")] [InlineData("numberingChange")]
    [InlineData("delText")] [InlineData("delInstrText")]
    [InlineData("conflictIns")] [InlineData("conflictDel")]
    public void Revision_content_cell_and_property_families_reject_valid_sources(string name)
    {
        var template = RevisionFixture(name);
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{}", "template.tracked_revision");
        Assert.Contains(analysis.Errors, e => e.Location!.EndsWith("tracked revision " + name, StringComparison.Ordinal));
        Assert.All(analysis.Errors, e => Assert.Equal("template.tracked_revision", e.Code));
        if (name.StartsWith("cell") || name == "tcPrChange") Assert.Contains("table 1, row 1, cell 1", analysis.Errors[0].Location);
        if (name is "trPrChange" or "tblPrExChange") Assert.Contains("table 1, row 1", analysis.Errors[0].Location);
        if (name == "sectPrChange") Assert.Contains("document properties", analysis.Errors[0].Location);
    }

    [Theory]
    [InlineData("moveFromRange")] [InlineData("moveToRange")]
    [InlineData("customXmlInsRange")] [InlineData("customXmlDelRange")]
    [InlineData("customXmlMoveFromRange")] [InlineData("customXmlMoveToRange")]
    [InlineData("customXmlConflictInsRange")] [InlineData("customXmlConflictDelRange")]
    public void Both_ends_of_each_revision_range_are_rejected_in_source_order(string family)
    {
        var ns = family.Contains("Conflict") ? XNamespace.Get("http://schemas.microsoft.com/office/word/2010/wordml") : W;
        var start = Revision(ns + (family + "Start"));
        if (family.StartsWith("move")) start.Add(new XAttribute(W + "name", "Move"));
        var end = new XElement(ns + (family + "End"), new XAttribute(W + "id", 1));
        var template = Docx.Create(Docx.P(start, Docx.Run("{Name}"), end));
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{}", "template.tracked_revision");
        Assert.Equal(new[] { family + "Start", family + "End" }, analysis.Errors.Select(e => e.Location!.Split("tracked revision ")[1]));
        Assert.All(analysis.Errors, e => Assert.Contains("paragraph 1", e.Location));
        // An isolated end is intentionally not a valid paired revision range; detection must still work.
        var isolated = Reject(Docx.Create(Docx.P(new XElement(end))), "{}", "template.tracked_revision");
        Assert.EndsWith("tracked revision " + family + "End", Assert.Single(isolated.Errors).Location);
    }

    [Fact]
    public void Revision_ranges_in_all_stories_keep_stable_part_and_paragraph_order_and_error_cap()
    {
        XElement Paragraph() => Docx.P(Revision(W + "moveToRangeStart", new XAttribute(W + "name", "Moved")),
            Docx.Run("x"), new XElement(W + "moveToRangeEnd", new XAttribute(W + "id", 1)));
        var template = Prepare(Docx.Create(Paragraph()), document =>
        {
            var main = document.MainDocumentPart!;
            Save(main.AddNewPart<HeaderPart>(), new XElement(W + "hdr", Paragraph()));
            Save(main.AddNewPart<FooterPart>(), new XElement(W + "ftr", Paragraph()));
            Save(main.AddNewPart<FootnotesPart>(), new XElement(W + "footnotes", new XElement(W + "footnote", new XAttribute(W + "id", 1), Paragraph())));
            Save(main.AddNewPart<EndnotesPart>(), new XElement(W + "endnotes", new XElement(W + "endnote", new XAttribute(W + "id", 1), Paragraph())));
        });
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{}", "template.tracked_revision");
        Assert.Equal(new[] { "/word/document.xml", "/word/header1.xml", "/word/footer1.xml", "/word/footnotes.xml", "/word/endnotes.xml" },
            analysis.Errors.Select(e => e.Part).Distinct());
        Assert.Equal(10, analysis.Errors.Count);
        var many = Docx.Create(Enumerable.Range(1, 105).Select(id => Docx.P(Revision(W + "ins", Docx.Run("text"), id))).ToArray());
        Docx.AssertValid(many, FileFormatVersions.Office2010);
        Assert.Equal(105, Reject(many, "{}", "template.tracked_revision").Errors.Count);
    }

    [Fact]
    public void Deleted_payload_detection_does_not_depend_on_a_wrapper()
    {
        // Deliberately nonconformant isolated payload probes supplement valid deletion wrappers above.
        foreach (var name in new[] { "delText", "delInstrText" })
            Assert.Single(Reject(Docx.Create(Docx.P(new XElement(W + "r", new XElement(W + name, "old")))), "{}", "template.tracked_revision").Errors);
    }

    [Fact]
    public void Nonrevision_metadata_and_track_future_edits_are_not_tracked_content()
    {
        var template = Prepare(WithComments(Docx.Create(Docx.P(
            new XAttribute(W + "rsidR", "00AABBCC"), Start("bookmark", 0), Start("permission", 1), Start("comment", 2),
            new XElement(W + "proofErr", new XAttribute(W + "type", "spellStart")), Docx.Run("{Name}"),
            new XElement(W + "proofErr", new XAttribute(W + "type", "spellEnd")), End("comment", 2), End("permission", 1), End("bookmark", 0)),
            Docx.Table(new XElement(W + "tr", new XElement(W + "tc", new XElement(W + "tcPr", new XElement(W + "vMerge", new XAttribute(W + "val", "restart"))), Docx.P(Docx.Run("merged")))))), 2),
            document => Save(document.MainDocumentPart!.AddNewPart<DocumentSettingsPart>(), new XElement(W + "settings", new XElement(W + "trackRevisions"))));
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        using var input = new MemoryStream(template);
        Assert.True(Analyzer().Analyze(input).Valid);
        Assert.Equal("OKmerged", Docx.Text(Docx.Read(Docx.Merge(template, "{\"Name\":\"OK\"}", validationVersion: FileFormatVersions.Office2010).Document)));
    }

    [Theory]
    [InlineData("paragraph-mark")] [InlineData("row")] [InlineData("custom-xml")]
    public void Revision_detection_covers_property_marks_and_custom_xml_wrappers(string placement)
    {
        var template = placement switch
        {
            "paragraph-mark" => Docx.Create(Docx.P(new XElement(W + "pPr", new XElement(W + "rPr", Revision(W + "ins"))), Docx.Run("x"))),
            "row" => Docx.Create(Docx.Table(new XElement(W + "tr", new XElement(W + "trPr", Revision(W + "del")),
                new XElement(W + "tc", Docx.P(Docx.Run("x")))))),
            _ => Docx.Create(new XElement(W + "customXml", new XAttribute(W + "element", "custom"), Docx.P(
                Revision(W + "customXmlInsRangeStart"), Docx.Run("x"), new XElement(W + "customXmlInsRangeEnd", new XAttribute(W + "id", 1)))))
        };
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        var analysis = Reject(template, "{}", "template.tracked_revision");
        Assert.Equal(placement == "custom-xml" ? 2 : 1, analysis.Errors.Count);
    }

    [Fact]
    public void Revision_like_content_inside_opaque_pictures_retains_the_existing_exclusion()
    {
        XNamespace v = "urn:schemas-microsoft-com:vml";
        var picture = new XElement(W + "pict", new XElement(v + "shape", new XElement(v + "textbox",
            new XElement(W + "txbxContent", Docx.P(Revision(W + "ins", Docx.Run("{Hidden}")))))));
        var template = Docx.Create(Docx.P(new XElement(W + "r", picture), Docx.Run("{Name}")));
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        using var input = new MemoryStream(template);
        Assert.True(Analyzer().Analyze(input).Valid);
        var output = Docx.Merge(template, "{\"Name\":\"OK\"}", validationVersion: FileFormatVersions.Office2010).Document;
        Assert.True(XNode.DeepEquals(Docx.Read(template).Descendants(W + "pict").Single(), Docx.Read(output).Descendants(W + "pict").Single()));
        Assert.EndsWith("OK", Docx.Text(Docx.Read(output)));
    }

    private static XElement Revision(XName name, object? content = null, int id = 1) => new(name,
        new XAttribute(W + "id", id), new XAttribute(W + "author", "Reviewer"), new XAttribute(W + "date", "2026-10-18T00:00:00Z"), content);

    private static byte[] RevisionFixture(string name)
    {
        XNamespace w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        if (name is "ins" or "del" or "moveFrom" or "moveTo") return Docx.Create(Docx.P(Revision(W + name, Docx.Run("x"))));
        if (name is "conflictIns" or "conflictDel") return Docx.Create(Docx.P(Revision(w14 + name, Docx.Run("x"))));
        if (name is "delText" or "delInstrText") return Docx.Create(Docx.P(Revision(W + "del", new XElement(W + "r", new XElement(W + name, "old")))));
        if (name == "rPrChange") return Docx.Create(Docx.P(new XElement(W + "r", new XElement(W + "rPr", Revision(W + name, new XElement(W + "rPr"))), new XElement(W + "t", "x"))));
        if (name == "pPrChange") return Docx.Create(Docx.P(new XElement(W + "pPr", Revision(W + name, new XElement(W + "pPr"))), Docx.Run("x")));
        if (name == "numberingChange") return Docx.Create(Docx.P(new XElement(W + "pPr", new XElement(W + "numPr", Revision(W + name, new XAttribute(W + "original", "1")))), Docx.Run("x")));
        if (name == "sectPrChange") return Docx.Create(Docx.P(Docx.Run("x")), new XElement(W + "sectPr", Revision(W + name, new XElement(W + "sectPr"))));
        var table = Docx.Table(Docx.Row(Docx.Run("x")));
        if (name == "tblPrChange") table.Element(W + "tblPr")!.Add(Revision(W + name, new XElement(W + "tblPr")));
        else if (name == "tblGridChange") table.Element(W + "tblGrid")!.Add(new XElement(W + name, new XAttribute(W + "id", 1), new XElement(W + "tblGrid")));
        else if (name == "tblPrExChange") table.Element(W + "tr")!.AddFirst(new XElement(W + "tblPrEx", Revision(W + name, new XElement(W + "tblPrEx"))));
        else if (name == "trPrChange") table.Element(W + "tr")!.AddFirst(new XElement(W + "trPr", Revision(W + name, new XElement(W + "trPr"))));
        else table.Descendants(W + "tc").First().AddFirst(new XElement(W + "tcPr", Revision(W + name, name == "tcPrChange" ? new XElement(W + "tcPr") : null)));
        return Docx.Create(table);
    }
}
