using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature011RawTests
{
    private static readonly XNamespace W = WordXml.WordprocessingNamespace;
    private static TemplateAnalyzer Analyzer() => Feature010IdentityTests.Analyzer();
    private static XElement P(string text) => Docx.P(Docx.Run(text));

    [Fact]
    public void Former_closers_are_ordinary_scalars_and_Raw_property_remains_accessible()
    {
        var source = Docx.Create(P("{EndIf}|{EndFor}|{./Raw}"));
        using var input = new MemoryStream(source);
        var analysis = Analyzer().Analyze(input);
        Assert.True(analysis.Valid);
        Assert.Equal(new[] { "EndIf", "EndFor", "Raw" }, analysis.Requirements.Select(item => item.Name));
        Assert.All(analysis.Occurrences, item => Assert.Equal("field", item.Kind));
        Assert.Equal("yes|next|literal", Docx.Text(Docx.Read(Docx.Merge(source,
            """{"EndIf":"yes","EndFor":"next","Raw":"literal"}""").Document)));
        Assert.Equal("||", Docx.Text(Docx.Read(Docx.Merge(source, "{}").Document)));
        Feature010IdentityTests.Reject(Docx.Create(P("{If:Show}"), P("{EndIf}")), "{}", "template.invalid_structure");
    }

    [Fact]
    public void Raw_payload_is_not_tokenized_and_active_dependencies_resume_in_source_context()
    {
        var source = Docx.Create(P("{For:Items}"), P("{Raw}"),
            P("{Name} {broken { ../ {Name unknown} }} {{ {Else} {/For} {Raw}"), P("{/Raw}"),
            P("{$.Title} {Name} {../EndFor}"), P("{/For}"));
        using var input = new MemoryStream(source);
        var analysis = Analyzer().Analyze(input);
        Assert.True(analysis.Valid, string.Join(";", analysis.Errors.Select(item => item.Message)));
        Assert.Empty(analysis.Warnings);
        Assert.Equal(new[] { "$.Items", "$.Title", "$.Items[].Name", "$.EndFor" },
            analysis.Occurrences.Select(item => item.Path));
        Assert.Equal(new[] { "Items", "Title", "EndFor" }, analysis.Requirements.Select(item => item.Name));
        Assert.Equal("Name", Assert.Single(analysis.Requirements[0].Fields!).Name);
    }

    [Fact]
    public void First_eligible_close_wins_and_wrong_parent_inline_escaped_and_opaque_closes_stay_literal()
    {
        var source = Docx.Create(Docx.P(Docx.Run("  {Ra"), Docx.Run("w}  ")),
            P("inline {/Raw}"), P("{{/Raw}}"), Docx.Table(Docx.Row(Docx.Run("{/Raw}"))),
            Docx.P(Docx.Field(" REF X ", "{/Raw}")),
            new XElement(W + "sdt", new XElement(W + "sdtContent", P("{/Raw}"))),
            P("{Raw} {Bad syntax}"), P("{/Raw}"), P("{After}"));
        using var input = new MemoryStream(source);
        var analysis = Analyzer().Analyze(input);
        Assert.True(analysis.Valid, string.Join(";", analysis.Errors.Select(item => item.Message)));
        Assert.Equal("After", Assert.Single(analysis.Requirements).Name);
        Assert.Equal("{After}", Assert.Single(analysis.Occurrences).Expression);
    }

    [Theory]
    [InlineData("{Raw}")]
    [InlineData("{/Raw}")]
    [InlineData("inline {Raw}")]
    [InlineData("inline {/Raw}")]
    public void Invalid_active_delimiters_have_shared_structural_errors(string text) =>
        Feature010IdentityTests.Reject(Docx.Create(P(text)), "{}", "template.invalid_structure");

    [Fact]
    public void Raw_does_not_pair_across_cells_tables_or_paragraph_row_layouts()
    {
        var variants = new[]
        {
            Docx.Create(Docx.Table(new XElement(W + "tr", new XElement(W + "tc", P("{Raw}")),
                new XElement(W + "tc", P("{/Raw}"))))),
            Docx.Create(Docx.Table(Docx.Row(Docx.Run("{Raw}"))), Docx.Table(Docx.Row(Docx.Run("{/Raw}")))),
            Docx.Create(P("{Raw}"), Docx.Table(Docx.Row(Docx.Run("{/Raw}"))))
        };
        foreach (var source in variants)
            Feature010IdentityTests.Reject(source, "{}", "template.invalid_structure");
    }

    [Fact]
    public void Raw_delimiters_cannot_pair_across_document_parts()
    {
        var source = Feature010IdentityTests.Prepare(Docx.Create(P("{Raw}")), document =>
            Feature010IdentityTests.Save(document.MainDocumentPart!.AddNewPart<HeaderPart>(),
                new XElement(W + "hdr", P("{/Raw}"))));

        var analysis = Feature010IdentityTests.Reject(source, "{}", "template.invalid_structure");

        Assert.All(analysis.Errors, error => Assert.Equal("template.invalid_structure", error.Code));
        Assert.Contains(analysis.Errors, error => error.Part == "/word/document.xml");
        Assert.Contains(analysis.Errors, error => error.Part!.Contains("header", StringComparison.Ordinal));
    }

    [Fact]
    public void Raw_delimiters_cannot_pair_across_independent_note_roots()
    {
        var source = Feature010IdentityTests.Prepare(Docx.Create(P("main")), document =>
            Feature010IdentityTests.Save(document.MainDocumentPart!.AddNewPart<FootnotesPart>(),
                new XElement(W + "footnotes",
                    new XElement(W + "footnote", new XAttribute(W + "id", 1), P("{Raw}")),
                    new XElement(W + "footnote", new XAttribute(W + "id", 2), P("{/Raw}")))));

        var analysis = Feature010IdentityTests.Reject(source, "{}", "template.invalid_structure");

        Assert.All(analysis.Errors, error => Assert.Equal("template.invalid_structure", error.Code));
        Assert.All(analysis.Errors, error => Assert.Equal("/word/footnotes.xml", error.Part));
    }

    [Fact]
    public void Nested_branches_and_zero_one_many_copies_preserve_literal_XML_and_context()
    {
        var literal = Docx.P(new XElement(W + "pPr", new XElement(W + "keepNext")),
            new XElement(W + "r", Docx.Properties(), new XElement(W + "t", "{Name unknown} {{ {Else} {/If} {Raw}")));
        var source = Docx.Create(P("{For:Items}"), P("{If:Show}"), P("{Raw}"), literal, P("{/Raw}"),
            P("{Name}"), P("{Else}"), P("{Raw}"), P("{Malformed"), P("{/Raw}"), P("{/If}"), P("{/For}"), P("{Tail}"));
        foreach (var count in new[] { 0, 1, 3 })
        {
            var items = Enumerable.Range(0, count).Select(index => new { Name = index.ToString(), Show = index % 2 == 0 });
            var output = Docx.Merge(source, System.Text.Json.JsonSerializer.Serialize(new { Items = items, Tail = "end" })).Document;
            var xml = Docx.Read(output);
            Assert.Equal(string.Concat(Enumerable.Range(0, count).Select(index => index % 2 == 0 ?
                Docx.Text(literal) + index : "{Malformed")) + "end", Docx.Text(xml));
            Assert.All(xml.Descendants(W + "p").Where(p => p.Element(W + "pPr") is not null),
                p => Assert.True(XNode.DeepEquals(literal, p)));
        }
    }

    [Fact]
    public void Raw_preserves_multiple_cell_paragraphs_complete_tables_and_nested_tables()
    {
        var table = Docx.Table(
            new XElement(W + "tr", new XElement(W + "tc", new XElement(W + "tcPr",
                new XElement(W + "vMerge", new XAttribute(W + "val", "restart"))), P("{invalid path}"))),
            new XElement(W + "tr", new XElement(W + "tc", new XElement(W + "tcPr",
                new XElement(W + "vMerge", new XAttribute(W + "val", "continue"))), P("nested"))));
        var body = new XElement(W + "tc", P("{Raw}"), P("{Malformed"), table, P("{{"), P("{/Raw}"), P("{After}"));
        var source = Docx.Create(Docx.Table(new XElement(W + "tr", body)));
        var outputBytes = Docx.Merge(source, "{\"After\":\"OK\"}").Document;
        var output = Docx.Read(outputBytes);
        Assert.True(XNode.DeepEquals(table, output.Descendants(W + "tbl").Last()));
        Assert.Equal("{Malformed{invalid path}nested{{OK", Docx.Text(output));
        Assert.Equal(2, output.Descendants(W + "vMerge").Count());
        Assert.Equal("", Docx.Text(Docx.Read(Docx.Merge(Docx.Create(P("{Raw}"), P("{/Raw}")), "{}").Document)));
    }

    [Fact]
    public void Row_Raw_preserves_a_surviving_opaque_row_content_control()
    {
        var control = new XElement(W + "sdt",
            new XElement(W + "sdtPr", new XElement(W + "alias", new XAttribute(W + "val", "opaque row")),
                new XElement(W + "id", new XAttribute(W + "val", 42)),
                new XElement(W + "lock", new XAttribute(W + "val", "sdtLocked"))),
            new XElement(W + "sdtContent", Docx.Row(Docx.Run("{Hidden} {Malformed {/Raw}"))));
        var table = Docx.Table(Docx.Row(Docx.Run("{Raw}")), control, Docx.Row(Docx.Run("{/Raw}")));
        var source = Docx.Create(table, P("{After}"));
        Docx.AssertValid(source);
        using var input = new MemoryStream(source);
        var analysis = Analyzer().Analyze(input);
        Assert.True(analysis.Valid, string.Join("; ", analysis.Errors.Select(error => error.Message)));
        Assert.Equal("After", Assert.Single(analysis.Requirements).Name);
        Assert.Equal("{After}", Assert.Single(analysis.Occurrences).Expression);

        var result = Docx.Merge(source, "{\"After\":\"done\"}");
        Assert.Empty(result.Warnings);
        var output = Docx.Read(result.Document);
        Assert.True(XNode.DeepEquals(control, output.Descendants(W + "sdt").Single()));
        Assert.True(XNode.DeepEquals(new XElement(W + "tbl", table.Elements().Take(2), control),
            output.Descendants(W + "tbl").Single()));
        Assert.Equal("{Hidden} {Malformed {/Raw}done", Docx.Text(output));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Row_Raw_checks_merge_origins_through_opaque_row_content_controls(bool deleteRestart)
    {
        XElement Row(string text, string? merge = null) => new(W + "tr", new XElement(W + "tc",
            merge is null ? null : new XElement(W + "tcPr",
                new XElement(W + "vMerge", new XAttribute(W + "val", merge))), P(text)));
        var control = new XElement(W + "sdt", new XElement(W + "sdtPr"),
            new XElement(W + "sdtContent", Row("{Hidden}", "continue")));
        var restart = Row("restart", "restart");
        var source = Docx.Create(Docx.Table(deleteRestart ? Row("{Raw}", "restart") : restart,
            deleteRestart ? control : Row("{Raw}", "continue"),
            deleteRestart ? Row("{/Raw}") : control,
            deleteRestart ? Row("outside") : Row("{/Raw}")));
        Docx.AssertValid(source);
        if (deleteRestart)
        {
            var analysis = Feature010IdentityTests.Reject(source, "{}", "template.invalid_structure");
            Assert.Contains("a surviving vertical merge chain", Assert.Single(analysis.Errors).Message);
        }
        else
        {
            using var input = new MemoryStream(source);
            Assert.True(Analyzer().Analyze(input).Valid);
            var output = Docx.Read(Docx.Merge(source, "{}").Document);
            Assert.True(XNode.DeepEquals(control, output.Descendants(W + "sdt").Single()));
            Assert.True(XNode.DeepEquals(restart, output.Descendants(W + "tbl").Single().Elements(W + "tr").Single()));
            Assert.Equal(2, output.Descendants(W + "vMerge").Count());
        }
    }

    [Fact]
    public void Field_state_continues_through_literal_Raw_without_activating_cached_delimiters()
    {
        var source = Docx.Create(P("{Raw}"), Docx.P(new XElement(W + "r", Docx.Character("begin"))), P("{/Raw}"),
            P("{Malformed"), Docx.P(new XElement(W + "r", Docx.Character("end"))), P("{/Raw}"), P("{After}"));
        using var input = new MemoryStream(source);
        var analysis = Analyzer().Analyze(input);
        Assert.True(analysis.Valid);
        Assert.Equal("After", Assert.Single(analysis.Requirements).Name);
        var output = Docx.Read(Docx.Merge(source, "{\"After\":\"done\"}").Document);
        Assert.Equal("{/Raw}{Malformeddone", Docx.Text(output));
        Assert.Equal(2, output.Descendants(W + "fldChar").Count());
    }

    [Theory]
    [InlineData("sdt", true)]
    [InlineData("customXml", true)]
    [InlineData("sdt", false)]
    [InlineData("customXml", false)]
    public void Row_Raw_checks_physical_cells_through_supported_wrappers(string wrapper, bool deleteRestart)
    {
        XElement Cell(string text, int span, string? merge = null) => new(W + "tc",
            new XElement(W + "tcPr", new XElement(W + "gridSpan", new XAttribute(W + "val", span)),
                merge is null ? null : new XElement(W + "vMerge", new XAttribute(W + "val", merge))), P(text));
        XElement Row(string text, string? merge = null) => new(W + "tr",
            new XElement(W + "trPr", new XElement(W + "gridBefore", new XAttribute(W + "val", 1))),
            Cell("", 2), Cell(text, 2, merge), Cell("", 1));

        var wrappedPrefix = Cell("", 2);
        wrappedPrefix.Add(new XElement(W + "sdt",
            new XElement(W + "sdtPr", new XElement(W + "tag", new XAttribute(W + "val", "opaque nested table"))),
            new XElement(W + "sdtContent", P("{Hidden} {Malformed {/Raw}"),
                Docx.Table(new XElement(W + "tr", Cell("nested restart", 1, "restart")),
                    new XElement(W + "tr", Cell("nested continuation", 1, "continue"))))), P(""));
        var wrappedCells = new[] { wrappedPrefix, Cell("wrapped", 2, deleteRestart ? "continue" : "restart") };
        var container = wrapper == "sdt"
            ? new XElement(W + "sdt", new XElement(W + "sdtPr",
                    new XElement(W + "alias", new XAttribute(W + "val", "physical cells")),
                    new XElement(W + "id", new XAttribute(W + "val", 73)),
                    new XElement(W + "lock", new XAttribute(W + "val", "sdtLocked"))),
                new XElement(W + "sdtContent", wrappedCells))
            : new XElement(W + "customXml", new XAttribute(W + "element", "cells"),
                new XElement(W + "customXmlPr", new XElement(W + "attr", new XAttribute(W + "name", "preserved"),
                    new XAttribute(W + "val", "metadata"))), wrappedCells);
        var wrappedRow = new XElement(W + "tr",
            new XElement(W + "trPr", new XElement(W + "gridBefore", new XAttribute(W + "val", 1))),
            container, Cell("unrelated suffix", 1));
        var opening = Row("{Raw}", deleteRestart ? "restart" : "continue");
        var closing = Row("{/Raw}");
        var body = Row("literal {Hidden}", "continue");
        var table = Docx.Table(deleteRestart ? opening : wrappedRow,
            deleteRestart ? wrappedRow : opening, deleteRestart ? closing : body);
        if (!deleteRestart) table.Add(closing);
        table.Element(W + "tblGrid")!.ReplaceNodes(Enumerable.Range(0, 6)
            .Select(_ => new XElement(W + "gridCol", new XAttribute(W + "w", 2000))));
        var unrelatedTable = Docx.Table(Docx.Row(Docx.Run("unrelated table")));
        var source = Docx.Create(table, unrelatedTable, P("{After}"));
        Docx.AssertValid(source, FileFormatVersions.Office2010);

        using var input = new MemoryStream(source);
        var analysis = Analyzer().Analyze(input);
        if (deleteRestart)
        {
            Assert.False(analysis.Valid);
            Assert.Contains("a surviving vertical merge chain", Assert.Single(analysis.Errors).Message);
            input.Position = 0;
            using var data = System.Text.Json.JsonDocument.Parse("{\"After\":\"done\"}");
            MergeResult? partial = null;
            var error = Assert.Throws<TemplateException>(() => partial = Docx.CreateEngine().Merge(input, data));
            Assert.Null(partial);
            Assert.Equal(analysis.Errors.Select(item => (item.Code, item.Part, item.Location, item.Message)),
                error.Errors.Select(item => (item.Code, item.Part, item.Location, item.Message)));
            Assert.All(error.Errors, item => Assert.Equal("template.invalid_structure", item.Code));
            Assert.Equal(source, input.ToArray());
            Assert.True(XNode.DeepEquals(table, Docx.Read(source).Descendants(W + "tbl").First()));
            return;
        }

        Assert.True(analysis.Valid, string.Join("; ", analysis.Errors.Select(item => item.Message)));
        Assert.Equal("After", Assert.Single(analysis.Requirements).Name);
        Assert.Equal("{After}", Assert.Single(analysis.Occurrences).Expression);
        var result = Docx.Merge(source, "{\"After\":\"done\"}", validationVersion: FileFormatVersions.Office2010);
        Assert.Empty(result.Warnings);
        var output = Docx.Read(result.Document);
        var actualTable = output.Descendants(W + "tbl").First();
        var expectedTable = new XElement(W + "tbl", table.Elements().Take(2), wrappedRow, body);
        Assert.True(XNode.DeepEquals(expectedTable, actualTable));
        Assert.True(XNode.DeepEquals(container, actualTable.Descendants(W + wrapper).First()));
        Assert.True(XNode.DeepEquals(unrelatedTable, output.Root!.Element(W + "body")!.Elements(W + "tbl").Last()));
        Assert.Equal("done", Docx.Text(output.Root.Element(W + "body")!.Elements(W + "p").Last()));

        // Inspect only outer physical cells; the preserved nested table has its own merge chain.
        var rows = actualTable.Elements(W + "tr").ToArray();
        Assert.Equal(2, rows.Length);
        var firstCells = (wrapper == "sdt" ? rows[0].Element(W + "sdt")!.Element(W + "sdtContent")! :
            rows[0].Element(W + "customXml")!).Elements(W + "tc").Concat(rows[0].Elements(W + "tc")).ToArray();
        var secondCells = rows[1].Elements(W + "tc").ToArray();
        foreach (var cells in new[] { firstCells, secondCells })
        {
            Assert.Equal(new[] { 2, 2, 1 }, cells.Select(cell => (int)cell.Element(W + "tcPr")!.Element(W + "gridSpan")!.Attribute(W + "val")!));
            Assert.Null(cells[0].Element(W + "tcPr")!.Element(W + "vMerge"));
            Assert.Null(cells[2].Element(W + "tcPr")!.Element(W + "vMerge"));
        }
        Assert.Equal(new[] { "restart", "continue" }, new[] { firstCells[1], secondCells[1] }
            .Select(cell => (string?)cell.Element(W + "tcPr")!.Element(W + "vMerge")!.Attribute(W + "val")));
        Assert.All(rows, row => Assert.Equal(1, (int)row.Element(W + "trPr")!.Element(W + "gridBefore")!.Attribute(W + "val")!));
    }
}
