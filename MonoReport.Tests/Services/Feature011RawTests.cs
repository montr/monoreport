using System.Xml.Linq;
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
        Directory.CreateDirectory("/private/tmp/feature011-evidence/artifacts");
        File.WriteAllBytes("/private/tmp/feature011-evidence/artifacts/raw-merged-table.docx", outputBytes);
        var output = Docx.Read(outputBytes);
        Assert.True(XNode.DeepEquals(table, output.Descendants(W + "tbl").Last()));
        Assert.Equal("{Malformed{invalid path}nested{{OK", Docx.Text(output));
        Assert.Equal(2, output.Descendants(W + "vMerge").Count());
        Assert.Equal("", Docx.Text(Docx.Read(Docx.Merge(Docx.Create(P("{Raw}"), P("{/Raw}")), "{}").Document)));
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
}
