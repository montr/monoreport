using System.Text.Json;
using System.Xml.Linq;
using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class UnifiedTemplateTests
{
    [Fact]
    public void Analyzer_builds_normative_requirements_and_occurrences()
    {
        var template = Docx.Create(
            Docx.P(Docx.Run("{Customer.Name}")),
            Docx.P(Docx.Run("{For:Orders}")),
            Docx.P(Docx.Run("{Number}")),
            Docx.P(Docx.Run("{If:Paid}")),
            Docx.P(Docx.Run("Оплачено")),
            Docx.P(Docx.Run("{/If}")),
            Docx.P(Docx.Run("{For:Items}")),
            Docx.P(Docx.Run("{Product}")),
            Docx.P(Docx.Run("{/For}")),
            Docx.P(Docx.Run("{/For}")));

        using var input = new MemoryStream(template);
        var result = Analyzer().Analyze(input);

        Assert.True(result.Valid, string.Join("; ", result.Errors.Select(error => error.Code + ":" + error.Location)));
        Assert.Collection(result.Requirements,
            customer =>
            {
                Assert.Equal(("Customer", "object"), (customer.Name, customer.Type));
                Assert.Equal(new[] { ("Name", "value") }, customer.Fields!.Select(field => (field.Name, field.Type)));
            },
            orders =>
            {
                Assert.Equal(("Orders", "array"), (orders.Name, orders.Type));
                var fields = Assert.IsAssignableFrom<IReadOnlyList<TemplateRequirement>>(orders.Fields);
                Assert.Equal(new[] { ("Number", "value"), ("Paid", "boolean"), ("Items", "array") },
                    fields.Select(field => (field.Name, field.Type)));
                Assert.Equal(("Product", "value"),
                    (fields[2].Fields![0].Name, fields[2].Fields![0].Type));
            });
        Assert.Equal(new[]
        {
            "$.Customer.Name", "$.Orders", "$.Orders[].Number", "$.Orders[].Paid",
            "$.Orders[].Items", "$.Orders[].Items[].Product"
        }, result.Occurrences.Select(item => item.Path));
    }

    [Fact]
    public void Analyzer_omits_fields_for_scalar_array_current_element()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{For:Tags}")), Docx.P(Docx.Run("{.}")),
            Docx.P(Docx.Run("{/For}")));
        using var input = new MemoryStream(template);

        var result = Analyzer().Analyze(input);

        var requirement = Assert.Single(result.Requirements);
        Assert.Equal(("Tags", "array", null), (requirement.Name, requirement.Type, requirement.Fields));
        Assert.Equal("$.Tags[]", result.Occurrences[1].Path);
    }

    [Fact]
    public void Renders_scalar_and_loop_from_one_plan_without_interpreting_inserted_syntax()
    {
        var template = Docx.Create(
            Docx.P(Docx.Run("Hello {Name}!")),
            Docx.P(Docx.Run("{For:Items}")),
            Docx.P(Docx.Run("[{. upper}]")),
            Docx.P(Docx.Run("{/For}")));

        var result = Docx.Merge(template, """{"Name":"{For:X}","Items":["one","{{two}}"]}""");

        Assert.Equal("Hello {For:X}![ONE][{{TWO}}]", Docx.Text(Docx.Read(result.Document)));
    }

    [Fact]
    public void Expressions_and_escapes_may_span_runs_and_emitted_braces_stay_literal()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{{"), Docx.Run("{Na"), Docx.Run("me}"),
            Docx.Run("}}"), Docx.Run("{A}"), Docx.Run("{B}")));
        var result = Docx.Merge(template, """{"Name":"X","A":"{B}","B":"Y"}""");
        Assert.Equal("{X}{B}Y", Docx.Text(Docx.Read(result.Document)));
    }

    [Fact]
    public void Scalar_output_with_boundary_whitespace_is_marked_for_word_preservation()
    {
        var result = Docx.Merge(Docx.Create(Docx.P(Docx.Run("{Value}"))), """{"Value":" padded "}""");
        var text = Assert.Single(Docx.Read(result.Document).Descendants(WordXml.TextElementName));
        Assert.Equal(" padded ", text.Value);
        Assert.Equal("preserve", (string?)text.Attribute(XNamespace.Xml + "space"));
    }

    [Theory]
    [InlineData("[]", "")]
    [InlineData("[\"A\"]", "A")]
    [InlineData("[\"A\",2,true,null]", "A2true")]
    public void For_handles_zero_one_many_and_mixed_scalar_elements(string array, string expected)
    {
        var template = Docx.Create(Docx.P(Docx.Run("{For:Items}")), Docx.P(Docx.Run("{.}")),
            Docx.P(Docx.Run("{/For}")));
        var result = Docx.Merge(template, $"{{\"Items\":{array}}}");
        Assert.Equal(expected, Docx.Text(Docx.Read(result.Document)));
    }

    [Fact]
    public void For_rejects_a_string_instead_of_iterating_its_characters()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{For:Items}")), Docx.P(Docx.Run("{.}")),
            Docx.P(Docx.Run("{/For}")));
        var error = Assert.Throws<TemplateException>(() => Docx.Merge(template, """{"Items":"ABC"}"""));
        Assert.Equal("generation.invalid_data", error.Code);
        Assert.Equal("$.Items", Assert.Single(error.Errors).DataPath);
    }

    [Fact]
    public void Mixed_array_property_access_uses_missing_value_policy()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{For:Items}")), Docx.P(Docx.Run("{Name}")),
            Docx.P(Docx.Run("{/For}")));

        var result = Docx.Merge(template, """{"Items":[{"Name":"A"},"x",1,true,null]}""",
            new MergeOptions { MissingValues = MissingValueBehavior.Warning });

        Assert.Equal("A", Docx.Text(Docx.Read(result.Document)));
        Assert.Equal(4, result.Warnings.Count);
    }

    [Theory]
    [InlineData(MissingValueBehavior.EmptyString)]
    [InlineData(MissingValueBehavior.Warning)]
    [InlineData(MissingValueBehavior.Error)]
    public void Property_access_on_every_primitive_and_null_element_uses_the_scalar_missing_policy(
        MissingValueBehavior behavior)
    {
        var template = Docx.Create(Docx.P(Docx.Run("{For:Items}")),
            Docx.P(Docx.Run("{Name}|{./Name}")), Docx.P(Docx.Run("{/For}")));
        var json = """{"Items":["text",12,true,null]}""";

        if (behavior == MissingValueBehavior.Error)
        {
            var error = Assert.Throws<TemplateException>(() => Docx.Merge(template, json,
                new MergeOptions { MissingValues = behavior }));
            Assert.Equal(8, error.Errors.Count);
            Assert.All(error.Errors, item => Assert.Equal("generation.invalid_data", item.Code));
            Assert.Equal(new[] { "$.Items[0].Name", "$.Items[0].Name" },
                error.Errors.Take(2).Select(item => item.DataPath));
            return;
        }

        var result = Docx.Merge(template, json, new MergeOptions { MissingValues = behavior });
        Assert.Equal("||||", Docx.Text(Docx.Read(result.Document)));
        Assert.Equal(behavior == MissingValueBehavior.Warning ? 8 : 0, result.Warnings.Count);
    }

    [Fact]
    public void Explicit_null_scalar_is_distinct_from_missing_and_objects_and_arrays_are_invalid_scalars()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Null}x{Missing}")));
        var result = Docx.Merge(template, """{"Null":null}""",
            new MergeOptions { MissingValues = MissingValueBehavior.Warning });
        Assert.Equal("x", Docx.Text(Docx.Read(result.Document)));
        Assert.Single(result.Warnings);

        foreach (var json in new[] { """{"Null":{}}""", """{"Null":[]}""" })
        {
            var error = Assert.Throws<TemplateException>(() => Docx.Merge(
                Docx.Create(Docx.P(Docx.Run("{Null}"))), json));
            Assert.Equal("generation.invalid_data", error.Code);
        }
    }

    [Fact]
    public void Structural_missing_and_null_are_quiet_rendering_choices_but_wrong_non_null_types_fail()
    {
        var template = Docx.Create(
            Docx.P(Docx.Run("{For:Items}")),
            Docx.P(Docx.Run("{If:Flag}")), Docx.P(Docx.Run("T")),
            Docx.P(Docx.Run("{Else}")), Docx.P(Docx.Run("F")), Docx.P(Docx.Run("{/If}")),
            Docx.P(Docx.Run("{For:Children}")), Docx.P(Docx.Run("X")), Docx.P(Docx.Run("{/For}")),
            Docx.P(Docx.Run("{/For}")));

        foreach (var behavior in Enum.GetValues<MissingValueBehavior>())
        {
            var result = Docx.Merge(template, """{"Items":["scalar",null,{"Flag":null,"Children":null}]}""",
                new MergeOptions { MissingValues = behavior });
            Assert.Equal("FFF", Docx.Text(Docx.Read(result.Document)));
            Assert.Equal(2, result.Warnings.Count);
            Assert.All(result.Warnings, warning =>
            {
                Assert.StartsWith("{", warning.FieldCode);
                Assert.StartsWith("$.Items[", warning.DataPath);
                Assert.Equal("/word/document.xml", warning.Part);
                Assert.Contains("paragraph", warning.Location);
            });
        }

        Assert.Throws<TemplateException>(() => Docx.Merge(template,
            """{"Items":[{"Flag":"yes","Children":{}}]}"""));
    }

    [Fact]
    public void Analyzer_reports_every_occurrence_in_an_incompatible_requirement_shape()
    {
        var template = Docx.Create(
            Docx.P(Docx.Run("{Value}")),
            Docx.P(Docx.Run("{For:Value}")), Docx.P(Docx.Run("{.}")), Docx.P(Docx.Run("{/For}")),
            Docx.P(Docx.Run("{Value}")));
        using var input = new MemoryStream(template);

        var result = Analyzer().Analyze(input);

        Assert.False(result.Valid);
        Assert.Equal(3, result.Errors.Count(error => error.Code == "template.requirement_conflict"));
        Assert.Equal(4, result.Occurrences.Count);
        Assert.Equal(new[] { "paragraph 1", "paragraph 2", "paragraph 5" }, result.Errors
            .Where(error => error.Code == "template.requirement_conflict")
            .Select(error => error.Location!.Split(", ").Last()));
    }

    [Fact]
    public void Boolean_if_and_scalar_output_are_compatible_and_first_occurrence_controls_order()
    {
        var template = Docx.Create(
            Docx.P(Docx.Run("{Second}")), Docx.P(Docx.Run("{First.Child}")),
            Docx.P(Docx.Run("{If:Second}")), Docx.P(Docx.Run("{Second}")), Docx.P(Docx.Run("{/If}")),
            Docx.P(Docx.Run("{First.Other}")));
        using var input = new MemoryStream(template);
        var result = Analyzer().Analyze(input);

        Assert.True(result.Valid);
        Assert.Equal(new[] { "Second", "First" }, result.Requirements.Select(item => item.Name));
        Assert.Equal("boolean", result.Requirements[0].Type);
        Assert.Equal(new[] { "Child", "Other" }, result.Requirements[1].Fields!.Select(item => item.Name));
        Assert.Equal(5, result.Occurrences.Count);
    }

    [Fact]
    public void Analyzer_uses_deterministic_story_order_and_preserves_cross_part_duplicates()
    {
        var template = MultiStoryTemplate();
        using var firstInput = new MemoryStream(template);
        using var secondInput = new MemoryStream(template);

        var first = Analyzer().Analyze(firstInput);
        var second = Analyzer().Analyze(secondInput);

        var expectedParts = new[]
        {
            "/word/document.xml", "/word/header1.xml", "/word/header2.xml",
            "/word/footer1.xml", "/word/footer2.xml", "/word/footnotes.xml", "/word/endnotes.xml"
        };
        Assert.Equal(expectedParts, first.Occurrences.Select(item => item.Part));
        Assert.Equal(expectedParts, second.Occurrences.Select(item => item.Part));
        Assert.Equal(7, first.Occurrences.Count);
        Assert.Single(first.Requirements);
    }

    [Fact]
    public void Analyzer_orders_failed_and_successful_story_diagnostics_by_traversal_order()
    {
        var w = WordXml.WordprocessingNamespace;
        using var prepared = new MemoryStream();
        prepared.Write(Docx.Create(Docx.P(Docx.Run("main"))));
        prepared.Position = 0;
        using (var document = WordprocessingDocument.Open(prepared, true))
        {
            var main = document.MainDocumentPart!;
            var header = main.AddNewPart<HeaderPart>();
            using (var target = header.GetStream(FileMode.Create, FileAccess.Write))
                new XDocument(new XElement(w + "hdr", Docx.P(Docx.Run("header")))).Save(target);
            var footer = main.AddNewPart<FooterPart>();
            using (var target = footer.GetStream(FileMode.Create, FileAccess.Write))
                new XDocument(new XElement(w + "ftr", Docx.P(Docx.Run("{Broken")))).Save(target);
        }
        using var corrupted = new MemoryStream(prepared.ToArray());
        using (var archive = new ZipArchive(corrupted, ZipArchiveMode.Update, leaveOpen: true))
        {
            var header = archive.GetEntry("word/header1.xml")!;
            header.Delete();
            using var target = archive.CreateEntry("word/header1.xml").Open();
            using var writer = new StreamWriter(target);
            writer.Write("<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p>");
        }
        corrupted.Position = 0;

        var analysis = Analyzer().Analyze(corrupted);

        Assert.Equal(new[] { "/word/header1.xml", "/word/footer1.xml" },
            analysis.Errors.Select(issue => issue.Part));
        Assert.Equal(new[] { "template.invalid_xml", "template.unmatched_brace" },
            analysis.Errors.Select(issue => issue.Code));
    }

    private static byte[] MultiStoryTemplate()
    {
        var w = WordXml.WordprocessingNamespace;
        using var stream = new MemoryStream();
        stream.Write(Docx.Create(Docx.P(Docx.Run("{Shared}"))));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var parts = new OpenXmlPart[]
            {
                main.AddNewPart<HeaderPart>(), main.AddNewPart<HeaderPart>(),
                main.AddNewPart<FooterPart>(), main.AddNewPart<FooterPart>(),
                main.AddNewPart<FootnotesPart>(), main.AddNewPart<EndnotesPart>()
            };
            foreach (var part in parts)
            {
                var root = part switch
                {
                    HeaderPart => new XElement(w + "hdr", Docx.P(Docx.Run("{Shared}"))),
                    FooterPart => new XElement(w + "ftr", Docx.P(Docx.Run("{Shared}"))),
                    FootnotesPart => new XElement(w + "footnotes", new XElement(w + "footnote",
                        new XAttribute(w + "id", "1"), Docx.P(Docx.Run("{Shared}")))),
                    _ => new XElement(w + "endnotes", new XElement(w + "endnote",
                        new XAttribute(w + "id", "1"), Docx.P(Docx.Run("{Shared}"))))
                };
                using var target = part.GetStream(FileMode.Create, FileAccess.Write);
                new XDocument(root).Save(target);
            }
        }
        return stream.ToArray();
    }

    private static TemplateAnalyzer Analyzer()
    {
        var reader = new DocxStoryReader();
        return new(reader, new DocxTemplatePlanner(new TemplateLanguageParser()));
    }
}
