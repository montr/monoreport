using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature010DocxTests
{
    [Fact]
    public void Nested_for_if_parent_and_root_context_render_recursively()
    {
        var template = Docx.Create(
            Docx.P(Docx.Run("{For:Orders}")),
            Docx.P(Docx.Run("{Number}:{$.Company}")),
            Docx.P(Docx.Run("{For:Items}")),
            Docx.P(Docx.Run("{../Number}/{Name}")),
            Docx.P(Docx.Run("{If:Paid}")),
            Docx.P(Docx.Run("+")),
            Docx.P(Docx.Run("{Else}")),
            Docx.P(Docx.Run("-")),
            Docx.P(Docx.Run("{EndIf}")),
            Docx.P(Docx.Run("{EndFor}")),
            Docx.P(Docx.Run("{EndFor}")));

        var result = Docx.Merge(template,
            """{"Company":"ACME","Orders":[{"Number":"1","Items":[{"Name":"A","Paid":true},{"Name":"B","Paid":false}]}]}""");

        Assert.Equal("1:ACME1/A+1/B-", Docx.Text(Docx.Read(result.Document)));
    }

    [Fact]
    public void Paragraph_region_can_repeat_a_complete_table_and_preserve_properties()
    {
        var table = Docx.Table(Docx.Row(Docx.Run("{.}")));
        var template = Docx.Create(Docx.P(Docx.Run("{For:Values}")), table,
            Docx.P(Docx.Run("{EndFor}")));

        var result = Docx.Merge(template, """{"Values":["A","B"]}""");
        var xml = Docx.Read(result.Document);

        Assert.Equal("AB", Docx.Text(xml));
        Assert.Equal(2, xml.Descendants(WordXml.TableElementName).Count());
        Assert.All(xml.Descendants(WordXml.TableElementName), item => Assert.Single(item.Elements(WordXml.TablePropertiesElementName)));
    }

    [Fact]
    public void Paragraph_if_selects_a_complete_table_and_removes_the_other_branch_and_markers()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{If:Show}")),
            Docx.Table(Docx.Row(Docx.Run("shown"))), Docx.P(Docx.Run("{Else}")),
            Docx.Table(Docx.Row(Docx.Run("hidden"))), Docx.P(Docx.Run("{EndIf}")));
        var result = Docx.Merge(template, """{"Show":true}""");
        var xml = Docx.Read(result.Document);
        Assert.Equal("shown", Docx.Text(xml));
        Assert.Single(xml.Descendants(WordXml.TableElementName));
    }

    [Fact]
    public void Row_region_repeats_rows_and_preserves_horizontal_merge_and_list_properties()
    {
        var w = WordXml.WordprocessingNamespace;
        var markerStart = new XElement(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName, Docx.P(Docx.Run("{For:Items}"))));
        var content = new XElement(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName,
                new XElement(w + "tcPr", new XElement(w + "gridSpan", new XAttribute(w + "val", "2"))),
                new XElement(WordXml.ParagraphElementName,
                    new XElement(w + "pPr", new XElement(w + "numPr", new XElement(w + "ilvl", new XAttribute(w + "val", "1")), new XElement(w + "numId", new XAttribute(w + "val", "7")))),
                    Docx.Run("{Name}"))));
        var markerEnd = new XElement(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName, Docx.P(Docx.Run("{EndFor}"))));
        var template = Docx.Create(Docx.Table(markerStart, content, markerEnd));

        var result = Docx.Merge(template, """{"Items":[{"Name":"A"},{"Name":"B"}]}""");
        var xml = Docx.Read(result.Document);

        Assert.Equal("AB", Docx.Text(xml));
        Assert.Equal(2, xml.Descendants(w + "gridSpan").Count());
        Assert.Equal(2, xml.Descendants(w + "numPr").Count());
    }

    [Fact]
    public void Repeated_numbered_bulleted_and_multilevel_paragraphs_keep_native_numbering_references()
    {
        var w = WordXml.WordprocessingNamespace;
        XElement ListParagraph(string text, int numId, int level) => new(WordXml.ParagraphElementName,
            new XElement(w + "pPr", new XElement(w + "numPr",
                new XElement(w + "ilvl", new XAttribute(w + "val", level)),
                new XElement(w + "numId", new XAttribute(w + "val", numId))),
                new XElement(w + "ind", new XAttribute(w + "left", 720 + level * 360))),
            Docx.Run(text));
        var template = Docx.Create(
            Docx.P(Docx.Run("{For:Items}")),
            ListParagraph("{Name}", 3, 0),
            ListParagraph("{Name} child", 3, 1),
            ListParagraph("{Name} bullet", 8, 0),
            Docx.P(Docx.Run("{EndFor}")));

        var result = Docx.Merge(template, """{"Items":[{"Name":"A"},{"Name":"B"}]}""");
        var xml = Docx.Read(result.Document);
        var properties = xml.Descendants(w + "numPr").ToArray();
        Assert.Equal(6, properties.Length);
        Assert.Equal(new[] { "3", "3", "8", "3", "3", "8" },
            properties.Select(item => (string?)item.Element(w + "numId")?.Attribute(w + "val")));
        Assert.Equal(new[] { "0", "1", "0", "0", "1", "0" },
            properties.Select(item => (string?)item.Element(w + "ilvl")?.Attribute(w + "val")));
        Assert.Equal(6, xml.Descendants(w + "ind").Count());
    }

    [Fact]
    public void Numbering_definition_part_is_preserved_byte_for_byte()
    {
        var w = WordXml.WordprocessingNamespace;
        var template = Docx.Create(Docx.P(Docx.Run("{For:Items}")),
            new XElement(WordXml.ParagraphElementName,
                new XElement(w + "pPr", new XElement(w + "numPr",
                    new XElement(w + "ilvl", new XAttribute(w + "val", "0")),
                    new XElement(w + "numId", new XAttribute(w + "val", "4")))), Docx.Run("{.}")),
            Docx.P(Docx.Run("{EndFor}")));
        var numbering = new XDocument(new XElement(w + "numbering",
            new XElement(w + "abstractNum", new XAttribute(w + "abstractNumId", "2")),
            new XElement(w + "num", new XAttribute(w + "numId", "4"),
                new XElement(w + "abstractNumId", new XAttribute(w + "val", "2")))));
        using var prepared = new MemoryStream();
        prepared.Write(template);
        prepared.Position = 0;
        using (var document = WordprocessingDocument.Open(prepared, true))
        {
            var part = document.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();
            using var target = part.GetStream(FileMode.Create, FileAccess.Write);
            numbering.Save(target, SaveOptions.DisableFormatting);
        }

        var preparedBytes = prepared.ToArray();
        string NumberingXml(byte[] bytes)
        {
            using var package = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(package, false);
            using var part = document.MainDocumentPart!.NumberingDefinitionsPart!.GetStream();
            using var reader = new StreamReader(part);
            return reader.ReadToEnd();
        }
        var expected = NumberingXml(preparedBytes);
        var result = Docx.Merge(preparedBytes, """{"Items":["A","B"]}""");
        Assert.Equal(expected, NumberingXml(result.Document));
    }

    [Fact]
    public void Nested_table_row_loop_operates_only_on_the_inner_table()
    {
        var inner = Docx.Table(Docx.Row(Docx.Run("{For:Values}")),
            Docx.Row(Docx.Run("{.}")), Docx.Row(Docx.Run("{EndFor}")));
        var outerRow = new XElement(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName, Docx.P(Docx.Run("Outer")), inner, Docx.P()));
        var result = Docx.Merge(Docx.Create(Docx.Table(outerRow)), """{"Values":["A","B"]}""");
        var xml = Docx.Read(result.Document);
        Assert.Equal("OuterAB", Docx.Text(xml));
        Assert.Single(xml.Descendants(WordXml.TableElementName).First().Elements(WordXml.TableRowElementName));
        Assert.Equal(2, xml.Descendants(WordXml.TableElementName).Last().Elements(WordXml.TableRowElementName).Count());
    }

    [Theory]
    [InlineData("{For:Items} extra")]
    [InlineData("{For:Items}{EndFor}")]
    public void Analyzer_rejects_mixed_or_multiple_control_markers(string text)
    {
        using var stream = new MemoryStream(Docx.Create(Docx.P(Docx.Run(text))));
        var analysis = Analyzer().Analyze(stream);
        Assert.False(analysis.Valid);
        Assert.Contains(analysis.Errors, issue => issue.Code == "template.invalid_structure");
    }

    [Fact]
    public void Analyzer_rejects_vertical_merge_in_row_region()
    {
        var w = WordXml.WordprocessingNamespace;
        var template = Docx.Create(Docx.Table(
            Docx.Row(Docx.Run("{For:Items}")),
            new XElement(WordXml.TableRowElementName, new XElement(WordXml.TableCellElementName,
                new XElement(w + "tcPr", new XElement(w + "vMerge", new XAttribute(w + "val", "restart"))),
                Docx.P(Docx.Run("{Name}")))),
            Docx.Row(Docx.Run("{EndFor}"))));

        using var stream = new MemoryStream(template);
        var result = Analyzer().Analyze(stream);

        Assert.False(result.Valid);
        Assert.Contains(result.Errors, error => error.Code == "template.invalid_structure");
    }

    [Fact]
    public void Vertical_merge_outside_a_row_region_and_inside_an_intact_cloned_table_is_preserved()
    {
        var w = WordXml.WordprocessingNamespace;
        XElement MergedRow(string value) => new(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName,
                new XElement(w + "tcPr", new XElement(w + "vMerge", new XAttribute(w + "val", value))),
                Docx.P(Docx.Run("merged"))));
        var rowTemplate = Docx.Create(Docx.Table(MergedRow("restart"), MergedRow("continue"),
            Docx.Row(Docx.Run("{For:Items}")), Docx.Row(Docx.Run("{.}")), Docx.Row(Docx.Run("{EndFor}"))));
        var rowResult = Docx.Merge(rowTemplate, """{"Items":["A","B"]}""");
        Assert.Equal(2, Docx.Read(rowResult.Document).Descendants(w + "vMerge").Count());

        var completeTable = Docx.Table(MergedRow("restart"), MergedRow("continue"));
        var paragraphTemplate = Docx.Create(Docx.P(Docx.Run("{For:Items}")), completeTable,
            Docx.P(Docx.Run("{EndFor}")));
        var paragraphResult = Docx.Merge(paragraphTemplate, """{"Items":[1,2]}""");
        Assert.Equal(4, Docx.Read(paragraphResult.Document).Descendants(w + "vMerge").Count());
    }

    [Fact]
    public void Conditional_row_keeps_only_the_selected_sibling_rows()
    {
        var template = Docx.Create(Docx.Table(Docx.Row(Docx.Run("{If:Show}")),
            Docx.Row(Docx.Run("yes")), Docx.Row(Docx.Run("{Else}")),
            Docx.Row(Docx.Run("no")), Docx.Row(Docx.Run("{EndIf}"))));
        var result = Docx.Merge(template, """{"Show":false}""");
        Assert.Equal("no", Docx.Text(Docx.Read(result.Document)));
    }

    [Theory]
    [InlineData("{EndFor}")]
    [InlineData("{Else}")]
    [InlineData("{For:Items}")]
    [InlineData("{If:Flag}\n{Else}\n{Else}\n{EndIf}")]
    public void Analyzer_rejects_unmatched_misplaced_and_duplicate_markers(string markerText)
    {
        var paragraphs = markerText.Split('\n').Select(text => Docx.P(Docx.Run(text))).ToArray();
        using var stream = new MemoryStream(Docx.Create(paragraphs));
        Assert.False(Analyzer().Analyze(stream).Valid);
    }

    [Fact]
    public void Hyperlinks_and_word_fields_are_opaque_and_split_expression_boundaries()
    {
        var hyperlink = new XElement(WordXml.HyperlinkElementName, Docx.Run("{Inside}"));
        var template = Docx.Create(Docx.P(Docx.Run("{Name}"), hyperlink,
            Docx.Field("PAGE", "{Cached}"), Docx.Run(" tail")));

        var result = Docx.Merge(template, """{"Name":"OK","Inside":"NO","Cached":"NO"}""");
        var xml = Docx.Read(result.Document);

        Assert.Equal("OK{Inside}{Cached} tail", Docx.Text(xml));
        Assert.Single(xml.Descendants(WordXml.SimpleFieldElementName));
        Assert.Single(xml.Descendants(WordXml.HyperlinkElementName));
    }

    [Fact]
    public void Drawing_and_text_box_content_remain_opaque_while_ordinary_text_renders()
    {
        var w = WordXml.WordprocessingNamespace;
        XNamespace v = "urn:schemas-microsoft-com:vml";
        var shape = new XElement(w + "pict", new XElement(v + "shape",
            new XElement(v + "textbox", new XElement(w + "txbxContent", Docx.P(Docx.Run("{Hidden}"))))));
        var template = Docx.Create(Docx.P(Docx.Run("{Visible}"), new XElement(w + "r", shape)));
        var result = Docx.Merge(template, """{"Visible":"OK","Hidden":"NO"}""");
        var xml = Docx.Read(result.Document);
        Assert.Equal("OK{Hidden}", Docx.Text(xml));
        Assert.Single(xml.Descendants(w + "pict"));
    }

    [Fact]
    public void Unsupported_drawing_relationship_and_binary_part_are_preserved()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Visible}")));
        var image = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        string relationshipId;
        using var prepared = new MemoryStream();
        prepared.Write(template);
        prepared.Position = 0;
        using (var document = WordprocessingDocument.Open(prepared, true))
        {
            var main = document.MainDocumentPart!;
            var imagePart = main.AddImagePart(ImagePartType.Png);
            using (var target = imagePart.GetStream(FileMode.Create, FileAccess.Write)) target.Write(image);
            relationshipId = main.GetIdOfPart(imagePart);
            XDocument xml;
            using (var source = main.GetStream()) xml = XDocument.Load(source);
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            xml.Descendants(WordXml.ParagraphElementName).First().Add(
                new XElement(WordXml.RunElementName,
                    new XElement(WordXml.DrawingElementName,
                        new XElement(wp + "inline",
                            new XElement(wp + "extent", new XAttribute("cx", 9525), new XAttribute("cy", 9525)),
                            new XElement(wp + "effectExtent", new XAttribute("l", 0), new XAttribute("t", 0),
                                new XAttribute("r", 0), new XAttribute("b", 0)),
                            new XElement(wp + "docPr", new XAttribute("id", 1), new XAttribute("name", "Picture 1")),
                            new XElement(wp + "cNvGraphicFramePr",
                                new XElement(a + "graphicFrameLocks", new XAttribute("noChangeAspect", 1))),
                            new XElement(a + "graphic",
                                new XElement(a + "graphicData",
                                    new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
                                    new XElement(pic + "pic",
                                        new XElement(pic + "nvPicPr",
                                            new XElement(pic + "cNvPr", new XAttribute("id", 0), new XAttribute("name", "image.png")),
                                            new XElement(pic + "cNvPicPr")),
                                        new XElement(pic + "blipFill",
                                            new XElement(a + "blip", new XAttribute(r + "embed", relationshipId)),
                                            new XElement(a + "stretch", new XElement(a + "fillRect"))),
                                        new XElement(pic + "spPr",
                                            new XElement(a + "xfrm",
                                                new XElement(a + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                                                new XElement(a + "ext", new XAttribute("cx", 9525), new XAttribute("cy", 9525))),
                                            new XElement(a + "prstGeom", new XAttribute("prst", "rect"),
                                                new XElement(a + "avLst"))))))))));
            using var output = main.GetStream(FileMode.Create, FileAccess.Write);
            xml.Save(output);
        }

        var result = Docx.Merge(prepared.ToArray(), """{"Visible":"OK"}""");
        using var resultStream = new MemoryStream(result.Document);
        using var generated = WordprocessingDocument.Open(resultStream, false);
        var preservedPart = generated.MainDocumentPart!.GetPartById(relationshipId);
        using var preserved = preservedPart.GetStream();
        using var bytes = new MemoryStream();
        preserved.CopyTo(bytes);
        Assert.Equal(image, bytes.ToArray());
        Assert.Single(Docx.Read(result.Document).Descendants(WordXml.DrawingElementName));
    }

    [Fact]
    public void Expression_text_is_not_joined_across_a_hyperlink_boundary()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Na"),
            new XElement(WordXml.HyperlinkElementName, Docx.Run("link")), Docx.Run("me}")));
        using var stream = new MemoryStream(template);

        var result = Analyzer().Analyze(stream);

        Assert.False(result.Valid);
        Assert.Empty(result.Occurrences);
        Assert.All(result.Errors, error => Assert.Contains(error.Code, new[] { "template.unmatched_brace" }));
    }

    [Theory]
    [InlineData("ins")]
    [InlineData("del")]
    [InlineData("moveFrom")]
    [InlineData("moveTo")]
    [InlineData("pPrChange")]
    public void Analyzer_and_generate_reject_tracked_revisions(string revision)
    {
        var w = WordXml.WordprocessingNamespace;
        XElement paragraph = revision == "pPrChange"
            ? Docx.P(new XElement(w + "pPr", new XElement(w + revision)), Docx.Run("text"))
            : Docx.P(new XElement(w + revision, Docx.Run("text")));
        var template = Docx.Create(paragraph);
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.False(analysis.Valid);
        Assert.Contains(analysis.Errors, error => error.Code == "template.tracked_revision" && error.Location is not null);
        Assert.Throws<TemplateException>(() => Docx.Merge(template, "{}"));
    }

    [Fact]
    public void Analyzer_detects_table_cell_run_and_section_property_revisions_outside_text()
    {
        var w = WordXml.WordprocessingNamespace;
        var table = new XElement(WordXml.TableElementName,
            new XElement(w + "tblPr", new XElement(w + "tblPrChange")),
            new XElement(WordXml.TableGridElementName),
            new XElement(WordXml.TableRowElementName,
                new XElement(w + "trPr", new XElement(w + "trPrChange")),
                new XElement(WordXml.TableCellElementName,
                    new XElement(w + "tcPr", new XElement(w + "tcPrChange")),
                    Docx.P(new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "rPrChange")),
                        new XElement(w + "t", "text"))))));
        var template = Docx.Create(table,
            new XElement(w + "sectPr", new XElement(w + "sectPrChange")));
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.False(analysis.Valid);
        Assert.Equal(5, analysis.Errors.Count(issue => issue.Code == "template.tracked_revision"));
        Assert.All(analysis.Errors, issue => Assert.StartsWith("/word/document.xml", issue.Location));
    }

    [Theory]
    [InlineData("{Text upper}", "abc", "ABC")]
    [InlineData("{Text lower}", "ABC", "abc")]
    [InlineData("{Text caps}", "hELLO wORLD", "HELLO WORLD")]
    [InlineData("{Text firstcap}", "--hello WORLD", "--Hello WORLD")]
    [InlineData("{Text upper}", "true", "TRUE")]
    [InlineData("{Text number:\"0.00\"}", " 1.2e2 ", "120.00")]
    public void Built_in_formatters_apply_approved_scalar_semantics(string expression, string value, string expected)
    {
        var jsonValue = value == "true" ? "true" : System.Text.Json.JsonSerializer.Serialize(value);
        var result = Docx.Merge(Docx.Create(Docx.P(Docx.Run(expression))), "{\"Text\":" + jsonValue + "}",
            new MergeOptions { Culture = CultureInfo.InvariantCulture });

        Assert.Equal(expected, Docx.Text(Docx.Read(result.Document)));
    }

    [Fact]
    public void Text_formatters_use_json_scalar_text_and_configured_culture()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Text upper}|{Number lower}|{Flag caps}|{Text firstcap}")));
        var result = Docx.Merge(template, """{"Text":"istanbul","Number":1.20e2,"Flag":false}""",
            new MergeOptions { Culture = CultureInfo.GetCultureInfo("tr-TR") });
        Assert.Equal("İSTANBUL|1.20e2|False|İstanbul", Docx.Text(Docx.Read(result.Document)));
    }

    [Theory]
    [InlineData("null")]
    [InlineData("\"\"")]
    public void Null_and_empty_string_skip_conversion_for_every_formatter(string jsonLiteral)
    {
        var expressions = new[] { "upper", "lower", "caps", "firstcap", "spellout", "number:\"0.00\"", "date:\"yyyy\"" };
        var template = Docx.Create(Docx.P(Docx.Run(string.Join('|', expressions.Select(name => $"{{Value {name}}}")))));
        var result = Docx.Merge(template, $"{{\"Value\":{jsonLiteral}}}");
        Assert.Equal("||||||", Docx.Text(Docx.Read(result.Document)));
    }

    [Theory]
    [InlineData("1,000")]
    [InlineData("1,5")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    public void Number_rejects_non_invariant_or_non_finite_strings(string value)
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Value number:\"0.00\"}")));
        Assert.Throws<TemplateException>(() => Docx.Merge(template,
            System.Text.Json.JsonSerializer.Serialize(new { Value = value })));
    }

    [Fact]
    public void Number_uses_invariant_input_and_configured_output_culture_without_losing_decimal_precision()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Value number:\"#,##0.000\"}")));
        var result = Docx.Merge(template, """{"Value":" -1.2345e2 "}""",
            new MergeOptions { Culture = CultureInfo.GetCultureInfo("de-DE") });
        Assert.Equal("-123,450", Docx.Text(Docx.Read(result.Document)));
    }

    [Theory]
    [InlineData("\"1e2\"")]
    [InlineData("\" -1.25 \"")]
    [InlineData("12.50")]
    public void Number_and_spellout_share_accepted_decimal_inputs(string literal)
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Value number:\"0.00\"}|{Value spellout}")));
        var result = Docx.Merge(template, $"{{\"Value\":{literal}}}");
        Assert.Contains('|', Docx.Text(Docx.Read(result.Document)));
    }

    [Theory]
    [InlineData("1,000")]
    [InlineData("1,5")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("not-a-number")]
    [InlineData("1e1000")]
    public void Number_and_spellout_share_rejected_numeric_strings(string value)
    {
        foreach (var formatter in new[] { "number:\"0.00\"", "spellout" })
        {
            var template = Docx.Create(Docx.P(Docx.Run($"{{Value {formatter}}}")));
            var error = Assert.Throws<TemplateException>(() => Docx.Merge(template,
                System.Text.Json.JsonSerializer.Serialize(new { Value = value })));
            Assert.Equal("generation.invalid_format", error.Code);
        }
    }

    [Fact]
    public void Date_preserves_explicit_offset_and_rejects_offsetless_datetime()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Value date:\"yyyy-MM-dd HH:mm zzz\"}")));
        var valid = Docx.Merge(template, """{"Value":"2026-01-02T03:04:05.123+03:30"}""");
        Assert.Equal("2026-01-02 03:04 +03:30", Docx.Text(Docx.Read(valid.Document)));

        Assert.Throws<TemplateException>(() => Docx.Merge(template, """{"Value":"2026-01-02T03:04:05"}"""));
    }

    [Theory]
    [InlineData("2026-01-02", "2026-01-02 00:00 +00:00")]
    [InlineData("2026-01-02T03:04:05Z", "2026-01-02 03:04 +00:00")]
    [InlineData("2026-01-02T03:04:05.1234567-04:30", "2026-01-02 03:04 -04:30")]
    public void Date_accepts_only_approved_iso_representations(string value, string expected)
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Value date:\"yyyy-MM-dd HH:mm zzz\"}")));
        var result = Docx.Merge(template, System.Text.Json.JsonSerializer.Serialize(new { Value = value }));
        Assert.Equal(expected, Docx.Text(Docx.Read(result.Document)));
    }

    [Theory]
    [InlineData("01/02/2026")]
    [InlineData("2026-01-02 03:04:05Z")]
    [InlineData("2026-02-30")]
    [InlineData("2026-01-02T03:04:05.12345678Z")]
    public void Date_rejects_ambiguous_or_unsupported_input(string value)
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Value date:\"yyyy\"}")));
        Assert.Throws<TemplateException>(() => Docx.Merge(template,
            System.Text.Json.JsonSerializer.Serialize(new { Value = value })));
    }

    [Fact]
    public void Spellout_uses_shared_numeric_parser_and_boolean_is_a_diagnostic()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Value spellout}")));
        var valid = Docx.Merge(template, """{"Value":" 1.25e1 "}""",
            new MergeOptions { Culture = CultureInfo.GetCultureInfo("en-US") });
        Assert.Equal("Двенадцать рублей 50 копеек", Docx.Text(Docx.Read(valid.Document)));

        var error = Assert.Throws<TemplateException>(() => Docx.Merge(template, """{"Value":true}"""));
        Assert.Equal("generation.invalid_format", error.Code);
    }

    [Theory]
    [InlineData("0", "Ноль рублей 00 копеек")]
    [InlineData("-1.005", "Минус один рубль 01 копейка")]
    [InlineData("2.5", "Два рубля 50 копеек")]
    [InlineData("1000", "Одна тысяча рублей 00 копеек")]
    public void Spellout_preserves_ruble_rounding_and_grammar(string value, string expected)
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Value spellout}")));
        var result = Docx.Merge(template, $"{{\"Value\":{value}}}",
            new MergeOptions { Culture = CultureInfo.GetCultureInfo("tr-TR") });
        Assert.Equal(expected, Docx.Text(Docx.Read(result.Document)));
    }

    [Fact]
    public void Spellout_accepts_the_upper_boundary_and_rejects_the_first_rounded_out_of_range_value()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Value spellout}")));
        var allowed = Docx.Merge(template, """{"Value":999999999999999999.994}""");
        Assert.EndsWith("99 копеек", Docx.Text(Docx.Read(allowed.Document)));
        Assert.Throws<TemplateException>(() => Docx.Merge(template,
            """{"Value":999999999999999999.995}"""));
    }

    [Fact]
    public void Formatter_failures_are_aggregated_without_partial_output()
    {
        var template = Docx.Create(Docx.P(Docx.Run("{Amount number:\"0.00\"}")),
            Docx.P(Docx.Run("{Date date:\"yyyy\"}")), Docx.P(Docx.Run("{Words spellout}")));
        var error = Assert.Throws<TemplateException>(() => Docx.Merge(template,
            """{"Amount":true,"Date":"not-a-date","Words":[]}"""));
        Assert.Equal(3, error.Errors.Count);
        Assert.Equal(new[] { "number", "date", "spellout" }, error.Errors.Select(item => item.Formatter));
        Assert.Equal(new[] { "$.Amount", "$.Date", "$.Words" }, error.Errors.Select(item => item.DataPath));
    }

    private static TemplateAnalyzer Analyzer()
    {
        var reader = new DocxStoryReader();
        return new(reader, new DocxTemplatePlanner(new TemplateLanguageParser()));
    }
}
