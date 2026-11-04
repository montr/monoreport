using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Cli.Tests;
using MonoReport.Core.Templates;
using MonoReport.Docx;
using MonoReport.Tests.Integration;
using Xunit;

namespace MonoReport.Tests.Docx;

public sealed class DocxStructureTests
{
    [Fact]
    public void Nested_for_if_parent_and_root_context_render_recursively()
    {
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{For:Orders}")),
            DocxFixture.P(DocxFixture.Run("{Number}:{$.Company}")),
            DocxFixture.P(DocxFixture.Run("{For:Items}")),
            DocxFixture.P(DocxFixture.Run("{../Number}/{Name}")),
            DocxFixture.P(DocxFixture.Run("{If:Paid}")),
            DocxFixture.P(DocxFixture.Run("+")),
            DocxFixture.P(DocxFixture.Run("{Else}")),
            DocxFixture.P(DocxFixture.Run("-")),
            DocxFixture.P(DocxFixture.Run("{/If}")),
            DocxFixture.P(DocxFixture.Run("{/For}")),
            DocxFixture.P(DocxFixture.Run("{/For}")));

        var result = DocxFixture.Generate(template,
            """{"Company":"ACME","Orders":[{"Number":"1","Items":[{"Name":"A","Paid":true},{"Name":"B","Paid":false}]}]}""");

        Assert.Equal("1:ACME1/A+1/B-", DocxFixture.Text(DocxFixture.Read(result.Document)));
    }

    [Fact]
    public void Paragraph_region_can_repeat_a_complete_table_and_preserve_properties()
    {
        var table = DocxFixture.Table(DocxFixture.Row(DocxFixture.Run("{.}")));
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Values}")), table,
            DocxFixture.P(DocxFixture.Run("{/For}")));

        var result = DocxFixture.Generate(template, """{"Values":["A","B"]}""");
        var xml = DocxFixture.Read(result.Document);

        Assert.Equal("AB", DocxFixture.Text(xml));
        Assert.Equal(2, xml.Descendants(WordXml.TableElementName).Count());
        Assert.All(xml.Descendants(WordXml.TableElementName),
            item => Assert.Single(item.Elements(WordXml.TablePropertiesElementName)));
    }

    [Fact]
    public void Paragraph_if_selects_a_complete_table_and_removes_the_other_branch_and_markers()
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{If:Show}")),
            DocxFixture.Table(DocxFixture.Row(DocxFixture.Run("shown"))), DocxFixture.P(DocxFixture.Run("{Else}")),
            DocxFixture.Table(DocxFixture.Row(DocxFixture.Run("hidden"))), DocxFixture.P(DocxFixture.Run("{/If}")));
        var result = DocxFixture.Generate(template, """{"Show":true}""");
        var xml = DocxFixture.Read(result.Document);
        Assert.Equal("shown", DocxFixture.Text(xml));
        Assert.Single(xml.Descendants(WordXml.TableElementName));
    }

    [Fact]
    public void Row_region_repeats_rows_and_preserves_horizontal_merge_and_list_properties()
    {
        var w = WordXml.WordprocessingNamespace;
        var markerStart = new XElement(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName, DocxFixture.P(DocxFixture.Run("{For:Items}"))));
        var content = new XElement(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName,
                new XElement(w + "tcPr", new XElement(w + "gridSpan", new XAttribute(w + "val", "2"))),
                new XElement(WordXml.ParagraphElementName,
                    new XElement(w + "pPr",
                        new XElement(w + "numPr", new XElement(w + "ilvl", new XAttribute(w + "val", "1")),
                            new XElement(w + "numId", new XAttribute(w + "val", "7")))),
                    DocxFixture.Run("{Name}"))));
        var markerEnd = new XElement(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName, DocxFixture.P(DocxFixture.Run("{/For}"))));
        var template = DocxFixture.Create(DocxFixture.Table(markerStart, content, markerEnd));

        var result = DocxFixture.Generate(template, """{"Items":[{"Name":"A"},{"Name":"B"}]}""");
        var xml = DocxFixture.Read(result.Document);

        Assert.Equal("AB", DocxFixture.Text(xml));
        Assert.Equal(2, xml.Descendants(w + "gridSpan").Count());
        Assert.Equal(2, xml.Descendants(w + "numPr").Count());
    }

    [Fact]
    public void Repeated_numbered_bulleted_and_multilevel_paragraphs_keep_native_numbering_references()
    {
        var w = WordXml.WordprocessingNamespace;

        XElement ListParagraph(string text, int numId, int level)
        {
            return new XElement(WordXml.ParagraphElementName,
                new XElement(w + "pPr", new XElement(w + "numPr",
                        new XElement(w + "ilvl", new XAttribute(w + "val", level)),
                        new XElement(w + "numId", new XAttribute(w + "val", numId))),
                    new XElement(w + "ind", new XAttribute(w + "left", 720 + level * 360))),
                DocxFixture.Run(text));
        }

        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{For:Items}")),
            ListParagraph("{Name}", 3, 0),
            ListParagraph("{Name} child", 3, 1),
            ListParagraph("{Name} bullet", 8, 0),
            DocxFixture.P(DocxFixture.Run("{/For}")));

        var result = DocxFixture.Generate(template, """{"Items":[{"Name":"A"},{"Name":"B"}]}""");
        var xml = DocxFixture.Read(result.Document);
        var properties = xml.Descendants(w + "numPr").ToArray();
        Assert.Equal(6, properties.Length);
        Assert.Equal(new[] {"3", "3", "8", "3", "3", "8"},
            properties.Select(item => (string?) item.Element(w + "numId")?.Attribute(w + "val")));
        Assert.Equal(new[] {"0", "1", "0", "0", "1", "0"},
            properties.Select(item => (string?) item.Element(w + "ilvl")?.Attribute(w + "val")));
        Assert.Equal(6, xml.Descendants(w + "ind").Count());
    }

    [Fact]
    public void Numbering_definition_part_is_preserved_byte_for_byte()
    {
        var w = WordXml.WordprocessingNamespace;
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Items}")),
            new XElement(WordXml.ParagraphElementName,
                new XElement(w + "pPr", new XElement(w + "numPr",
                    new XElement(w + "ilvl", new XAttribute(w + "val", "0")),
                    new XElement(w + "numId", new XAttribute(w + "val", "4")))), DocxFixture.Run("{.}")),
            DocxFixture.P(DocxFixture.Run("{/For}")));
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
        var result = DocxFixture.Generate(preparedBytes, """{"Items":["A","B"]}""");
        Assert.Equal(expected, NumberingXml(result.Document));
    }

    [Fact]
    public void Nested_table_row_loop_operates_only_on_the_inner_table()
    {
        var inner = DocxFixture.Table(DocxFixture.Row(DocxFixture.Run("{For:Values}")),
            DocxFixture.Row(DocxFixture.Run("{.}")), DocxFixture.Row(DocxFixture.Run("{/For}")));
        var outerRow = new XElement(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName, DocxFixture.P(DocxFixture.Run("Outer")), inner,
                DocxFixture.P()));
        var result = DocxFixture.Generate(DocxFixture.Create(DocxFixture.Table(outerRow)), """{"Values":["A","B"]}""");
        var xml = DocxFixture.Read(result.Document);
        Assert.Equal("OuterAB", DocxFixture.Text(xml));
        Assert.Single(xml.Descendants(WordXml.TableElementName).First().Elements(WordXml.TableRowElementName));
        Assert.Equal(2, xml.Descendants(WordXml.TableElementName).Last().Elements(WordXml.TableRowElementName).Count());
    }

    [Theory]
    [InlineData("{For:Items} extra")]
    [InlineData("{For:Items}{/For}")]
    public void Analyzer_rejects_mixed_or_multiple_control_markers(string text)
    {
        using var stream = new MemoryStream(DocxFixture.Create(DocxFixture.P(DocxFixture.Run(text))));
        var analysis = Analyzer().Analyze(stream);
        Assert.False(analysis.Valid);
        Assert.Contains(analysis.Errors, issue => issue.Code == "template.invalid_structure");
    }

    [Theory]
    [InlineData("{For:Items}", "symbol")]
    [InlineData("{If:Show}", "symbol")]
    [InlineData("{Else}", "symbol")]
    [InlineData("{/If}", "symbol")]
    [InlineData("{/For}", "symbol")]
    [InlineData("{For:Items}", "drawing")]
    [InlineData("{If:Show}", "drawing")]
    [InlineData("{Else}", "drawing")]
    [InlineData("{/If}", "drawing")]
    [InlineData("{/For}", "drawing")]
    public void Analyzer_rejects_marker_paragraphs_with_non_text_content(string marker, string content)
    {
        var contaminant = content == "symbol"
            ? new XElement(WordXml.WordprocessingNamespace + "sym",
                new XAttribute(WordXml.WordprocessingNamespace + "font", "Wingdings"),
                new XAttribute(WordXml.WordprocessingNamespace + "char", "F0FC"))
            : Drawing(41);

        XElement Marker(string text)
        {
            var paragraph = DocxFixture.P(DocxFixture.Run(text));
            if (text == marker) paragraph.Add(new XElement(WordXml.RunElementName, contaminant));
            return paragraph;
        }

        var paragraphs = marker switch
        {
            "{For:Items}" => new[] {Marker("{For:Items}"), DocxFixture.P(DocxFixture.Run("{.}")), Marker("{/For}")},
            "{/For}" => new[] {Marker("{For:Items}"), DocxFixture.P(DocxFixture.Run("{.}")), Marker("{/For}")},
            "{If:Show}" => new[] {Marker("{If:Show}"), DocxFixture.P(DocxFixture.Run("yes")), Marker("{/If}")},
            "{Else}" => new[]
            {
                Marker("{If:Show}"), DocxFixture.P(DocxFixture.Run("yes")), Marker("{Else}"),
                DocxFixture.P(DocxFixture.Run("no")), Marker("{/If}")
            },
            _ => new[] {Marker("{If:Show}"), DocxFixture.P(DocxFixture.Run("yes")), Marker("{/If}")}
        };
        var template = DocxFixture.Create(paragraphs);
        DocxFixture.AssertValid(template);
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.False(analysis.Valid);
        Assert.Contains(analysis.Errors, issue => issue.Code == "template.invalid_structure");
    }

    [Theory]
    [InlineData("For")]
    [InlineData("If")]
    public void Formatting_only_marker_paragraphs_remain_valid(string blockKind)
    {
        var w = WordXml.WordprocessingNamespace;

        XElement Marker(string text)
        {
            return new XElement(WordXml.ParagraphElementName,
                new XElement(WordXml.ParagraphPropertiesElementName,
                    new XElement(w + "keepNext")),
                new XElement(WordXml.RunElementName,
                    new XElement(WordXml.RunPropertiesElementName, new XElement(w + "b")),
                    new XElement(WordXml.TextElementName, text)));
        }

        var template = blockKind == "For"
            ? DocxFixture.Create(Marker("{For:Items}"), DocxFixture.P(DocxFixture.Run("{.}")), Marker("{/For}"))
            : DocxFixture.Create(Marker("{If:Show}"), DocxFixture.P(DocxFixture.Run("yes")), Marker("{Else}"),
                DocxFixture.P(DocxFixture.Run("no")), Marker("{/If}"));
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.True(analysis.Valid);
    }

    [Fact]
    public void Analyzer_rejects_vertical_merge_in_row_region()
    {
        var w = WordXml.WordprocessingNamespace;
        var template = DocxFixture.Create(DocxFixture.Table(
            DocxFixture.Row(DocxFixture.Run("{For:Items}")),
            new XElement(WordXml.TableRowElementName, new XElement(WordXml.TableCellElementName,
                new XElement(w + "tcPr", new XElement(w + "vMerge", new XAttribute(w + "val", "restart"))),
                DocxFixture.P(DocxFixture.Run("{Name}")))),
            DocxFixture.Row(DocxFixture.Run("{/For}"))));

        using var stream = new MemoryStream(template);
        var result = Analyzer().Analyze(stream);

        Assert.False(result.Valid);
        Assert.Contains(result.Errors, error => error.Code == "template.invalid_structure");
    }

    [Fact]
    public void Vertical_merge_outside_a_row_region_and_inside_an_intact_cloned_table_is_preserved()
    {
        var w = WordXml.WordprocessingNamespace;

        XElement MergedRow(string value)
        {
            return new XElement(WordXml.TableRowElementName,
                new XElement(WordXml.TableCellElementName,
                    new XElement(w + "tcPr", new XElement(w + "vMerge", new XAttribute(w + "val", value))),
                    DocxFixture.P(DocxFixture.Run("merged"))));
        }

        var rowTemplate = DocxFixture.Create(DocxFixture.Table(MergedRow("restart"), MergedRow("continue"),
            DocxFixture.Row(DocxFixture.Run("{For:Items}")), DocxFixture.Row(DocxFixture.Run("{.}")),
            DocxFixture.Row(DocxFixture.Run("{/For}"))));
        var rowResult = DocxFixture.Generate(rowTemplate, """{"Items":["A","B"]}""");
        Assert.Equal(2, DocxFixture.Read(rowResult.Document).Descendants(w + "vMerge").Count());

        var completeTable = DocxFixture.Table(MergedRow("restart"), MergedRow("continue"));
        var paragraphTemplate = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Items}")), completeTable,
            DocxFixture.P(DocxFixture.Run("{/For}")));
        var paragraphResult = DocxFixture.Generate(paragraphTemplate, """{"Items":[1,2]}""");
        Assert.Equal(4, DocxFixture.Read(paragraphResult.Document).Descendants(w + "vMerge").Count());
    }

    [Fact]
    public void Outer_row_loop_clones_complete_nested_table_with_vertical_merge()
    {
        var w = WordXml.WordprocessingNamespace;

        XElement MergedRow(string value)
        {
            return new XElement(WordXml.TableRowElementName,
                new XElement(WordXml.TableCellElementName,
                    new XElement(WordXml.TableCellPropertiesElementName,
                        new XElement(WordXml.VerticalMergeElementName, new XAttribute(w + "val", value))),
                    DocxFixture.P(DocxFixture.Run("merged"))));
        }

        var nested = DocxFixture.Table(MergedRow("restart"), MergedRow("continue"));
        var body = new XElement(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName, DocxFixture.P(DocxFixture.Run("{Name}")), nested));
        var template = DocxFixture.Create(DocxFixture.Table(
            DocxFixture.Row(DocxFixture.Run("{For:Items}")), body, DocxFixture.Row(DocxFixture.Run("{/For}"))));
        DocxFixture.AssertValid(template);

        var result = DocxFixture.Generate(template, """{"Items":[{"Name":"A"},{"Name":"B"}]}""");
        var xml = DocxFixture.Read(result.Document);
        var outer = xml.Descendants(WordXml.TableElementName).First();

        Assert.Equal(2, outer.Elements(WordXml.TableRowElementName).Count());
        Assert.Equal(2, outer.Descendants(WordXml.TableElementName).Count());
        Assert.Equal(4, outer.Descendants(WordXml.VerticalMergeElementName).Count());
        Assert.Equal("AmergedmergedBmergedmerged", DocxFixture.Text(xml));
    }

    [Fact]
    public void Conditional_row_keeps_only_the_selected_sibling_rows()
    {
        var template = DocxFixture.Create(DocxFixture.Table(DocxFixture.Row(DocxFixture.Run("{If:Show}")),
            DocxFixture.Row(DocxFixture.Run("yes")), DocxFixture.Row(DocxFixture.Run("{Else}")),
            DocxFixture.Row(DocxFixture.Run("no")), DocxFixture.Row(DocxFixture.Run("{/If}"))));
        var result = DocxFixture.Generate(template, """{"Show":false}""");
        Assert.Equal("no", DocxFixture.Text(DocxFixture.Read(result.Document)));
    }

    [Theory]
    [InlineData("{/For}")]
    [InlineData("{Else}")]
    [InlineData("{For:Items}")]
    [InlineData("{If:Flag}\n{Else}\n{Else}\n{/If}")]
    public void Analyzer_rejects_unmatched_misplaced_and_duplicate_markers(string markerText)
    {
        var paragraphs = markerText.Split('\n').Select(text => DocxFixture.P(DocxFixture.Run(text))).ToArray();
        using var stream = new MemoryStream(DocxFixture.Create(paragraphs));
        Assert.False(Analyzer().Analyze(stream).Valid);
    }

    [Fact]
    public void Hyperlinks_and_word_fields_are_opaque_and_split_expression_boundaries()
    {
        var hyperlink = new XElement(WordXml.HyperlinkElementName, DocxFixture.Run("{Inside}"));
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Name}"), hyperlink,
            DocxFixture.Field("PAGE", "{Cached}"), DocxFixture.Run(" tail")));

        var result = DocxFixture.Generate(template, """{"Name":"OK","Inside":"NO","Cached":"NO"}""");
        var xml = DocxFixture.Read(result.Document);

        Assert.Equal("OK{Inside}{Cached} tail", DocxFixture.Text(xml));
        Assert.Single(xml.Descendants(WordXml.SimpleFieldElementName));
        Assert.Single(xml.Descendants(WordXml.HyperlinkElementName));
    }

    [Fact]
    public void Drawing_and_text_box_content_remain_opaque_while_ordinary_text_renders()
    {
        var w = WordXml.WordprocessingNamespace;
        XNamespace v = "urn:schemas-microsoft-com:vml";
        var shape = new XElement(w + "pict", new XElement(v + "shape",
            new XElement(v + "textbox", new XElement(w + "txbxContent", DocxFixture.P(DocxFixture.Run("{Hidden}"))))));
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Visible}"), new XElement(w + "r", shape)));
        var result = DocxFixture.Generate(template, """{"Visible":"OK","Hidden":"NO"}""");
        var xml = DocxFixture.Read(result.Document);
        Assert.Equal("OK{Hidden}", DocxFixture.Text(xml));
        Assert.Single(xml.Descendants(w + "pict"));
    }

    [Fact]
    public void Unsupported_drawing_relationship_and_binary_part_are_preserved()
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Visible}")));
        var image = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        string relationshipId;
        using var prepared = new MemoryStream();
        prepared.Write(template);
        prepared.Position = 0;
        using (var document = WordprocessingDocument.Open(prepared, true))
        {
            var main = document.MainDocumentPart!;
            var imagePart = main.AddImagePart(ImagePartType.Png);
            using (var target = imagePart.GetStream(FileMode.Create, FileAccess.Write))
            {
                target.Write(image);
            }

            relationshipId = main.GetIdOfPart(imagePart);
            XDocument xml;
            using (var source = main.GetStream())
            {
                xml = XDocument.Load(source);
            }

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
                                            new XElement(pic + "cNvPr", new XAttribute("id", 0),
                                                new XAttribute("name", "image.png")),
                                            new XElement(pic + "cNvPicPr")),
                                        new XElement(pic + "blipFill",
                                            new XElement(a + "blip", new XAttribute(r + "embed", relationshipId)),
                                            new XElement(a + "stretch", new XElement(a + "fillRect"))),
                                        new XElement(pic + "spPr",
                                            new XElement(a + "xfrm",
                                                new XElement(a + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                                                new XElement(a + "ext", new XAttribute("cx", 9525),
                                                    new XAttribute("cy", 9525))),
                                            new XElement(a + "prstGeom", new XAttribute("prst", "rect"),
                                                new XElement(a + "avLst"))))))))));
            using var output = main.GetStream(FileMode.Create, FileAccess.Write);
            xml.Save(output);
        }

        var result = DocxFixture.Generate(prepared.ToArray(), """{"Visible":"OK"}""");
        using var resultStream = new MemoryStream(result.Document);
        using var generated = WordprocessingDocument.Open(resultStream, false);
        var preservedPart = generated.MainDocumentPart!.GetPartById(relationshipId);
        using var preserved = preservedPart.GetStream();
        using var bytes = new MemoryStream();
        preserved.CopyTo(bytes);
        Assert.Equal(image, bytes.ToArray());
        Assert.Single(DocxFixture.Read(result.Document).Descendants(WordXml.DrawingElementName));
    }

    [Fact]
    public void Expression_text_is_not_joined_across_a_hyperlink_boundary()
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Na"),
            new XElement(WordXml.HyperlinkElementName, DocxFixture.Run("link")), DocxFixture.Run("me}")));
        using var stream = new MemoryStream(template);

        var result = Analyzer().Analyze(stream);

        Assert.False(result.Valid);
        Assert.Empty(result.Occurrences);
        Assert.All(result.Errors, error => Assert.Contains(error.Code, new[] {"template.unmatched_brace"}));
    }

    [Theory]
    [InlineData("symbol")]
    [InlineData("tab")]
    [InlineData("break")]
    public void Expression_text_is_not_joined_across_non_text_run_content(string kind)
    {
        var w = WordXml.WordprocessingNamespace;
        var boundary = kind switch
        {
            "symbol" => new XElement(w + "sym", new XAttribute(w + "font", "Wingdings"),
                new XAttribute(w + "char", "F0FC")),
            "tab" => new XElement(WordXml.TabElementName),
            _ => new XElement(WordXml.BreakElementName)
        };
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Na"),
            new XElement(WordXml.RunElementName, boundary), DocxFixture.Run("me}")));
        DocxFixture.AssertValid(template);
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.False(analysis.Valid);
        Assert.Empty(analysis.Occurrences);
        Assert.All(analysis.Errors, issue => Assert.Equal("template.unmatched_brace", issue.Code));
    }

    [Theory]
    [InlineData("hyperlink")]
    [InlineData("field")]
    [InlineData("drawing")]
    [InlineData("picture")]
    public void Analyzer_rejects_marker_rows_with_additional_opaque_visible_content(string kind)
    {
        var start = new XElement(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName, DocxFixture.P(DocxFixture.Run("{For:Items}"))),
            new XElement(WordXml.TableCellElementName, DocxFixture.P(OpaqueContent(kind, true))));
        var template = DocxFixture.Create(DocxFixture.Table(start, DocxFixture.Row(DocxFixture.Run("{.}")),
            DocxFixture.Row(DocxFixture.Run("{/For}"))));
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.False(analysis.Valid);
        Assert.Contains(analysis.Errors, issue => issue.Code == "template.invalid_structure");
    }

    [Theory]
    [InlineData("hyperlink")]
    [InlineData("field")]
    public void Analyzer_rejects_marker_rows_with_visible_symbols_inside_opaque_content(string kind)
    {
        var w = WordXml.WordprocessingNamespace;
        var symbol = new XElement(w + "sym", new XAttribute(w + "font", "Wingdings"),
            new XAttribute(w + "char", "F0FC"));
        var opaque = kind == "hyperlink"
            ? new XElement(WordXml.HyperlinkElementName, new XElement(WordXml.RunElementName, symbol))
            : new XElement(WordXml.SimpleFieldElementName, new XAttribute(WordXml.InstructionAttributeName, "PAGE"),
                new XElement(WordXml.RunElementName, symbol));
        var start = new XElement(WordXml.TableRowElementName,
            new XElement(WordXml.TableCellElementName, DocxFixture.P(DocxFixture.Run("{For:Items}"))),
            new XElement(WordXml.TableCellElementName, DocxFixture.P(opaque)));
        var template = DocxFixture.Create(DocxFixture.Table(start, DocxFixture.Row(DocxFixture.Run("{.}")),
            DocxFixture.Row(DocxFixture.Run("{/For}"))));
        DocxFixture.AssertValid(template);
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.False(analysis.Valid);
        Assert.Contains(analysis.Errors, issue => issue.Code == "template.invalid_structure");
    }

    [Theory]
    [InlineData("hyperlink")]
    [InlineData("field")]
    [InlineData("drawing")]
    [InlineData("picture")]
    public void Empty_opaque_content_splits_expression_text_groups(string kind)
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Na"), OpaqueContent(kind, false),
            DocxFixture.Run("me}")));
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.False(analysis.Valid);
        Assert.Empty(analysis.Occurrences);
        Assert.All(analysis.Errors, issue => Assert.Equal("template.unmatched_brace", issue.Code));
    }

    [Fact]
    public void Complex_word_field_state_remains_opaque_across_paragraphs()
    {
        var template = DocxFixture.Create(
            DocxFixture.P(new XElement(WordXml.RunElementName, DocxFixture.Character(WordXml.FieldBeginValue),
                new XElement(WordXml.InstructionTextElementName, "PAGE"))),
            DocxFixture.P(DocxFixture.Run("{Hidden}"), new XElement(WordXml.RunElementName,
                DocxFixture.Character(WordXml.FieldEndValue))),
            DocxFixture.P(DocxFixture.Run("{Visible}")));
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.True(analysis.Valid);
        var occurrence = Assert.Single(analysis.Occurrences);
        Assert.Equal("{Visible}", occurrence.Expression);
        var result = DocxFixture.Generate(template, """{"Hidden":"NO","Visible":"OK"}""");
        Assert.Equal("{Hidden}OK", DocxFixture.Text(DocxFixture.Read(result.Document)));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Complex_word_field_state_crosses_hyperlink_wrappers_without_processing_hyperlink_text(
        bool beginInsideHyperlink)
    {
        var w = WordXml.WordprocessingNamespace;

        XElement Hyperlink(params object[] content)
        {
            return new XElement(WordXml.HyperlinkElementName,
                new XAttribute(WordXml.AnchorAttributeName, "target"), content);
        }

        var field = beginInsideHyperlink
            ? DocxFixture.P(Hyperlink(new XElement(WordXml.RunElementName,
                    DocxFixture.Character(WordXml.FieldBeginValue),
                    new XElement(WordXml.InstructionTextElementName, " PAGE "))),
                DocxFixture.Run("{Hidden}"),
                new XElement(WordXml.RunElementName, DocxFixture.Character(WordXml.FieldEndValue)),
                DocxFixture.Run("{Visible}"))
            : DocxFixture.P(new XElement(WordXml.RunElementName, DocxFixture.Character(WordXml.FieldBeginValue),
                    new XElement(WordXml.InstructionTextElementName, " PAGE ")),
                Hyperlink(DocxFixture.Run("{Hidden}"),
                    new XElement(WordXml.RunElementName, DocxFixture.Character(WordXml.FieldEndValue))),
                DocxFixture.Run("{Visible}"));
        var template = DocxFixture.Create(
            new XElement(WordXml.ParagraphElementName,
                new XElement(w + "bookmarkStart", new XAttribute(w + "id", 0),
                    new XAttribute(w + "name", "target")),
                DocxFixture.Run("target"), new XElement(w + "bookmarkEnd", new XAttribute(w + "id", 0))),
            field);
        DocxFixture.AssertValid(template);
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.True(analysis.Valid);
        var occurrence = Assert.Single(analysis.Occurrences);
        Assert.Equal("{Visible}", occurrence.Expression);
        var result = DocxFixture.Generate(template, """{"Hidden":"NO","Visible":"OK"}""");
        Assert.Equal("target{Hidden}OK", DocxFixture.Text(DocxFixture.Read(result.Document)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Complete_bookmark_pairs_on_operation_markers_are_rejected(bool rowMarker)
    {
        var w = WordXml.WordprocessingNamespace;

        XElement MarkerParagraph(string marker, int id)
        {
            return new XElement(WordXml.ParagraphElementName,
                new XElement(w + "bookmarkStart", new XAttribute(w + "id", id),
                    new XAttribute(w + "name", "marker" + id)),
                DocxFixture.Run(marker), new XElement(w + "bookmarkEnd", new XAttribute(w + "id", id)));
        }

        var template = rowMarker
            ? DocxFixture.Create(DocxFixture.Table(
                new XElement(WordXml.TableRowElementName,
                    new XElement(WordXml.TableCellElementName, MarkerParagraph("{For:Items}", 4))),
                DocxFixture.Row(DocxFixture.Run("{.}")),
                new XElement(WordXml.TableRowElementName,
                    new XElement(WordXml.TableCellElementName, MarkerParagraph("{/For}", 5)))))
            : DocxFixture.Create(MarkerParagraph("{For:Items}", 4), DocxFixture.P(DocxFixture.Run("{.}")),
                MarkerParagraph("{/For}", 5));
        DocxFixture.AssertValid(template);
        DocxIdentityTests.Reject(template, """{"Items":["A","B"]}""", "template.invalid_structure");
    }

    [Fact]
    public void Repeated_permission_ranges_receive_matching_local_ids()
    {
        var w = WordXml.WordprocessingNamespace;
        var body = new XElement(WordXml.ParagraphElementName,
            new XElement(w + "permStart", new XAttribute(w + "id", 9), new XAttribute(w + "edGrp", "everyone")),
            DocxFixture.Run("{Name}"),
            new XElement(w + "permEnd", new XAttribute(w + "id", 9)));
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Items}")), body,
            DocxFixture.P(DocxFixture.Run("{/For}")));
        DocxFixture.AssertValid(template);

        var result = DocxFixture.Generate(template, """{"Items":[{"Name":"A"},{"Name":"B"}]}""");
        var xml = DocxFixture.Read(result.Document);
        var permissionStarts = xml.Descendants(w + "permStart").ToArray();
        var permissionEnds = xml.Descendants(w + "permEnd").ToArray();
        Assert.Equal(permissionStarts.Select(item => (string) item.Attribute(w + "id")!),
            permissionEnds.Select(item => (string) item.Attribute(w + "id")!));
        Assert.Equal(2, permissionStarts.Select(item => (string) item.Attribute(w + "id")!).Distinct().Count());
    }

    [Fact]
    public void Repeated_comment_ranges_are_rejected_before_cloning()
    {
        var w = WordXml.WordprocessingNamespace;
        var body = new XElement(WordXml.ParagraphElementName,
            new XElement(w + "commentRangeStart", new XAttribute(w + "id", 0)),
            DocxFixture.Run("{Name}"),
            new XElement(w + "commentRangeEnd", new XAttribute(w + "id", 0)),
            new XElement(WordXml.RunElementName,
                new XElement(w + "commentReference", new XAttribute(w + "id", 0))));
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Items}")), body,
            DocxFixture.P(DocxFixture.Run("{/For}")));
        using var prepared = new MemoryStream();
        prepared.Write(template);
        prepared.Position = 0;
        using (var document = WordprocessingDocument.Open(prepared, true))
        {
            var comments = document.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>();
            using var target = comments.GetStream(FileMode.Create, FileAccess.Write);
            new XDocument(new XElement(w + "comments",
                    new XElement(w + "comment", new XAttribute(w + "id", 0), new XAttribute(w + "author", "A"),
                        new XAttribute(w + "initials", "A"),
                        new XAttribute(w + "date", "2026-10-18T00:00:00Z"), DocxFixture.P(DocxFixture.Run("comment")))))
                .Save(target);
        }

        var preparedBytes = prepared.ToArray();
        DocxFixture.AssertValid(preparedBytes);

        DocxIdentityTests.Reject(preparedBytes, """{"Items":[{"Name":"A"},{"Name":"B"}]}""",
            "template.invalid_structure");
    }

    [Fact]
    public void Repeated_drawings_receive_unique_document_property_ids()
    {
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{For:Items}")),
            DocxFixture.P(new XElement(WordXml.RunElementName, Drawing(7))),
            DocxFixture.P(DocxFixture.Run("{/For}")));
        var image = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        string relationshipId;
        using var prepared = new MemoryStream();
        prepared.Write(template);
        prepared.Position = 0;
        using (var document = WordprocessingDocument.Open(prepared, true))
        {
            var main = document.MainDocumentPart!;
            var imagePart = main.AddImagePart(ImagePartType.Png);
            using (var target = imagePart.GetStream(FileMode.Create, FileAccess.Write))
            {
                target.Write(image);
            }

            relationshipId = main.GetIdOfPart(imagePart);
            XDocument preparedXml;
            using (var source = main.GetStream())
            {
                preparedXml = XDocument.Load(source);
            }

            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            preparedXml.Descendants(a + "blip").Single().SetAttributeValue(WordXml.RelationshipsNamespace + "embed",
                relationshipId);
            using var output = main.GetStream(FileMode.Create, FileAccess.Write);
            preparedXml.Save(output);
        }

        var result = DocxFixture.Generate(prepared.ToArray(), """{"Items":[1,2]}""");
        var xml = DocxFixture.Read(result.Document);
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var ids = xml.Descendants(wp + "docPr").Select(item => (uint) item.Attribute("id")!).ToArray();

        Assert.Equal(2, ids.Length);
        Assert.Equal(ids.Length, ids.Distinct().Count());
        Assert.All(xml.Descendants(drawing + "blip"), blip =>
            Assert.Equal(relationshipId, (string?) blip.Attribute(WordXml.RelationshipsNamespace + "embed")));
        using var resultStream = new MemoryStream(result.Document);
        using var generated = WordprocessingDocument.Open(resultStream, false);
        using var preserved = generated.MainDocumentPart!.GetPartById(relationshipId).GetStream();
        using var bytes = new MemoryStream();
        preserved.CopyTo(bytes);
        Assert.Equal(image, bytes.ToArray());
    }

    [Fact]
    public void Repeated_drawings_receive_ids_unique_across_main_and_header_stories()
    {
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{For:Items}")),
            DocxFixture.P(new XElement(WordXml.RunElementName, Drawing(7))),
            DocxFixture.P(DocxFixture.Run("{/For}")));
        using var prepared = new MemoryStream();
        prepared.Write(template);
        prepared.Position = 0;
        using (var document = WordprocessingDocument.Open(prepared, true))
        {
            var header = document.MainDocumentPart!.AddNewPart<HeaderPart>();
            using var target = header.GetStream(FileMode.Create, FileAccess.Write);
            new XDocument(new XElement(WordXml.HeaderElementName,
                DocxFixture.P(new XElement(WordXml.RunElementName, Drawing(1))))).Save(target);
        }

        var result = DocxFixture.Generate(prepared.ToArray(), """{"Items":[1,2]}""");
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        var stories = DocxFixture.ReadStories(result.Document);
        var properties = stories.Values.SelectMany(story => story.Descendants(wp + "docPr")).ToArray();
        var ids = properties.Select(item => (uint) item.Attribute("id")!).ToArray();

        Assert.Equal(3, ids.Length);
        Assert.Equal(ids.Length, ids.Distinct().Count());
        Assert.Contains(properties, item => (string?) item.Attribute("name") == "Picture 1");
        Assert.Equal(2, properties.Count(item => (string?) item.Attribute("name") == "Picture 7"));
    }

    [Fact]
    public void Repeated_drawings_reserve_ids_from_comments_and_glossary_parts()
    {
        var w = WordXml.WordprocessingNamespace;
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{For:Items}")),
            DocxFixture.P(new XElement(WordXml.RunElementName, Drawing(1))),
            DocxFixture.P(DocxFixture.Run("{/For}")));
        using var prepared = new MemoryStream();
        prepared.Write(template);
        prepared.Position = 0;
        using (var document = WordprocessingDocument.Open(prepared, true))
        {
            var main = document.MainDocumentPart!;
            var comments = main.AddNewPart<WordprocessingCommentsPart>();
            using (var target = comments.GetStream(FileMode.Create, FileAccess.Write))
            {
                new XDocument(new XElement(w + "comments",
                    new XElement(w + "comment", new XAttribute(w + "id", 0), new XAttribute(w + "author", "A"),
                        new XAttribute(w + "initials", "A"),
                        new XAttribute(w + "date", "2026-10-18T00:00:00Z"),
                        DocxFixture.P(new XElement(WordXml.RunElementName, Drawing(2)))))).Save(target);
            }

            var glossary = main.AddNewPart<GlossaryDocumentPart>();
            using var glossaryTarget = glossary.GetStream(FileMode.Create, FileAccess.Write);
            new XDocument(new XElement(w + "glossaryDocument",
                new XElement(w + "docParts",
                    new XElement(w + "docPart",
                        new XElement(w + "docPartPr", new XElement(w + "name", new XAttribute(w + "val", "Entry"))),
                        new XElement(w + "docPartBody",
                            DocxFixture.P(new XElement(WordXml.RunElementName, Drawing(3)))))))).Save(glossaryTarget);
        }

        var preparedBytes = prepared.ToArray();
        DocxFixture.AssertValid(preparedBytes);

        var result = DocxFixture.Generate(preparedBytes, """{"Items":[1,2]}""");
        using var resultStream = new MemoryStream(result.Document);
        using var generated = WordprocessingDocument.Open(resultStream, false);
        var ids = generated.MainDocumentPart!.Document!.Descendants<DocProperties>()
            .Select(item => item.Id!.Value).ToArray();

        Assert.Equal(2, ids.Length);
        Assert.All(ids, id => Assert.DoesNotContain(id, new uint[] {1, 2, 3}));
        Assert.Equal(2, ids.Distinct().Count());
    }

    [Fact]
    public void Analyzer_diagnostics_follow_source_order_across_processing_phases()
    {
        var template = DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{Value}")),
            DocxFixture.P(DocxFixture.Run("{For:Value}")),
            DocxFixture.P(DocxFixture.Run("{/For}")),
            DocxFixture.P(DocxFixture.Run("{Broken")));
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.False(analysis.Valid);
        Assert.Equal(new[] {1, 2, 4}, analysis.Errors.Select(issue =>
            int.Parse(issue.Location!.Split("paragraph ")[1].Split(',')[0])));
        Assert.Equal(
            new[] {"template.requirement_conflict", "template.requirement_conflict", "template.unmatched_brace"},
            analysis.Errors.Select(issue => issue.Code));

        using var sameParagraph = new MemoryStream(DocxFixture.Create(
            DocxFixture.P(DocxFixture.Run("{Value}{For:Value}{/For}{Broken"))));
        var sameParagraphAnalysis = Analyzer().Analyze(sameParagraph);
        Assert.Equal(new[]
        {
            "template.requirement_conflict",
            "template.invalid_structure",
            "template.requirement_conflict",
            "template.invalid_structure",
            "template.unmatched_brace"
        }, sameParagraphAnalysis.Errors.Select(issue => issue.Code));
    }

    [Fact]
    public void Analyzer_anchors_parser_diagnostics_to_the_actual_text_node_after_a_formatting_revision()
    {
        var w = WordXml.WordprocessingNamespace;
        var template = DocxFixture.Create(DocxFixture.P(
            DocxFixture.Run("prefix"),
            new XElement(WordXml.RunElementName,
                new XElement(WordXml.RunPropertiesElementName,
                    new XElement(w + "rPrChange", new XAttribute(w + "id", 1),
                        new XAttribute(w + "author", "Reviewer"),
                        new XAttribute(w + "date", "2026-10-18T00:00:00Z"),
                        new XElement(WordXml.RunPropertiesElementName))),
                new XElement(WordXml.TextElementName, "middle")),
            DocxFixture.Run("{Broken")));
        DocxFixture.AssertValid(template);
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.Equal(new[] {"template.tracked_revision", "template.unmatched_brace"},
            analysis.Errors.Select(issue => issue.Code));
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
        var paragraph = revision == "pPrChange"
            ? DocxFixture.P(new XElement(w + "pPr", new XElement(w + revision)), DocxFixture.Run("text"))
            : DocxFixture.P(new XElement(w + revision, DocxFixture.Run("text")));
        var template = DocxFixture.Create(paragraph);
        using var stream = new MemoryStream(template);

        var analysis = Analyzer().Analyze(stream);

        Assert.False(analysis.Valid);
        Assert.Contains(analysis.Errors,
            error => error.Code == "template.tracked_revision" && error.Location is not null);
        Assert.Throws<TemplateException>(() => DocxFixture.Generate(template, "{}"));
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
                    DocxFixture.P(new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "rPrChange")),
                        new XElement(w + "t", "text"))))));
        var template = DocxFixture.Create(table,
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
    [InlineData("{Text number \"0.00\"}", " 1.2e2 ", "120.00")]
    public void Built_in_formatters_apply_approved_scalar_semantics(string expression, string value, string expected)
    {
        var jsonValue = value == "true" ? "true" : JsonSerializer.Serialize(value);
        var result = DocxFixture.Generate(DocxFixture.Create(DocxFixture.P(DocxFixture.Run(expression))),
            "{\"Text\":" + jsonValue + "}",
            new GenerationOptions {Culture = CultureInfo.InvariantCulture});

        Assert.Equal(expected, DocxFixture.Text(DocxFixture.Read(result.Document)));
    }

    [Fact]
    public void Text_formatters_use_json_scalar_text_and_configured_culture()
    {
        var template =
            DocxFixture.Create(
                DocxFixture.P(DocxFixture.Run("{Text upper}|{Number lower}|{Flag caps}|{Text firstcap}")));
        var result = DocxFixture.Generate(template, """{"Text":"istanbul","Number":1.20e2,"Flag":false}""",
            new GenerationOptions {Culture = CultureInfo.GetCultureInfo("tr-TR")});
        Assert.Equal("İSTANBUL|1.20e2|False|İstanbul", DocxFixture.Text(DocxFixture.Read(result.Document)));
    }

    [Theory]
    [InlineData("null")]
    [InlineData("\"\"")]
    public void Null_and_empty_string_skip_conversion_for_every_formatter(string jsonLiteral)
    {
        var expressions = new[] {"upper", "lower", "caps", "firstcap", "spellout", "number \"0.00\"", "date \"yyyy\""};
        var template =
            DocxFixture.Create(
                DocxFixture.P(DocxFixture.Run(string.Join('|', expressions.Select(name => $"{{Value {name}}}")))));
        var result = DocxFixture.Generate(template, $"{{\"Value\":{jsonLiteral}}}");
        Assert.Equal("||||||", DocxFixture.Text(DocxFixture.Read(result.Document)));
    }

    [Theory]
    [InlineData("1,000")]
    [InlineData("1,5")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    public void Number_rejects_non_invariant_or_non_finite_strings(string value)
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Value number \"0.00\"}")));
        Assert.Throws<TemplateException>(() => DocxFixture.Generate(template,
            JsonSerializer.Serialize(new {Value = value})));
    }

    [Fact]
    public void Number_uses_invariant_input_and_configured_output_culture_without_losing_decimal_precision()
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Value number \"#,##0.000\"}")));
        var result = DocxFixture.Generate(template, """{"Value":" -1.2345e2 "}""",
            new GenerationOptions {Culture = CultureInfo.GetCultureInfo("de-DE")});
        Assert.Equal("-123,450", DocxFixture.Text(DocxFixture.Read(result.Document)));
    }

    [Theory]
    [InlineData("\"1e2\"", "100.00|One hundred rubles 00 copecks")]
    [InlineData("\" -1.25 \"", "-1.25|Minus one ruble 25 copecks")]
    [InlineData("12.50", "12.50|Twelve rubles 50 copecks")]
    public void Number_and_spellout_render_complete_output_for_accepted_decimal_inputs(string literal,
        string expected)
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run(
            "{Value number \"0.00\"}|{Value spellout mode=\"money\" currency=\"RUB\"}")));
        var result = DocxFixture.Generate(template, $"{{\"Value\":{literal}}}",
            new GenerationOptions {Culture = CultureInfo.GetCultureInfo("en-US")});
        Assert.Equal(expected, DocxFixture.Text(DocxFixture.Read(result.Document)));
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
        foreach (var formatter in new[] {"number \"0.00\"", "spellout mode=\"money\" currency=\"RUB\""})
        {
            var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run($"{{Value {formatter}}}")));
            var error = Assert.Throws<TemplateException>(() => DocxFixture.Generate(template,
                JsonSerializer.Serialize(new {Value = value})));
            Assert.Equal("generation.invalid_format", error.Code);
        }
    }

    [Fact]
    public void Date_preserves_explicit_offset_and_rejects_offsetless_datetime()
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Value date \"yyyy-MM-dd HH:mm zzz\"}")));
        var valid = DocxFixture.Generate(template, """{"Value":"2026-01-02T03:04:05.123+03:30"}""");
        Assert.Equal("2026-01-02 03:04 +03:30", DocxFixture.Text(DocxFixture.Read(valid.Document)));

        Assert.Throws<TemplateException>(() => DocxFixture.Generate(template, """{"Value":"2026-01-02T03:04:05"}"""));
    }

    [Theory]
    [InlineData("2026-01-02", "2026-01-02 00:00 +00:00")]
    [InlineData("2026-01-02T03:04:05Z", "2026-01-02 03:04 +00:00")]
    [InlineData("2026-01-02T03:04:05.1234567-04:30", "2026-01-02 03:04 -04:30")]
    public void Date_accepts_only_approved_iso_representations(string value, string expected)
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Value date \"yyyy-MM-dd HH:mm zzz\"}")));
        var result = DocxFixture.Generate(template, JsonSerializer.Serialize(new {Value = value}));
        Assert.Equal(expected, DocxFixture.Text(DocxFixture.Read(result.Document)));
    }

    [Theory]
    [InlineData("01/02/2026")]
    [InlineData("2026-01-02 03:04:05Z")]
    [InlineData("2026-02-30")]
    [InlineData("2026-01-02T03:04:05.12345678Z")]
    public void Date_rejects_ambiguous_or_unsupported_input(string value)
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Value date \"yyyy\"}")));
        Assert.Throws<TemplateException>(() => DocxFixture.Generate(template,
            JsonSerializer.Serialize(new {Value = value})));
    }

    [Fact]
    public void Formatter_failures_are_aggregated_without_partial_output()
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Amount number \"0.00\"}")),
            DocxFixture.P(DocxFixture.Run("{Date date \"yyyy\"}")), DocxFixture.P(DocxFixture.Run("{Words spellout}")));
        var error = Assert.Throws<TemplateException>(() => DocxFixture.Generate(template,
            """{"Amount":true,"Date":"not-a-date","Words":[]}"""));
        Assert.Equal(3, error.Errors.Count);
        Assert.Equal(new[] {"number", "date", "spellout"}, error.Errors.Select(item => item.Formatter));
        Assert.Equal(new[] {"$.Amount", "$.Date", "$.Words"}, error.Errors.Select(item => item.DataPath));
    }

    private static TemplateAnalyzer Analyzer()
    {
        return TestComposition.Analyzer;
    }

    private static XElement OpaqueContent(string kind, bool visible)
    {
        var text = visible ? "visible" : string.Empty;
        XNamespace v = "urn:schemas-microsoft-com:vml";
        return kind switch
        {
            "hyperlink" => new XElement(WordXml.HyperlinkElementName, visible ? DocxFixture.Run(text) : null),
            "field" => DocxFixture.Field("PAGE", text),
            "drawing" => Drawing(1),
            "picture" => new XElement(WordXml.PictureElementName,
                new XElement(v + "shape",
                    visible ? new XElement(v + "textbox", DocxFixture.P(DocxFixture.Run(text))) : null)),
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
    }

    private static XElement Drawing(uint id)
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        return new XElement(WordXml.DrawingElementName,
            new XElement(wp + "inline",
                new XElement(wp + "extent", new XAttribute("cx", 9525), new XAttribute("cy", 9525)),
                new XElement(wp + "effectExtent", new XAttribute("l", 0), new XAttribute("t", 0),
                    new XAttribute("r", 0), new XAttribute("b", 0)),
                new XElement(wp + "docPr", new XAttribute("id", id), new XAttribute("name", $"Picture {id}")),
                new XElement(wp + "cNvGraphicFramePr",
                    new XElement(a + "graphicFrameLocks", new XAttribute("noChangeAspect", 1))),
                new XElement(a + "graphic",
                    new XElement(a + "graphicData",
                        new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
                        new XElement(pic + "pic",
                            new XElement(pic + "nvPicPr",
                                new XElement(pic + "cNvPr", new XAttribute("id", 0),
                                    new XAttribute("name", "image.png")),
                                new XElement(pic + "cNvPicPr")),
                            new XElement(pic + "blipFill", new XElement(a + "blip"),
                                new XElement(a + "stretch", new XElement(a + "fillRect"))),
                            new XElement(pic + "spPr",
                                new XElement(a + "xfrm",
                                    new XElement(a + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                                    new XElement(a + "ext", new XAttribute("cx", 9525), new XAttribute("cy", 9525))),
                                new XElement(a + "prstGeom", new XAttribute("prst", "rect"),
                                    new XElement(a + "avLst"))))))));
    }
}