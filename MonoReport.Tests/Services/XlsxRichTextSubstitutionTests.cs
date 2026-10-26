using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class XlsxRichTextSubstitutionTests
{
    private readonly XlsxFixture _fixture = new();
    private static readonly XNamespace S = XlsxPackageCompatibilityTests.S;

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Cross_run_fields_escapes_and_decoded_text_keep_surrounding_and_opening_brace_properties(bool shared)
    {
        var source = Rich(shared, new Run(new RunProperties(new Italic()), new Text(" \n漢😀 ")),
            new Run(new RunProperties(new Bold(), new Italic { Val = false }, new Color { Theme = 6, Tint = 0.4 }), new Text("{Na")),
            new Run(new RunProperties(new Color { Rgb = "FFFF0000" }), new Text("me}|{Empty}{Other} {")),
            new Run(new RunProperties(new Underline()), new Text("{literal}} _x0041_ _x005F_x0042_\n ")));
        using var services = _fixture.Services();
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(new MemoryStream(source));
        Assert.True(analysis.Valid);
        Assert.Equal(new[] { "Name", "Empty", "Other" }, analysis.Requirements.Select(requirement => requirement.Name));
        var result = _fixture.Merge(source, "{\"Name\":\"Longer\",\"Empty\":\"\",\"Other\":\"=X{Y}\"}");
        const string expected = " \n漢😀 Longer|=X{Y} {literal} A _x0042_\n ";
        Assert.Equal(expected, _fixture.ReadCellText(result.Document, "A1"));
        using var document = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single();
        Assert.Null(cell.CellFormula); Assert.Equal(1U, cell.StyleIndex!.Value);
        Assert.Null(document.WorkbookPart.SharedStringTablePart);
        var properties = CharacterProperties(result.Document, "A1");
        Assert.NotNull(properties[0]!.Element(S + "i"));
        var replacement = properties[expected.IndexOf("Longer", StringComparison.Ordinal)]!;
        Assert.NotNull(replacement.Element(S + "b")); Assert.Equal("0", (string?)replacement.Element(S + "i")!.Attribute("val"));
        Assert.Equal("6", (string?)replacement.Element(S + "color")!.Attribute("theme"));
        Assert.Equal("0.4", (string?)replacement.Element(S + "color")!.Attribute("tint"));
        Assert.All(properties.Skip(expected.IndexOf("Longer", StringComparison.Ordinal)).Take(6), item => Assert.True(XNode.DeepEquals(replacement, item)));
        Assert.Equal("FFFF0000", (string?)properties[expected.IndexOf("{literal}", StringComparison.Ordinal)]!.Element(S + "color")!.Attribute("rgb"));
        Assert.NotNull(properties[^1]!.Element(S + "u"));
    }

    [Theory]
    [InlineData("{Name}", "Ada", null)]
    [InlineData("{Name upper}", "ADA", null)]
    [InlineData("{Missing}", "", DiagnosticCodes.GenerationInvalidData)]
    [InlineData("{Name", null, DiagnosticCodes.TemplateUnmatchedBrace)]
    [InlineData("{../Name}", null, DiagnosticCodes.TemplateInvalidContext)]
    [InlineData("{For:Items}", null, DiagnosticCodes.XlsxUnsupportedFeature)]
    public void Plain_and_rich_sources_have_equivalent_analysis_and_real_language_diagnostics(string text, string? expected, string? errorCode)
    {
        var plain = _fixture.Create((_, part) => part.Worksheet!.Descendants<Text>().Single().Text = text);
        var rich = Rich(false, new Run(new RunProperties(new Bold()), new Text(text)));
        using var services = _fixture.Services();
        var analyzer = services.GetRequiredService<TemplateAnalyzer>();
        var plainAnalysis = analyzer.Analyze(new MemoryStream(plain)); var richAnalysis = analyzer.Analyze(new MemoryStream(rich));
        Assert.Equal(System.Text.Json.JsonSerializer.Serialize(plainAnalysis), System.Text.Json.JsonSerializer.Serialize(richAnalysis));
        if (!richAnalysis.Valid) Assert.Contains(richAnalysis.Errors, issue => issue.Code == errorCode);
        else if (errorCode is not null)
            Assert.Equal(errorCode, Assert.Throws<TemplateException>(() => _fixture.Merge(rich, "{}", new MergeOptions { MissingValues = MissingValueBehavior.Error })).Code);
        else Assert.Equal(expected, _fixture.ReadCellText(_fixture.Merge(rich).Document, "A1"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Phonetic_hints_on_literals_survive_and_shift_while_replaced_hints_are_removed(bool shared)
    {
        // Offsets count decoded UTF-16 code units. Excel verification of this convention remains Pending.
        var source = Rich(shared, new Run(new Text("漢😀 _x0041_ {Name} 字")));
        source = _fixture.RewritePart(source, "xl/worksheets/sheet1.xml", text =>
        {
            var xml = XDocument.Parse(text); xml.Descendants(S + "row").Single().Add(new XElement(S + "c", new XAttribute("r", "D3"), new XAttribute("t", "str"),
                new XElement(S + "f", "PHONETIC(B3)"), new XElement(S + "v", "old cache")));
            return xml.ToString();
        });
        source = _fixture.RewritePart(source, shared ? "xl/sharedStrings.xml" : "xl/worksheets/sheet1.xml", text =>
        {
            var xml = XDocument.Parse(text); var container = xml.Descendants(S + (shared ? "si" : "is")).First();
            container.Add(new XElement(S + "rPh", new XAttribute("sb", "0"), new XAttribute("eb", "3"), new XElement(S + "t", "unchanged")),
                new XElement(S + "rPh", new XAttribute("sb", "6"), new XAttribute("eb", "12"), new XElement(S + "t", "removed")),
                new XElement(S + "rPh", new XAttribute("sb", "13"), new XAttribute("eb", "14"), new XElement(S + "t", "shifted")),
                new XElement(S + "phoneticPr", new XAttribute("fontId", "0")));
            return xml.ToString();
        });
        var result = _fixture.Merge(source, "{\"Name\":\"Q\"}").Document;
        Assert.Equal("漢😀 A Q 字", _fixture.ReadCellText(result, "A1"));
        using var document = SpreadsheetDocument.Open(new MemoryStream(result), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var sheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        var formula = sheet.Descendants<Cell>().Single(cell => cell.CellReference == "C1");
        Assert.Equal("PHONETIC(A1)", formula.CellFormula!.Text); Assert.Null(formula.CellValue);
        var hints = sheet.Descendants<PhoneticRun>().ToArray();
        Assert.Equal(new[] { "unchanged", "shifted" }, hints.Select(hint => hint.InnerText));
        Assert.Equal(new uint[] { 0, 8 }, hints.Select(hint => hint.BaseTextStartIndex!.Value));
        Assert.Equal(new uint[] { 3, 9 }, hints.Select(hint => hint.EndingBaseIndex!.Value));
    }

    [Fact]
    public void Shared_rich_sources_bind_independently_in_vertical_and_horizontal_header_body_footer_instances()
    {
        var twoAxis = new XlsxTwoAxisFixture();
        var source = twoAxis.Create((workbook, part) =>
        {
            workbook.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable(new SharedStringItem(
                new Run(new RunProperties(new Italic()), new Text("[")), new Run(new RunProperties(new Bold()), new Text("{Na")),
                new Run(new Text("me}")), new Run(new RunProperties(new Underline()), new Text("]"))));
            foreach (var cell in part.Worksheet!.Descendants<Cell>().Where(cell => cell.CellReference!.Value is "E8" or "B9" or "B10" or "E10" or "B11"))
            {
                cell.RemoveAllChildren(); cell.DataType = CellValues.SharedString; cell.CellValue = new CellValue("0");
            }
            part.Worksheet.Descendants<Row>().Single(row => row.RowIndex!.Value == 9).PrependChild(_fixture.Text("A9", "$l(10).Items"));
            part.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "A10").InlineString!.Text!.Text = "$l(20).Items";
            part.Worksheet.Descendants<Row>().Single(row => row.RowIndex!.Value == 11).PrependChild(_fixture.Text("A11", "$l(10).Items"));
        });
        var data = System.Text.Json.JsonSerializer.Serialize(new
        {
            Bids = new[] { new { Company = "Acme" } }, Criteria = new[] { new { Name = "Header A" }, new { Name = "Header B" } },
            Items = new[] { (10, "p", "Parent"), (20, "c", "First") }.SelectMany(item => Enumerable.Range(0, 2).Select(index =>
                new Dictionary<string, object> { ["$level"] = item.Item1, ["$path"] = item.Item2,
                    ["Name"] = item.Item1 == 10 ? item.Item3 : index == 0 ? "First" : "Second", ["Quantity"] = 1 }))
        });
        var result = _fixture.Merge(source, data).Document;
        using var document = SpreadsheetDocument.Open(new MemoryStream(result), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var values = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>()
            .Where(cell => cell.InlineString?.Elements<Run>().Any() == true).ToArray();
        Assert.Contains(values, cell => cell.InlineString!.InnerText == "[Header A]");
        Assert.Contains(values, cell => cell.InlineString!.InnerText == "[Header B]");
        Assert.Contains(values, cell => cell.InlineString!.InnerText == "[First]");
        Assert.Contains(values, cell => cell.InlineString!.InnerText == "[Second]");
        Assert.Equal(2, values.Count(cell => cell.InlineString!.InnerText == "[Parent]"));
        Assert.All(values, cell => { Assert.NotEmpty(cell.Descendants<Bold>()); Assert.NotEmpty(cell.Descendants<Underline>()); });
    }

    private byte[] Rich(bool shared, params Run[] runs) => _fixture.Create((workbook, part) =>
    {
        new XlsxWorksheetRetentionTests().Styles(workbook);
        var cell = part.Worksheet!.Descendants<Cell>().Single(); cell.StyleIndex = 1;
        if (shared)
        {
            workbook.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable(new SharedStringItem(runs));
            cell.RemoveAllChildren(); cell.DataType = CellValues.SharedString; cell.CellValue = new CellValue("0");
        }
        else cell.InlineString = new InlineString(runs);
    });

    private XElement?[] CharacterProperties(byte[] bytes, string address)
    {
        using var document = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        using var stream = document.WorkbookPart!.WorksheetParts.Single().GetStream();
        var cell = XDocument.Load(stream).Descendants(S + "c").Single(cell => (string?)cell.Attribute("r") == address);
        return cell.Descendants(S + "r").SelectMany(run => Enumerable.Repeat(run.Element(S + "rPr"), run.Element(S + "t")!.Value.Length)).ToArray();
    }
}
