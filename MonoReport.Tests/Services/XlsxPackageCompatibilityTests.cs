using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class XlsxPackageCompatibilityTests
{
    private readonly XlsxFixture _fixture = new();
    internal static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    internal const string PathSentinel = "/test-owned/compatibility/";

    internal byte[] Metadata(byte[] source, string prefix = "office15", string variant = "safe") => _fixture.RewritePart(source, "xl/workbook.xml", text =>
    {
        var xml = XDocument.Parse(text);
        XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        XNamespace ac = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac";
        XNamespace revision = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
        XNamespace revision6 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6";
        XNamespace revision10 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";
        var choice = new XElement(mc + "Choice", new XAttribute("Requires", prefix),
            new XElement(ac + "absPath", new XAttribute("url", PathSentinel)));
        var alternate = new XElement(mc + "AlternateContent", new XAttribute(XNamespace.Xmlns + prefix,
            "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"), choice);
        if (variant == "choice") choice.Add(new XElement(S + "definedNames"));
        if (variant == "fallback") alternate.Add(new XElement(mc + "Fallback", new XElement(S + "definedNames")));
        if (variant == "attribute") choice.Elements().Single().SetAttributeValue("meaningful", "keep");
        if (variant == "unknown-extension") xml.Root!.Add(new XElement(S + "extLst", new XElement(S + "ext", new XAttribute("uri", "test"))));
        if (variant == "unknown-revision") xml.Root!.Add(new XElement(revision + "revisionPtr", new XAttribute("revIDLastSave", "0"), new XAttribute("meaningful", "keep")));
        xml.Root!.AddFirst(alternate, new XElement(revision + "revisionPtr", new XAttribute("revIDLastSave", "0"),
            new XAttribute("documentId", "13_ncr:1_{11111111-1111-1111-1111-111111111111}"),
            new XAttribute(revision6 + "coauthVersionLast", "47"), new XAttribute(revision6 + "coauthVersionMax", "47"),
            new XAttribute(revision10 + "uidLastSave", "{00000000-0000-0000-0000-000000000000}")));
        return xml.ToString();
    });

    [Theory]
    [InlineData("office15")]
    [InlineData("renamed")]
    public void Recognized_metadata_is_removed_from_output_without_changing_source(string prefix)
    {
        var source = Metadata(_fixture.Create(), prefix);
        var original = source.ToArray();
        using var services = _fixture.Services();
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(new MemoryStream(source)).Valid);
        var result = _fixture.Merge(source);
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(output));
        using var reader = new StreamReader(output.WorkbookPart!.GetStream());
        var xml = reader.ReadToEnd();
        Assert.DoesNotContain(PathSentinel, xml); Assert.DoesNotContain("revisionPtr", xml); Assert.DoesNotContain("AlternateContent", xml);
        Assert.Equal("Ada", output.WorkbookPart.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single().InlineString!.InnerText);
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData("choice")]
    [InlineData("fallback")]
    [InlineData("attribute")]
    [InlineData("unknown-extension")]
    [InlineData("unknown-revision")]
    [InlineData("relationship")]
    [InlineData("malformed")]
    public void Meaningful_or_malformed_workbook_content_is_rejected_at_the_workbook_part(string variant)
    {
        var source = _fixture.Create((workbook, _) =>
        {
            if (variant == "relationship") workbook.AddExternalRelationship("https://test.invalid/unsupported", new Uri("https://test.invalid/target"));
        });
        source = Metadata(source, variant: variant);
        if (variant == "malformed") source = _fixture.RewritePart(source, "xl/workbook.xml", text => text[..^5]);
        using var services = _fixture.Services();
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(new MemoryStream(source));
        var issue = Assert.Single(analysis.Errors);
        Assert.Equal("/xl/workbook.xml", issue.Location);
        Assert.Equal(variant == "malformed" ? DiagnosticCodes.TemplateInvalidXml : DiagnosticCodes.XlsxUnsupportedFeature, issue.Code);
        var original = source.ToArray(); MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source));
        Assert.Null(result); Assert.Equal(original, source); Assert.Equal(issue.Code, error.Errors[0].Code);
    }

    [Theory]
    [InlineData("extension")]
    [InlineData("attribute")]
    [InlineData("relationship")]
    [InlineData("internal-relationship")]
    public void Unsupported_shared_string_table_content_is_rejected_before_conversion(string variant)
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            var strings = workbook.AddNewPart<SharedStringTablePart>();
            strings.SharedStringTable = new SharedStringTable(new SharedStringItem(
                new Run(new RunProperties(new Bold()), new Text("{Na")), new Run(new Text("me}"))))
                { Count = 1, UniqueCount = 1 };
            if (variant == "relationship") strings.AddExternalRelationship("https://test.invalid/unsupported", new Uri("https://test.invalid/target"));
            if (variant == "internal-relationship")
            {
                using var payload = new MemoryStream(new byte[] { 1 });
                strings.AddExtendedPart("https://test.invalid/unsupported", "application/octet-stream", ".bin").FeedData(payload);
            }
            var cell = worksheet.Worksheet!.Descendants<Cell>().Single();
            cell.RemoveAllChildren(); cell.DataType = CellValues.SharedString; cell.CellValue = new CellValue("0");
        });
        if (variant is "extension" or "attribute") source = _fixture.RewritePart(source, "xl/sharedStrings.xml", text =>
        {
            var xml = XDocument.Parse(text);
            XNamespace unknown = "https://test.invalid/shared-string-metadata";
            if (variant == "extension") xml.Root!.Add(new XElement(S + "extLst", new XElement(S + "ext",
                new XAttribute("uri", "test"), new XElement(unknown + "payload", "must not disappear"))));
            else xml.Root!.SetAttributeValue(unknown + "metadata", "must not disappear");
            return xml.ToString();
        });
        if (variant == "extension")
        {
            using var document = SpreadsheetDocument.Open(new MemoryStream(source), false);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        var original = source.ToArray();
        using var services = _fixture.Services();
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(new MemoryStream(source));
        Assert.False(analysis.Valid);
        var issue = Assert.Single(analysis.Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, issue.Code);
        Assert.Equal("/xl/sharedStrings.xml", issue.Location);
        Assert.Contains("Remove unsupported shared-string table", issue.Message);
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source));
        Assert.Null(result); Assert.Equal(original, source);
        Assert.Equal(issue.Code, error.Errors[0].Code); Assert.Equal(issue.Location, error.Errors[0].Location);
        Assert.Equal(issue.Message, error.Errors[0].Message);
    }

    [Theory]
    [InlineData("B3:C4", "B3", "0", "B3", "A1:B2", "A1", "A1", false)]
    [InlineData("A3:XFD3", "B3", "0", "B3", "A1:C1", "A1", "A1", false)]
    [InlineData("B3 B5:D5", "C5", "1", "B3", "A1 A3:C3", "B3", "A1", false)]
    [InlineData("F9:G10", "F9", "0", "F9", "A1", "A1", "A1", false)]
    [InlineData("B3", "B3", "0", "B3", "A1", "A1", "A1", true)]
    [InlineData("B3:C4", "A1", null, "A1", "A1:B2", "A1", "A1", false)]
    public void Valid_views_normalize_without_changing_report_semantics(string ranges, string active, string? activeId,
        string top, string expectedRanges, string expectedActive, string expectedTop, bool repeated)
    {
        var baseline = repeated ? _fixture.CreateRepeated() : _fixture.Create();
        var source = Views(baseline, ranges, active, activeId, top);
        const string data = "{\"Name\":\"Ada\",\"Items\":[{\"$level\":10,\"Name\":\"A\",\"Price\":2},{\"$level\":10,\"Name\":\"B\",\"Price\":3}]}";
        using var services = _fixture.Services();
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(new MemoryStream(source)).Valid);
        var output = _fixture.Merge(source, data).Document;
        using var document = SpreadsheetDocument.Open(new MemoryStream(output), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var selection = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Selection>().Single();
        Assert.Equal(expectedRanges, selection.SequenceOfReferences!.InnerText); Assert.Equal(expectedActive, selection.ActiveCell!.Value);
        if (activeId is not null) Assert.Equal(expectedRanges.Contains(' ') ? 1U : 0U, selection.ActiveCellId!.Value);
        Assert.Equal(expectedTop, document.WorkbookPart.WorksheetParts.Single().Worksheet!.Descendants<SheetView>().Single().TopLeftCell!.Value);
        Assert.Equal(Content(_fixture.Merge(baseline, data).Document), Content(output));
    }

    [Theory]
    [InlineData("B0", "B3", "0", "B3")]
    [InlineData("B3:", "B3", "0", "B3")]
    [InlineData("B3", "XFE3", "0", "B3")]
    [InlineData("B3", "B3", "0", "A0")]
    [InlineData("B3", "B3", "-1", "B3")]
    [InlineData("B3", "B3", "1", "B3")]
    [InlineData("B3 C4", "B3", "1", "B3")]
    [InlineData("", "B3", "0", "B3")]
    public void Malformed_coordinates_and_inconsistent_active_range_state_fail_before_fallback(string ranges, string active, string activeId, string top)
    {
        var source = Views(_fixture.Create(), ranges, active, activeId, top);
        using var services = _fixture.Services();
        var issue = Assert.Single(services.GetRequiredService<TemplateAnalyzer>().Analyze(new MemoryStream(source)).Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, issue.Code); Assert.Contains("sheet 'Report'", issue.Location);
        Assert.Equal(issue.Code, Assert.Throws<TemplateException>(() => _fixture.Merge(source)).Code);
    }

    internal byte[] Views(byte[] source, string ranges, string active, string? activeId, string top) =>
        _fixture.RewritePart(source, "xl/worksheets/sheet1.xml", text =>
        {
            var xml = XDocument.Parse(text);
            xml.Root!.AddFirst(new XElement(S + "sheetViews", new XElement(S + "sheetView", new XAttribute("workbookViewId", "0"),
                new XAttribute("topLeftCell", top), new XAttribute("showGridLines", "0"),
                new XElement(S + "selection", new XAttribute("sqref", ranges), new XAttribute("activeCell", active),
                    activeId is null ? null : new XAttribute("activeCellId", activeId)))));
            return xml.ToString();
        });

    private string[] Content(byte[] bytes)
    {
        using var document = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var sheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        return new[] { sheet.GetFirstChild<SheetDimension>()!.Reference!.Value! }.Concat(sheet.Descendants<Cell>().Select(cell => cell.OuterXml)).ToArray();
    }
}
