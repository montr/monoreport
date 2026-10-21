using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature012M1RevisionTests
{
    private readonly XlsxFixture _fixture = new();

    [Theory]
    [InlineData("_x0041_")]
    [InlineData("_x000D_")]
    [InlineData("_x005F_")]
    [InlineData("_x005F_x0041_")]
    [InlineData("_x0041_x0042_")]
    [InlineData("_x00aB_")]
    [InlineData("_x00000041_")]
    [InlineData("_X0041_")]
    [InlineData("a\rb")]
    [InlineData("a\r\nb")]
    [InlineData("a\nb\tc")]
    [InlineData("😀 {Other} =1")]
    public void F1_Generated_strings_round_trip_under_spreadsheet_semantics(string value)
    {
        var source = _fixture.Create((_, worksheet) => worksheet.Worksheet!.Descendants<Row>().Single()
            .Append(_fixture.Text("C3", "Value={Name}"), _fixture.Text("D3", "{Name lower}")));
        var result = _fixture.Merge(source, JsonSerializer.Serialize(new { Name = value, Other = "never reparse" }));
        Assert.Equal(value, _fixture.ReadCellText(result.Document, "A1"));
        Assert.Equal("Value=" + value, _fixture.ReadCellText(result.Document, "B1"));
        Assert.Equal(value.ToLowerInvariant(), _fixture.ReadCellText(result.Document, "C1"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void F1_Authored_string_escapes_are_interpreted_once_and_unchanged_text_is_preserved(bool shared)
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            var row = worksheet.Worksheet!.Descendants<Row>().Single();
            row.RemoveAllChildren();
            row.Append(_fixture.Text("B3", "_x0041_"), _fixture.Text("C3", "_x005F_x0041_"), _fixture.Text("D3", "_x000D_"));
            var field = _fixture.Text("B4", "_x005F_x0041_/_x000D_/_x0001_/{Name}");
            if (shared)
            {
                workbook.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable(new SharedStringItem(new Text(field.InlineString!.InnerText)));
                field.RemoveAllChildren(); field.DataType = CellValues.SharedString; field.CellValue = new CellValue("0");
            }
            var data = worksheet.Worksheet.GetFirstChild<SheetData>()!;
            data.Append(new Row(field) { RowIndex = 4 }, new Row(_fixture.Text("B5", "AUTHORED_CR")) { RowIndex = 5 });
        });
        source = _fixture.RewritePart(source, "xl/worksheets/sheet1.xml", xml => xml.Replace("AUTHORED_CR", "a&#xD;b", StringComparison.Ordinal));
        var result = _fixture.Merge(source, "{\"Name\":\"_x000D_\"}");
        Assert.Equal("A", _fixture.ReadCellText(result.Document, "A1"));
        Assert.Equal("_x0041_", _fixture.ReadCellText(result.Document, "B1"));
        Assert.Equal("\r", _fixture.ReadCellText(result.Document, "C1"));
        Assert.Equal("_x0041_/\r/\u0001/_x000D_", _fixture.ReadCellText(result.Document, "A2"));
        Assert.Equal("a\rb", _fixture.ReadCellText(result.Document, "A3"));
    }

    [Theory]
    [InlineData("\\uD800", "{Name}")]
    [InlineData("\\uDC00", "{Name}")]
    [InlineData("\\uD800", "Value={Name}")]
    [InlineData("\\uDC00", "{Name upper}")]
    public void F2_Malformed_json_surrogates_have_safe_coded_field_diagnostics(string surrogate, string template)
    {
        var source = _fixture.Create((_, worksheet) => worksheet.Worksheet!.Descendants<Text>().Single().Text = template);
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        using var data = JsonDocument.Parse("{\"Name\":\"private" + surrogate + "\"}");
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = services.GetRequiredService<MergeEngine>().Merge(input, data));
        Assert.Null(result); Assert.Equal(source, input.ToArray());
        var issue = Assert.Single(error.Errors);
        Assert.Equal(DiagnosticCodes.GenerationInvalidData, issue.Code); Assert.Equal("$.Name", issue.DataPath);
        Assert.Contains("B3", issue.Location); Assert.Contains("Unicode", issue.Message); Assert.Contains("surrogate", issue.Message);
        Assert.DoesNotContain("private", JsonSerializer.Serialize(error.Errors));
    }

    [Theory]
    [InlineData("'Report'!$B$3:$D$5", "B3")]
    [InlineData("'Report'!$A$1:$D$5", "F9")]
    public void F3_Affected_sync_reference_rejects_in_shared_preflight(string printArea, string anchor)
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = printArea;
            worksheet.Worksheet!.PrependChild(new SheetProperties { SyncHorizontal = true, SyncReference = anchor });
        });
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(analysis.Valid);
        var issue = Assert.Single(analysis.Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, issue.Code); Assert.Contains("syncRef", issue.Message); Assert.Contains("Remove", issue.Message);
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source));
        Assert.Equal(issue.Code, error.Code); Assert.Equal(issue.Message, error.Errors[0].Message);
    }

    [Fact]
    public void F3_Unaffected_sync_reference_remains_unchanged()
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$A$1:$D$5";
            worksheet.Worksheet!.PrependChild(new SheetProperties { SyncHorizontal = true, SyncReference = "B3" });
        });
        var result = _fixture.Merge(source);
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Equal("B3", output.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetProperties>()!.SyncReference!.Value);
    }

    [Theory]
    [InlineData(97, false)] [InlineData(98, false)] [InlineData(99, false)] [InlineData(100, false)]
    [InlineData(97, true)] [InlineData(98, true)] [InlineData(99, true)] [InlineData(100, true)]
    public void F4_Error_cap_keeps_earlier_conflicts_before_later_parse_errors(int syntaxErrors, bool sameCell)
    {
        var source = _fixture.Create((_, worksheet) =>
        {
            var row = worksheet.Worksheet!.Descendants<Row>().Single(); row.RemoveAllChildren();
            var syntax = string.Join(' ', Enumerable.Repeat("{!}", syntaxErrors));
            if (sameCell) row.Append(_fixture.Text("B3", "{Thing} {Thing.Name} " + syntax));
            else row.Append(_fixture.Text("B3", "{Thing}"), _fixture.Text("C3", "{Thing.Name}"), _fixture.Text("D3", syntax));
        });
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(analysis.Valid); Assert.Equal(Math.Min(100, syntaxErrors + 2), analysis.Errors.Count);
        Assert.All(analysis.Errors.Take(2), issue => Assert.Equal(DiagnosticCodes.TemplateRequirementConflict, issue.Code));
        Assert.Equal(new[] { "{Thing}", "{Thing.Name}" }, analysis.Errors.Take(2).Select(issue => issue.Field));
        Assert.All(analysis.Errors.Skip(2), issue => Assert.Equal(DiagnosticCodes.TemplateInvalidExpression, issue.Code));
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source));
        Assert.Equal(DiagnosticCodes.TemplateRequirementConflict, error.Code);
        Assert.Equal(syntaxErrors >= 98, error.LimitReached);
        Assert.Equal(analysis.Errors.Select(issue => (issue.Code, issue.Field, issue.Location)), error.Errors.Select(issue => (issue.Code, issue.Field, issue.Location)));
        Assert.Contains("B3", error.Location);
    }

    [Theory]
    [InlineData(SpreadsheetDocumentType.Template)]
    [InlineData(SpreadsheetDocumentType.MacroEnabledWorkbook)]
    [InlineData(SpreadsheetDocumentType.MacroEnabledTemplate)]
    public void F5_Unsupported_excel_kinds_explain_the_ordinary_xlsx_remedy(SpreadsheetDocumentType kind)
    {
        var source = _fixture.Create(kind: kind);
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        var issue = Assert.Single(analysis.Errors);
        Assert.Equal(DiagnosticCodes.TemplateInvalidPackage, issue.Code);
        Assert.Contains("Word", issue.Message); Assert.Contains("ordinary", issue.Message); Assert.Contains(".xlsx", issue.Message);
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source));
        Assert.Equal(issue.Code, error.Code); Assert.Equal(issue.Message, error.Errors[0].Message);
    }

    [Theory]
    [InlineData(97)] [InlineData(98)] [InlineData(99)] [InlineData(100)]
    public void F4_Later_sheet_conflict_still_reports_the_earlier_use_before_applying_the_cap(int syntaxErrors)
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            var row = worksheet.Worksheet!.Descendants<Row>().Single(); row.RemoveAllChildren();
            row.Append(_fixture.Text("B3", "{Thing}"), _fixture.Text("C3", string.Join(' ', Enumerable.Repeat("{!}", syntaxErrors))));
            var other = workbook.AddNewPart<WorksheetPart>();
            other.Worksheet = new Worksheet(new SheetData(new Row(_fixture.Text("A1", "{Thing.Name}"), _fixture.Text("B1", "{!}")) { RowIndex = 1 }));
            workbook.Workbook!.Sheets!.Append(new Sheet { Name = "A later", SheetId = 2, Id = workbook.GetIdOfPart(other) });
            workbook.Workbook.DefinedNames!.Append(new DefinedName("'A later'!$A$1:$B$1") { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
        });
        var expected = new List<string> { DiagnosticCodes.TemplateRequirementConflict };
        expected.AddRange(Enumerable.Repeat(DiagnosticCodes.TemplateInvalidExpression, syntaxErrors));
        expected.Add(DiagnosticCodes.TemplateRequirementConflict); expected.Add(DiagnosticCodes.TemplateInvalidExpression);
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.Equal(expected.Take(100), analysis.Errors.Select(issue => issue.Code));
        Assert.Contains("Report", analysis.Errors[0].Location); Assert.Contains("B3", analysis.Errors[0].Location);
        if (syntaxErrors <= 98) Assert.Contains("A later", analysis.Errors[syntaxErrors + 1].Location);
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source));
        Assert.Equal(DiagnosticCodes.TemplateRequirementConflict, error.Code); Assert.True(error.LimitReached);
        Assert.Equal(analysis.Errors.Select(issue => (issue.Code, issue.Location)), error.Errors.Select(issue => (issue.Code, issue.Location)));
    }
}
