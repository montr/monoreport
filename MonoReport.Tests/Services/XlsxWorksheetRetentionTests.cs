using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class XlsxWorksheetRetentionTests
{
    private readonly XlsxFixture _fixture = new();

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Identity_sheets_keep_sparse_content_and_finalize_references_in_mixed_and_uncontrolled_workbooks(bool repeated)
    {
        var compatibility = new XlsxPackageCompatibilityTests();
        var source = compatibility.Metadata(compatibility.Views(Mixed(repeated), "A3:XFD3", "B3", "0", "A1"));
        var original = source.ToArray();
        using var services = _fixture.Services();
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(new MemoryStream(source));
        Assert.True(analysis.Valid, string.Join("; ", analysis.Errors.Select(error => error.Message)));
        var result = _fixture.Merge(source, "{\"Name\":\"Ada\",\"Number\":12.5,\"Items\":[{\"$level\":10,\"Name\":\"One\",\"Price\":2},{\"$level\":10,\"Name\":\"Two\",\"Price\":3}]}");
        Assert.Empty(result.Warnings); Assert.Equal(original, source);
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(output));
        var workbook = output.WorkbookPart!;
        var sheets = workbook.Workbook!.Sheets!.Elements<Sheet>().ToArray();
        Assert.Equal(new[] { "Report", "Empty", "Formats", "Scalar", "References" }, sheets.Select(sheet => sheet.Name!.Value));
        Assert.Equal(new uint[] { 1, 7, 9, 11, 15 }, sheets.Select(sheet => sheet.SheetId!.Value));
        Assert.Equal(SheetStateValues.Hidden, sheets[2].State!.Value); Assert.Equal(SheetStateValues.VeryHidden, sheets[4].State!.Value);
        Worksheet Sheet(string name) => ((WorksheetPart)workbook.GetPartById(sheets.Single(sheet => sheet.Name == name).Id!)).Worksheet!;
        Assert.Empty(Sheet("Empty").Descendants<Row>()); Assert.Empty(Sheet("Empty").Descendants<Cell>());
        Assert.Null(Sheet("Empty").GetFirstChild<SheetDimension>());
        var format = Sheet("Formats");
        Assert.Equal("XFD1048576", Assert.Single(format.Descendants<Cell>()).CellReference!.Value);
        Assert.Equal(1U, Assert.Single(format.Descendants<Cell>()).StyleIndex!.Value);
        Assert.Equal(27, Assert.Single(format.Descendants<Row>()).Height!.Value);
        Assert.Equal(19, Assert.Single(format.Descendants<Column>()).Width!.Value);
        Assert.Equal("XFD1048576", format.GetFirstChild<SheetDimension>()!.Reference!.Value);
        var scalar = Sheet("Scalar").Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal(new[] { "E7", "F7" }, scalar.Keys);
        Assert.Single(Sheet("Scalar").Descendants<Row>());
        Assert.Equal("E7:E8", Assert.Single(Sheet("Scalar").Descendants<MergeCell>()).Reference!.Value);
        Assert.Equal("Person Ada!", scalar["E7"].InlineString!.InnerText);
        Assert.NotEmpty(scalar["E7"].Descendants<Bold>());
        Assert.Equal(CellValues.Number, scalar["F7"].DataType!.Value); Assert.Equal("12.5", scalar["F7"].CellValue!.Text);
        Assert.Equal(1U, scalar["F7"].StyleIndex!.Value); Assert.Null(scalar["F7"].InlineString);
        var reference = Sheet("References").Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        var reportRange = repeated ? "Report!$B$1:$B$2" : "Report!$C$3:$C$3";
        Assert.Equal("SUM(" + reportRange + ")+Scalar!F7+Empty!$XFD$1048576+Formats!Z999", reference["A2"].CellFormula!.Text);
        Assert.Equal("SUM(" + reportRange + ")+Scalar!G7+Empty!$XFD$1048576+Formats!AA999", reference["B2"].CellFormula!.Text);
        Assert.All(reference.Values, cell => { Assert.Null(cell.CellFormula!.FormulaType); Assert.Null(cell.CellFormula.SharedIndex); Assert.Null(cell.CellValue); });
        var report = Sheet("Report").Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal(repeated ? "Scalar!$F$7+B1" : "Scalar!$F$7+C3", report[repeated ? "C1" : "D3"].CellFormula!.Text);
        var names = workbook.Workbook.DefinedNames!.Elements<DefinedName>().ToArray();
        Assert.Equal(repeated ? 1 : 0, names.Count(name => name.Name == "_xlnm.Print_Area"));
        Assert.Equal(repeated ? "Report!$B$1:$B$2" : "Report!$C$3:$C$3", names.Single(name => name.Name == "ReportValues").Text);
        Assert.Equal("Empty!$XFD$1048576", names.Single(name => name.Name == "BlankLimit").Text);
        Assert.True(workbook.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);
        Assert.True(workbook.Workbook.CalculationProperties.ForceFullCalculation!.Value);
        Assert.Equal(CalculateModeValues.Auto, workbook.Workbook.CalculationProperties.CalculationMode!.Value);
        Assert.Null(workbook.CalculationChainPart);
        using var reader = new StreamReader(workbook.GetStream());
        var xml = reader.ReadToEnd(); Assert.DoesNotContain(XlsxPackageCompatibilityTests.PathSentinel, xml); Assert.DoesNotContain("revisionPtr", xml);
    }

    private byte[] Mixed(bool repeated) => _fixture.Create((workbook, part) =>
    {
        Styles(workbook);
        part.Worksheet = new Worksheet(new SheetData(new Row(
            _fixture.Text("B3", "{Name}"), repeated ? _fixture.Text("C3", "{Price}") : new Cell { CellReference = "C3", CellValue = new CellValue("2") },
            new Cell { CellReference = "D3", CellFormula = new CellFormula("Scalar!$F$7+C3"), CellValue = new CellValue("99") }) { RowIndex = 3 }));
        if (repeated) part.Worksheet.Descendants<Row>().Single().PrependChild(_fixture.Text("A3", "$l(10).Items"));
        else workbook.Workbook!.DefinedNames!.RemoveAllChildren();
        void Add(uint id, string name, Worksheet worksheet, SheetStateValues? state = null)
        {
            var sheet = workbook.AddNewPart<WorksheetPart>(); sheet.Worksheet = worksheet;
            workbook.Workbook!.Sheets!.Append(new Sheet { Name = name, SheetId = id, State = state, Id = workbook.GetIdOfPart(sheet) });
        }
        Add(7, "Empty", new Worksheet(new SheetData()));
        Add(9, "Formats", new Worksheet(new SheetDimension { Reference = "XFD1048576" },
            new Columns(new Column { Min = 16384, Max = 16384, Width = 19, CustomWidth = true }),
            new SheetData(new Row(new Cell { CellReference = "XFD1048576", StyleIndex = 1 }) { RowIndex = 1048576, Height = 27, CustomHeight = true })), SheetStateValues.Hidden);
        Add(11, "Scalar", new Worksheet(new SheetData(new Row(
            new Cell { CellReference = "E7", DataType = CellValues.InlineString, InlineString = new InlineString(
                new Run(new RunProperties(new Italic()), new Text("Person ")),
                new Run(new RunProperties(new Bold()), new Text("{Na")), new Run(new Text("me}!"))) },
            new Cell { CellReference = "F7", StyleIndex = 1, DataType = CellValues.InlineString,
                InlineString = new InlineString(new Run(new RunProperties(new Bold()), new Text("{Number}"))) }) { RowIndex = 7 }),
            new MergeCells(new MergeCell { Reference = "E7:E8" })));
        Add(15, "References", new Worksheet(new SheetData(new Row(
            new Cell { CellReference = "A2", CellFormula = new CellFormula("SUM(Report!$C$3:$C$3)+Scalar!F7+Empty!$XFD$1048576+Formats!Z999")
                { FormulaType = CellFormulaValues.Shared, SharedIndex = 3, Reference = "A2:B2" }, CellValue = new CellValue("999") },
            new Cell { CellReference = "B2", CellFormula = new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 3 }, CellValue = new CellValue("999") }) { RowIndex = 2 })), SheetStateValues.VeryHidden);
        workbook.Workbook!.DefinedNames!.Append(new DefinedName("Report!$C$3:$C$3") { Name = "ReportValues" },
            new DefinedName("Empty!$XFD$1048576") { Name = "BlankLimit" });
    });

    [Theory]
    [InlineData("$l(10).Items", false)]
    [InlineData("$l(bad).Items", false)]
    [InlineData("$header(1)", false)]
    [InlineData("$l(10).Items", true)]
    public void Missing_boundary_cannot_hide_valid_malformed_or_misplaced_controls(string control, bool rich)
    {
        var source = _fixture.Create((workbook, part) =>
        {
            workbook.Workbook!.DefinedNames!.RemoveAllChildren();
            var cell = part.Worksheet!.Descendants<Cell>().Single();
            cell.InlineString = rich ? new InlineString(new Run(new Text(control[..3])), new Run(new Text(control[3..]))) : new InlineString(new Text(control));
        });
        using var services = _fixture.Services();
        var issue = Assert.Single(services.GetRequiredService<TemplateAnalyzer>().Analyze(new MemoryStream(source)).Errors);
        Assert.Equal(DiagnosticCodes.XlsxInvalidPrintArea, issue.Code); Assert.Contains("B3", issue.Location);
        Assert.Equal(issue.Code, Assert.Throws<TemplateException>(() => _fixture.Merge(source)).Code);
    }

    [Fact]
    public void Identity_formula_cannot_refer_to_discarded_content_on_a_cropped_sheet()
    {
        var source = _fixture.Create((workbook, part) =>
        {
            var reference = workbook.AddNewPart<WorksheetPart>();
            reference.Worksheet = new Worksheet(new SheetData(new Row(new Cell { CellReference = "A1", CellFormula = new CellFormula("Report!A1") }) { RowIndex = 1 }));
            workbook.Workbook!.Sheets!.Append(new Sheet { Name = "Reference", SheetId = 2, Id = workbook.GetIdOfPart(reference) });
        });
        using var services = _fixture.Services();
        Assert.False(services.GetRequiredService<TemplateAnalyzer>().Analyze(new MemoryStream(source)).Valid);
        MergeResult? result = null;
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, Assert.Throws<TemplateException>(() => result = _fixture.Merge(source)).Code); Assert.Null(result);
    }

    internal void Styles(WorkbookPart workbook) => workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()),
        new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
        new Borders(new Border()), new CellStyleFormats(new CellFormat()), new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 4, ApplyNumberFormat = true }));
}
