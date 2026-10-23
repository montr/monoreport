using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature012CropTests
{
    private readonly XlsxFixture _fixture = new();

    [Fact]
    public void StaticCrop_Physically_removes_external_cells_strings_columns_and_obsolete_calculation_chain()
    {
        const string sentinel = "DISPOSABLE_PRIVATE_SENTINEL";
        var source = _fixture.Create((workbook, worksheet) =>
        {
            var data = worksheet.Worksheet!.GetFirstChild<SheetData>()!;
            var row = data.GetFirstChild<Row>()!;
            row.Height = 22; row.CustomHeight = true; row.Hidden = true;
            data.PrependChild(new Row(new Cell { CellReference = "A1", CellFormula = new CellFormula("1+1"), CellValue = new CellValue("2") }) { RowIndex = 1 });
            data.Append(new Row(new Cell { CellReference = "F9", DataType = CellValues.SharedString, CellValue = new CellValue("0") }) { RowIndex = 9 });
            workbook.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable(new SharedStringItem(new Text(sentinel)));
            workbook.AddNewPart<CalculationChainPart>().CalculationChain = new CalculationChain(new CalculationCell { CellReference = "A1", SheetId = 1 });
            worksheet.Worksheet.InsertAt(new Columns(new Column { Min = 1, Max = 1, Width = 14, CustomWidth = true }), 0);
            worksheet.Worksheet.Append(new PageMargins { Left = 0.1, Right = 0.2, Top = 0.3, Bottom = 0.4, Header = 0.1, Footer = 0.1 },
                new PageSetup { Orientation = OrientationValues.Landscape, FitToWidth = 1 });
        });
        var result = _fixture.Merge(source);
        Assert.Equal(MonoReport.Models.DocumentFormat.Xlsx, result.Format);
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2010).Validate(output));
        var workbook = output.WorkbookPart!;
        var sheet = workbook.WorksheetParts.Single().Worksheet!;
        Assert.Equal("Ada", Assert.Single(sheet.Descendants<Cell>()).InlineString!.InnerText);
        Assert.Equal("A1", Assert.Single(sheet.Descendants<Cell>()).CellReference!.Value);
        Assert.Null(sheet.GetFirstChild<Columns>());
        Assert.Null(workbook.SharedStringTablePart);
        Assert.Null(workbook.CalculationChainPart);
        Assert.Equal("'Report'!$A$1:$C$3", workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text);
        Assert.Equal("A1:C3", sheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal(22, Assert.Single(sheet.Descendants<Row>()).Height!.Value);
        Assert.True(Assert.Single(sheet.Descendants<Row>()).Hidden!.Value);
        Assert.Equal(0.3, sheet.GetFirstChild<PageMargins>()!.Top!.Value);
        Assert.Equal(OrientationValues.Landscape, sheet.GetFirstChild<PageSetup>()!.Orientation!.Value);
        Assert.True(workbook.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);
        using var zip = new ZipArchive(new MemoryStream(result.Document));
        foreach (var entry in zip.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal)))
        {
            using var reader = new StreamReader(entry.Open());
            Assert.DoesNotContain(sentinel, reader.ReadToEnd());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StaticCrop_Preserves_zero_width_merges_rich_nonanchor_style_dates_and_selection(bool date1904)
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            workbook.Workbook!.InsertAt(new WorkbookProperties { Date1904 = date1904 }, 0);
            workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()),
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(new Border()), new CellStyleFormats(new CellFormat()), new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 14 }));
            var data = worksheet.Worksheet!.GetFirstChild<SheetData>()!;
            data.Append(new Row(new Cell { CellReference = "B4", StyleIndex = 1, CellValue = new CellValue("60.5") },
                new Cell { CellReference = "C4", StyleIndex = 1 }) { RowIndex = 4, Height = 0, CustomHeight = true });
            data.Append(new Row(new Cell { CellReference = "B5", DataType = CellValues.SharedString, CellValue = new CellValue("0") }) { RowIndex = 5 });
            workbook.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable(new SharedStringItem(
                new Run(new RunProperties(new Bold()), new Text("Unchanged")), new PhoneticRun(new Text("reading")) { BaseTextStartIndex = 0, EndingBaseIndex = 2 }));
            data.Remove();
            worksheet.Worksheet = new Worksheet(new SheetViews(new SheetView(new Selection { ActiveCell = "B3", SequenceOfReferences = new ListValue<StringValue> { InnerText = "B3:C4" } }) { WorkbookViewId = 0, TopLeftCell = "B3" }),
                new Columns(new Column { Min = 2, Max = 2, Width = 0, CustomWidth = true, Hidden = true },
                    new Column { Min = 3, Max = 4, Width = 12, CustomWidth = true }), data,
                new MergeCells(new MergeCell { Reference = "B4:C4" }),
                new RowBreaks(new Break { Id = 4, Min = 1, Max = 3, ManualPageBreak = true }) { Count = 1, ManualBreakCount = 1 });
        });
        var result = _fixture.Merge(source);
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2010).Validate(output));
        var workbook = output.WorkbookPart!;
        Assert.Equal(date1904, workbook.Workbook!.WorkbookProperties!.Date1904!.Value);
        var sheet = workbook.WorksheetParts.Single().Worksheet!;
        var columns = sheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
        Assert.Equal(1U, columns[0].Min!.Value); Assert.Equal(0, columns[0].Width!.Value); Assert.True(columns[0].Hidden!.Value);
        Assert.Null(columns[1].Hidden);
        Assert.Equal("A2:B2", sheet.Descendants<MergeCell>().Single().Reference!.Value);
        Assert.Equal(1U, sheet.Descendants<Cell>().Single(cell => cell.CellReference == "B2").StyleIndex!.Value);
        var date = sheet.Descendants<Cell>().Single(cell => cell.CellReference == "A2");
        Assert.Equal("60.5", date.CellValue!.Text); Assert.Equal(1U, date.StyleIndex!.Value);
        Assert.Equal(0, sheet.Descendants<Row>().Single(row => row.RowIndex!.Value == 2).Height!.Value);
        Assert.Null(sheet.Descendants<Row>().Single(row => row.RowIndex!.Value == 2).Hidden);
        var rich = sheet.Descendants<Cell>().Single(cell => cell.CellReference == "A3").InlineString!;
        Assert.NotNull(rich.Descendants<Bold>().Single());
        Assert.Equal("Unchanged", rich.Elements<Run>().Single().InnerText);
        Assert.Equal("reading", rich.Elements<PhoneticRun>().Single().InnerText);
        var selection = sheet.Descendants<Selection>().Single();
        Assert.Equal("A1", selection.ActiveCell!.Value);
        Assert.Equal("A1:B2", selection.SequenceOfReferences!.InnerText);
        Assert.Equal("A1", sheet.Descendants<SheetView>().Single().TopLeftCell!.Value);
        var pageBreak = sheet.Descendants<Break>().Single();
        Assert.Equal(2U, pageBreak.Id!.Value); Assert.Equal(0U, pageBreak.Min!.Value); Assert.Equal(2U, pageBreak.Max!.Value);
    }

    [Theory]
    [InlineData("Report!$B$3:$D$5")]
    [InlineData("'Report'!$B$3:$D$5")]
    public void StaticCrop_Uses_resolved_workbook_scoped_printarea_identity(string area)
    {
        var source = _fixture.Create((workbook, _) =>
        {
            var name = workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single();
            name.LocalSheetId = null; name.Text = area;
        });
        var result = _fixture.Merge(source);
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        var name = output.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().Single();
        Assert.Null(name.LocalSheetId); Assert.Equal("'Report'!$A$1:$C$3", name.Text);
    }

    [Theory]
    [InlineData("hyperlink", "hyperlinks")]
    [InlineData("print-title", "print titles")]
    [InlineData("pane", "pane")]
    [InlineData("partial-merge", "partial merge")]
    [InlineData("invalid-column", "coordinates")]
    [InlineData("relationship-only", "relationship")]
    [InlineData("nested-extension", "nested extension")]
    public void StaticCrop_Analyze_and_merge_reject_unproven_structures_with_published_remedies(string scenario, string reason)
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            var sheet = worksheet.Worksheet!;
            if (scenario == "hyperlink") sheet.Append(new Hyperlinks(new Hyperlink { Reference = "B3", Location = "Report!B4" }));
            if (scenario == "print-title") workbook.Workbook!.DefinedNames!.Append(new DefinedName("Report!$B$3:$D$3") { Name = "_xlnm.Print_Titles", LocalSheetId = 0 });
            if (scenario == "pane") sheet.PrependChild(new SheetViews(new SheetView(new Pane { State = PaneStateValues.Frozen, VerticalSplit = 2, TopLeftCell = "B3" }) { WorkbookViewId = 0 }));
            if (scenario == "partial-merge") sheet.Append(new MergeCells(new MergeCell { Reference = "A3:B3" }));
            if (scenario == "invalid-column") sheet.PrependChild(new Columns(new Column { Min = 3, Max = 2, Width = 10 }));
            if (scenario == "relationship-only") worksheet.AddNewPart<PivotTablePart>();
            if (scenario == "nested-extension") sheet.PrependChild(new SheetViews(new SheetView(new ExtensionList(new Extension { Uri = "test" })) { WorkbookViewId = 0 }));
        });
        using var services = _fixture.Services();
        using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(analysis.Valid);
        var issue = Assert.Single(analysis.Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, issue.Code);
        Assert.Contains(reason, issue.Message);
        MergeResult? result = null;
        using var data = JsonDocument.Parse("{\"Name\":\"Ada\"}");
        input.Position = 0;
        var error = Assert.Throws<TemplateException>(() => result = services.GetRequiredService<MergeEngine>().Merge(input, data));
        Assert.Null(result); Assert.Equal(source, input.ToArray());
        Assert.Equal(issue.Code, error.Errors[0].Code);
        Assert.Equal(issue.Message, error.Errors[0].Message);
        Assert.NotNull(error.Errors[0].Location);
    }
}
