using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class XlsxExcelSavedWorkbookTests
{
    [Fact]
    public void Excel_saved_employees_analyzes_and_generates_nested_rows_with_authored_formatting()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "XlsxCompatibility", "employees-excel-saved.template.xlsx");
        var source = File.ReadAllBytes(path);
        var original = source.ToArray();
        var fixture = new XlsxFixture();
        using var services = fixture.Services();
        using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.True(analysis.Valid, string.Join("; ", analysis.Errors.Select(error => error.Message)));
        var result = fixture.Merge(source, """
            {"Data":[
              {"$level":10,"Department":"Engineering"},
              {"$level":40,"Name":"Ada","Position":"Developer"},
              {"$level":40,"Name":"Беа","Position":"=Designer"},
              {"$level":10,"Department":"Operations"},
              {"$level":40,"Name":"Chen","Position":"Coordinator"}
            ]}
            """);
        Assert.Empty(result.Warnings);
        Assert.Equal(MonoReport.Models.DocumentFormat.Xlsx, result.Format);
        Assert.Equal(original, source);
        Assert.Equal(original, File.ReadAllBytes(path));

        using var template = SpreadsheetDocument.Open(new MemoryStream(source), false);
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        // Excel authored zero DPI values. Pin these two source findings; accept no additional output errors.
        var sourceErrors = validator.Validate(template).ToArray();
        var outputErrors = validator.Validate(output).ToArray();
        foreach (var errors in new[] { sourceErrors, outputErrors })
        {
            Assert.Equal(2, errors.Length);
            foreach (var attribute in new[] { "horizontalDpi", "verticalDpi" })
            {
                var error = Assert.Single(errors, item => item.Description.Contains("'" + attribute + "'", StringComparison.Ordinal));
                Assert.Equal(ValidationErrorType.Schema, error.ErrorType);
                Assert.Equal("Sem_AttributeValueDataTypeDetailed", error.Id);
                Assert.Equal("/xl/worksheets/sheet1.xml", error.Part!.Uri.ToString());
                Assert.IsType<PageSetup>(error.Node);
                Assert.Equal("0", error.Node!.GetAttribute(attribute, string.Empty).Value);
            }
        }
        Assert.Equal(sourceErrors.Select(error => error.Description), outputErrors.Select(error => error.Description));
        var workbook = output.WorkbookPart!;
        var sheet = Assert.Single(workbook.Workbook!.Sheets!.Elements<Sheet>());
        Assert.Equal("Employees", sheet.Name!.Value);
        Assert.Equal(1U, sheet.SheetId!.Value);
        var worksheet = ((WorksheetPart)workbook.GetPartById(sheet.Id!)).Worksheet!;
        var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("A1:D7", worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        var printArea = Assert.Single(workbook.Workbook.DefinedNames!.Elements<DefinedName>());
        Assert.Equal("_xlnm.Print_Area", printArea.Name!.Value);
        Assert.Equal(0U, printArea.LocalSheetId!.Value);
        Assert.Equal("'Employees'!$A$1:$D$7", printArea.Text);
        Assert.Equal(new[] { "B3:D3", "B6:D6" }, worksheet.Descendants<MergeCell>().Select(merge => merge.Reference!.Value));

        void Text(string address, string expected, uint style)
        {
            var cell = cells[address];
            Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
            Assert.Equal(expected, cell.InlineString!.InnerText);
            Assert.Equal(style, cell.StyleIndex!.Value);
            Assert.Null(cell.CellFormula);
        }
        Text("B2", "Department", 8); Text("C2", "Employee", 9); Text("D2", "Position", 10);
        Text("B3", "Engineering", 2); Text("B6", "Operations", 2);
        Text("C4", "Ada", 6); Text("D4", "Developer", 7);
        Text("C5", "Беа", 6); Text("D5", "=Designer", 7);
        Text("C7", "Chen", 6); Text("D7", "Coordinator", 7);
        Assert.Equal(new uint[] { 1, 2, 3, 4, 5, 6, 7 }, worksheet.Descendants<Row>().Select(row => row.RowIndex!.Value));
        Assert.All(worksheet.Descendants<Row>(), row => Assert.Equal(15D, row.Height!.Value));
        Assert.All(cells.Values, cell => Assert.Null(cell.CellFormula));
        Assert.DoesNotContain(cells.Values, cell => cell.InlineString?.InnerText.Contains("$level", StringComparison.Ordinal) == true);
        foreach (var row in new[] { 4, 5, 7 })
        {
            Assert.Equal(5U, cells["B" + row].StyleIndex!.Value);
            Assert.Null(cells["B" + row].CellValue); Assert.Null(cells["B" + row].InlineString);
        }
        foreach (var row in new[] { 3, 6 })
        {
            Assert.Equal(3U, cells["C" + row].StyleIndex!.Value);
            Assert.Equal(4U, cells["D" + row].StyleIndex!.Value);
        }

        var styles = workbook.WorkbookStylesPart!.Stylesheet!;
        var formats = styles.CellFormats!.Elements<CellFormat>().ToArray();
        Assert.Equal((3U, 3U, 1U, 0U), (formats[8].FontId!.Value, formats[8].FillId!.Value, formats[8].BorderId!.Value, formats[8].NumberFormatId!.Value));
        Assert.Equal((2U, 2U, 7U), (formats[2].FontId!.Value, formats[2].FillId!.Value, formats[2].BorderId!.Value));
        Assert.Equal((1U, 0U, 5U), (formats[6].FontId!.Value, formats[6].FillId!.Value, formats[6].BorderId!.Value));
        var fonts = styles.Fonts!.Elements<Font>().ToArray();
        Assert.Equal("Calibri", fonts[3].FontName!.Val!.Value); Assert.Equal(10D, fonts[3].FontSize!.Val!.Value);
        Assert.NotNull(fonts[3].Bold); Assert.Equal(2U, fonts[3].Color!.Theme!.Value);
        Assert.NotNull(fonts[2].Bold); Assert.Null(fonts[1].Bold);
        var fills = styles.Fills!.Elements<Fill>().ToArray();
        Assert.Equal(PatternValues.Solid, fills[3].PatternFill!.PatternType!.Value);
        Assert.Equal(1U, fills[3].PatternFill!.ForegroundColor!.Theme!.Value);
        Assert.Equal(PatternValues.Solid, fills[2].PatternFill!.PatternType!.Value);
        Assert.Equal(0U, fills[2].PatternFill!.ForegroundColor!.Theme!.Value);
        var borders = styles.Borders!.Elements<Border>().ToArray();
        Assert.Equal(BorderStyleValues.Medium, borders[1].LeftBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.Hair, borders[1].RightBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.Medium, borders[5].BottomBorder!.Style!.Value);
        var columns = worksheet.Descendants<Column>().ToArray();
        Assert.Equal(new uint[] { 1, 2, 3, 4 }, columns.Select(column => column.Min!.Value));
        Assert.Equal(new[] { 4.6640625, 14.6640625, 23, 23.1640625 }, columns.Select(column => column.Width!.Value));
        var page = worksheet.GetFirstChild<PageSetup>()!;
        Assert.Equal(OrientationValues.Portrait, page.Orientation!.Value); Assert.Equal(9U, page.PaperSize!.Value);
        Assert.Equal(0.7D, worksheet.GetFirstChild<PageMargins>()!.Left!.Value);
        var view = Assert.Single(worksheet.Descendants<SheetView>());
        Assert.Equal(SheetViewValues.PageBreakPreview, view.View!.Value); Assert.Equal(125U, view.ZoomScale!.Value);
        var selection = Assert.Single(view.Elements<Selection>());
        Assert.Equal("A1", selection.ActiveCell!.Value); Assert.Equal("A1", selection.SequenceOfReferences!.InnerText);

        Assert.NotNull(template.WorkbookPart!.SharedStringTablePart);
        Assert.Contains("absPath", template.WorkbookPart.Workbook!.OuterXml);
        Assert.Contains("revisionPtr", template.WorkbookPart.Workbook.OuterXml);
        Assert.Equal("E9", Assert.Single(template.WorkbookPart.WorksheetParts.Single().Worksheet!.Descendants<Selection>()).ActiveCell!.Value);
        Assert.Null(workbook.SharedStringTablePart);
        Assert.DoesNotContain("absPath", workbook.Workbook.OuterXml);
        Assert.DoesNotContain("/test-owned/compatibility/", workbook.Workbook.OuterXml);
        Assert.DoesNotContain("revisionPtr", workbook.Workbook.OuterXml);
        Assert.Null(workbook.CalculationChainPart);
        var calculation = workbook.Workbook.CalculationProperties!;
        Assert.Equal(CalculateModeValues.Auto, calculation.CalculationMode!.Value);
        Assert.True(calculation.FullCalculationOnLoad!.Value); Assert.True(calculation.ForceFullCalculation!.Value);
    }
}
