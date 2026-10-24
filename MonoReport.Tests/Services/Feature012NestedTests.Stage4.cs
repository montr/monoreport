using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using MonoReport.Models;
using System.Xml.Linq;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012NestedTests
{
    [Theory]
    [InlineData(2, 2, "A1:AH6", "'Report'!$A$1:$AH$6")]
    [InlineData(0, 2, "A1:J6", "'Report'!$A$1:$J$6")]
    public void Stage4_Deep_render_preserves_styles_dimensions_workbook_and_print_settings_after_crop(
        int branches, int leaves, string dimension, string printArea)
    {
        var source = Stage4FidelityTemplate();
        var rendered = RenderNested(source, Stage3BranchingData(branches, leaves));

        using var package = SpreadsheetDocument.Open(new MemoryStream(rendered.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var workbook = package.WorkbookPart!.Workbook!;
        var sheet = package.WorkbookPart.WorksheetParts.Single().Worksheet!;
        var formats = package.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ToArray();
        Assert.True(workbook.WorkbookProperties!.Date1904!.Value);
        Assert.Equal(dimension, sheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal(printArea, workbook.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "_xlnm.Print_Area").Text);
        Assert.Equal(1U, formats[1].FontId!.Value);
        Assert.Equal(1U, formats[1].FillId!.Value);
        Assert.Equal(1U, formats[1].BorderId!.Value);
        Assert.Equal(HorizontalAlignmentValues.Center, formats[1].Alignment!.Horizontal!.Value);
        Assert.Equal(2U, formats[2].NumberFormatId!.Value);

        var outputCells = sheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!, StringComparer.Ordinal);
        if (branches > 0 && leaves > 0)
        {
            Assert.All(new[] { "D4", "G4", "T4", "W4" }, address => Assert.Contains(address, outputCells.Keys));
            Assert.All(new[] { "D4", "G4", "T4", "W4" }, address => Assert.Equal(1U, outputCells[address].StyleIndex!.Value));
            Assert.All(new[] { "E5", "F5", "L5", "M5", "U5", "V5", "AB5", "AC5" },
                address => Assert.Equal(2U, outputCells[address].StyleIndex!.Value));
            Assert.Equal(new[] { "C3:R3", "S3:AH3" }, sheet.Descendants<MergeCell>()
                .Select(merge => merge.Reference!.Value!).OrderBy(reference => reference, StringComparer.Ordinal));
        }
        else
        {
            Assert.DoesNotContain(outputCells.Keys, address => new[] { "D4", "G4", "T4", "W4", "E5", "F5" }.Contains(address));
        }

        Assert.Equal(new[] { (3U, 18d, false), (4U, 27d, true), (5U, 31d, false) },
            sheet.Descendants<Row>().Where(row => row.RowIndex!.Value is >= 3 and <= 5)
                .Select(row => (row.RowIndex!.Value, row.Height!.Value, row.Hidden!.Value)));
        Assert.Equal(new[] { 3.5d, 0.75d, 0.75d, 0.75d },
            new[] { 0, 1, 2, 3 }.Select(index => sheet.GetFirstChild<Columns>()!.Elements<Column>().ElementAt(index).Width!.Value));
        Assert.True(sheet.GetFirstChild<Columns>()!.Elements<Column>().First().Hidden!.Value);
        Assert.Equal("landscape", sheet.GetFirstChild<PageSetup>()!.GetAttribute("orientation", "").Value);
        Assert.Equal(1U, sheet.GetFirstChild<PageSetup>()!.FitToWidth!.Value);
        Assert.Equal(0.3d, sheet.GetFirstChild<PageMargins>()!.Left!.Value);
        Assert.Equal("&A | &P", sheet.GetFirstChild<HeaderFooter>()!.GetFirstChild<OddHeader>()!.Text);
        Assert.True(sheet.GetFirstChild<PrintOptions>()!.HorizontalCentered!.Value);
        Assert.Empty(rendered.Context.Errors);
    }

    [Theory]
    [InlineData("conditionalFormatting", "conditionalFormatting")]
    [InlineData("dataValidations", "dataValidations")]
    public void Stage4_Affected_worksheet_features_are_rejected_with_a_remedy_and_no_workbook(string feature, string diagnostic)
    {
        var source = _fixture.RewritePart(Stage3BranchingTemplate(), "xl/worksheets/sheet1.xml", xml =>
        {
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var worksheet = XDocument.Parse(xml);
            var element = feature == "conditionalFormatting"
                ? new XElement(ns + feature, new XAttribute("sqref", "E8:E9"),
                    new XElement(ns + "cfRule", new XAttribute("type", "expression"), new XAttribute("priority", "1"),
                        new XElement(ns + "formula", "A1>0")))
                : new XElement(ns + feature, new XElement(ns + "dataValidation", new XAttribute("type", "list"),
                    new XAttribute("sqref", "E8:E9"), new XElement(ns + "formula1", "\"A,B\"")));
            worksheet.Root!.Add(element);
            return worksheet.ToString(SaveOptions.DisableFormatting);
        });
        using (var inputPackage = SpreadsheetDocument.Open(new MemoryStream(source), false))
            Assert.Single(inputPackage.WorkbookPart!.WorksheetParts.Single().Worksheet!.ChildElements,
                child => child.LocalName == feature);

        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, Stage3BranchingData(2, 2)));

        Assert.Null(result);
        Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, error.Code);
        Assert.Contains(error.Errors, issue => issue.Location!.Contains("Report", StringComparison.Ordinal) &&
            issue.Message.Contains(diagnostic, StringComparison.Ordinal) &&
            issue.Message.Contains("remove", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("drawing")]
    [InlineData("table")]
    public void Stage4_Unhandled_drawing_and_table_relationships_are_rejected_before_nested_render(string feature)
    {
        var source = Stage3BranchingTemplate();
        using var writable = new MemoryStream();
        writable.Write(source);
        writable.Position = 0;
        using (var package = SpreadsheetDocument.Open(writable, true))
        {
            var worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
            if (feature == "drawing")
            {
                var drawing = worksheetPart.AddNewPart<DrawingsPart>();
                drawing.WorksheetDrawing = new WorksheetDrawing();
            }
            else
            {
                var table = worksheetPart.AddNewPart<TableDefinitionPart>();
                table.Table = new Table { Id = 1, Name = "ReportTable", DisplayName = "ReportTable", Reference = "E8:E9" };
            }
        }
        source = writable.ToArray();

        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, Stage3BranchingData(2, 2)));

        Assert.Null(result);
        Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, error.Code);
        Assert.Contains(error.Errors, issue => issue.Location!.Contains("Report", StringComparison.Ordinal) &&
            issue.Message.Contains("unsupported worksheet relationship", StringComparison.OrdinalIgnoreCase));
    }

    private byte[] Stage4FidelityTemplate()
    {
        var bytes = Stage3BranchingTemplate();
        using var buffer = new MemoryStream();
        buffer.Write(bytes);
        buffer.Position = 0;
        using (var package = SpreadsheetDocument.Open(buffer, true))
        {
            var workbook = package.WorkbookPart!;
            workbook.Workbook!.InsertAt(new WorkbookProperties { Date1904 = true }, 0);
            var stylePart = workbook.AddNewPart<WorkbookStylesPart>();
            stylePart.Stylesheet = new Stylesheet(
                new Fonts(new Font(), new Font(new Bold())),
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor { Rgb = "FFFFFF00" },
                        BackgroundColor = new BackgroundColor { Indexed = 64 } })),
                new Borders(new Border(), new Border(new LeftBorder { Style = BorderStyleValues.Thin })),
                new CellStyleFormats(new CellFormat()),
                new CellFormats(new CellFormat(),
                    new CellFormat { FontId = 1, FillId = 1, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true,
                        ApplyAlignment = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } },
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true, NumberFormatId = 2, ApplyNumberFormat = true }));

            var sheet = workbook.WorksheetParts.Single().Worksheet!;
            foreach (var cell in sheet.Descendants<Cell>())
            {
                if (cell.CellReference!.Value is "E8" or "H8") cell.StyleIndex = 1;
                if (cell.CellReference.Value is "F9" or "I9") cell.StyleIndex = 2;
            }
            sheet.InsertBefore(new Columns(
                new Column { Min = 1, Max = 1, Width = 2.25, CustomWidth = true },
                new Column { Min = 2, Max = 2, Width = 3.5, CustomWidth = true, Hidden = true },
                new Column { Min = 3, Max = 4, Width = 0.75, CustomWidth = true }), sheet.GetFirstChild<SheetData>());
            foreach (var (rowIndex, height) in new[] { (7U, 18d), (8U, 27d), (9U, 31d) })
            {
                var row = sheet.Descendants<Row>().Single(item => item.RowIndex!.Value == rowIndex);
                row.Height = height; row.CustomHeight = true; row.Hidden = rowIndex == 8;
            }
            sheet.Append(new PrintOptions { HorizontalCentered = true },
                new PageMargins { Left = 0.3, Right = 0.4, Top = 0.5, Bottom = 0.6, Header = 0.2, Footer = 0.2 },
                new PageSetup { FitToWidth = 1, FitToHeight = 0 },
                new HeaderFooter(new OddHeader("&A | &P")));
            sheet.GetFirstChild<PageSetup>()!.SetAttribute(new OpenXmlAttribute("orientation", "", "landscape"));
        }
        return buffer.ToArray();
    }
}
