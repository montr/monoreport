using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MonoReport.Tests.Services;

internal sealed class XlsxNestedFixture
{
    private readonly XlsxFixture _fixture = new();

    internal byte[] CreateActivated() => _fixture.Create((workbook, part) =>
    {
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$7:$G$16";
        part.Worksheet = new Worksheet(new SheetData(
            new Row(_fixture.Text("D1", "$h(1)"), _fixture.Text("E1", "$h(1)"), _fixture.Text("F1", "$h(1)")) { RowIndex = 1 },
            new Row(_fixture.Text("D2", "$h(2)"), _fixture.Text("F2", "$h(2)")) { RowIndex = 2 },
            new Row(_fixture.Text("D3", "$h(3)"), _fixture.Text("F3", "$h(3)")) { RowIndex = 3 },
            new Row(_fixture.Text("A7", "$h(1).Bids"), _fixture.Text("D7", "{Name}")) { RowIndex = 7 },
            new Row(_fixture.Text("A8", "$h(1).Bids"), _fixture.Text("D8", "{Name}")) { RowIndex = 8 },
            new Row(_fixture.Text("A9", "$h(2).Criteria"), _fixture.Text("D9", "{Name}"), _fixture.Text("F9", "{Name}")) { RowIndex = 9 },
            new Row(_fixture.Text("A10", "$h(2).Criteria"), _fixture.Text("D10", "{Name}"), _fixture.Text("F10", "{Name}")) { RowIndex = 10 },
            new Row(_fixture.Text("A11", "$h(3).Third"), _fixture.Text("D11", "{Name}"), _fixture.Text("F11", "{Name}")) { RowIndex = 11 },
            new Row(_fixture.Text("A12", "$h(3).Third"), _fixture.Text("D12", "{Name}"), _fixture.Text("F12", "{Name}")) { RowIndex = 12 },
            new Row(_fixture.Text("A14", "$l(10).Items"), _fixture.Text("B14", "{Name}"), _fixture.Text("C14", "{Value}")) { RowIndex = 14 },
            new Row(_fixture.Text("A15", "$l(20).Items"), _fixture.Text("B15", "{Name}/{../Name}"), _fixture.Text("C15", "{Value}")) { RowIndex = 15 },
            new Row(_fixture.Text("A16", "$l(10).Items"), _fixture.Text("B16", "Total"), new Cell { CellReference = "C16", CellFormula = new CellFormula("SUM(C14:C15)") }) { RowIndex = 16 }),
            new MergeCells(new MergeCell { Reference = "B7:C7" }));
    });

    internal static string ActivatedData() => JsonSerializer.Serialize(new
    {
        Bids = new[] { new { Name = "bid" } }, Criteria = new[] { new { Name = "criterion" } },
        Third = new[] { new { Name = "third" } },
        Items = new[]
        {
            new Dictionary<string, object> { ["$level"] = 10, ["$path"] = "parent", ["Name"] = "parent", ["Value"] = 4 },
            new Dictionary<string, object> { ["$level"] = 20, ["$path"] = "child", ["Name"] = "child", ["Value"] = 5 }
        }
    });

    internal static void AssertActivated(byte[] bytes)
    {
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package));
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("A1:F10", worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal("bid", cells["C1"].InlineString!.InnerText);
        Assert.Equal("criterion", cells["C4"].InlineString!.InnerText);
        Assert.Equal("third", cells["C6"].InlineString!.InnerText);
        Assert.Equal("child/parent", cells["A9"].InlineString!.InnerText);
        Assert.Equal("SUM(B8:B9)", cells["B10"].CellFormula!.Text);
        Assert.Equal("A1:B1", Assert.Single(worksheet.Descendants<MergeCell>()).Reference!.Value);
    }

    internal byte[] Create(bool shared, Action<WorkbookPart, WorksheetPart>? customize = null) => _fixture.Create((workbook, part) =>
    {
        workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()),
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })), new Borders(new Border()),
            new CellStyleFormats(new CellFormat()), new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 1 },
                new CellFormat { NumberFormatId = 2 }, new CellFormat { NumberFormatId = 3 }));
        Cell Formula(string address, string text, bool master = false)
        {
            var formula = new CellFormula(text);
            if (shared && address is "E9" or "E12")
            {
                formula = new CellFormula(master ? text : "ignored follower text") { FormulaType = CellFormulaValues.Shared, SharedIndex = 7 };
                if (master) formula.Reference = "E9:E12";
            }
            return new Cell { CellReference = address, CellValue = new CellValue("99999"), CellFormula = formula };
        }
        var name = _fixture.Text("B10", "{Name}"); name.StyleIndex = 1;
        var price = _fixture.Text("E10", "{Price}"); price.StyleIndex = 3;
        part.Worksheet = new Worksheet(new Columns(new Column { Min = 2, Max = 3, Width = 18, CustomWidth = true },
                new Column { Min = 4, Max = 4, Width = 0, CustomWidth = true, Hidden = true }, new Column { Min = 5, Max = 6, Width = 14, CustomWidth = true }),
            new SheetData(
                new Row(_fixture.Text("D1", "$h(1)"), _fixture.Text("E1", "$h(1)")) { RowIndex = 1 },
                new Row(_fixture.Text("E2", "$h(2)")) { RowIndex = 2 },
                new Row(_fixture.Text("B6", "DISPOSABLE ABOVE")) { RowIndex = 6 },
                new Row(_fixture.Text("A7", "$header(1).Bids"), _fixture.Text("D7", "{Company}"), new Cell { CellReference = "E7" }, _fixture.Text("F7", "Static right")) { RowIndex = 7 },
                new Row(_fixture.Text("A8", "$header(2).Criteria"), _fixture.Text("E8", "{Name}")) { RowIndex = 8 },
                new Row(_fixture.Text("A9", "$l(10).Items"), _fixture.Text("B9", "{Name}"), Formula("E9", "SUBTOTAL(9,E$10:E$11)")) { RowIndex = 9 },
                new Row(_fixture.Text("A10", "$l(20).Items"), name, _fixture.Text("C10", "{Quantity}"), _fixture.Text("D10", "{Name}"), price) { RowIndex = 10, Height = 18, CustomHeight = true },
                new Row(_fixture.Text("A11", "$level(20).Items"), new Cell { CellReference = "B11", StyleIndex = 2 }, _fixture.Text("C11", "{../Name}"), _fixture.Text("E11", "{Delta}")) { RowIndex = 11, Height = 19, CustomHeight = true },
                new Row(_fixture.Text("A12", "$l(10).Items"), _fixture.Text("B12", "Footer: {Name}"), Formula("E12", "SUBTOTAL(9,E$10:E$11)", master: true)) { RowIndex = 12 },
                new Row(_fixture.Text("B13", "Grand"), Formula("E13", "SUBTOTAL(9,E9:E12)")) { RowIndex = 13 },
                new Row(_fixture.Text("B14", "Separator")) { RowIndex = 14, Height = 0, CustomHeight = true, Hidden = true },
                new Row(_fixture.Text("A15", "$l(5).Other"), _fixture.Text("B15", "{Name}")) { RowIndex = 15 },
                new Row(_fixture.Text("B16", "After second table")) { RowIndex = 16 },
                new Row(_fixture.Text("B17", "DISPOSABLE BELOW"), _fixture.Text("G17", "DISPOSABLE RIGHT")) { RowIndex = 17 }),
            new MergeCells(new MergeCell { Reference = "D7:E7" }, new MergeCell { Reference = "B10:B11" }));
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$7:$F$16";
        var summary = workbook.AddNewPart<WorksheetPart>();
        summary.Worksheet = new Worksheet(new SheetData(new Row(Formula("B4", "SUM(Report!D13:E13)")) { RowIndex = 4 }));
        workbook.Workbook.Sheets!.Append(new Sheet { Name = "Summary", SheetId = 2, Id = workbook.GetIdOfPart(summary) });
        workbook.Workbook.DefinedNames.Append(new DefinedName("'Summary'!$B$4:$B$4") { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
        customize?.Invoke(workbook, part);
    });

    internal string Data(int criteria) => JsonSerializer.Serialize(new
    {
        Bids = new[] { new { Company = "Alpha" }, new { Company = "Beta" } },
        Criteria = Enumerable.Range(0, criteria).Select(index => new { Name = "Criterion " + index }),
        Items = new[] { (10, "z", "Lot 1", 100), (20, "opaque-child", "A", 10), (20, "a", "B", 20), (10, "b", "Lot 2", 200), (20, "no/slash/ancestry", "C", 30) }
            .SelectMany((item, index) => Enumerable.Range(0, criteria * 2).Select(tuple => new Dictionary<string, object>
            {
                ["$level"] = item.Item1, ["$path"] = item.Item2, ["Name"] = tuple == 0 ? item.Item3 : "not ordinary name " + tuple,
                ["Quantity"] = index + 1, ["Price"] = item.Item4 + tuple, ["Delta"] = 1
            })),
        Other = new[] { "X", "Y", "Z" }.Select(name => new Dictionary<string, object> { ["$level"] = 5, ["Name"] = name })
    });

    internal void AssertGenerated(byte[] bytes)
    {
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package));
        var workbook = package.WorkbookPart!;
        var report = workbook.WorksheetParts.Single(part => part.Worksheet!.GetFirstChild<SheetDimension>()!.Reference == "A1:K18").Worksheet!;
        var cells = report.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal(109, cells.Count); Assert.Equal(30, cells.Values.Count(cell => cell.CellFormula is not null));
        Assert.Equal("Lot 1", cells["A3"].InlineString!.InnerText);
        Assert.Equal("Lot 2", cells["A9"].InlineString!.InnerText);
        Assert.Equal("X", cells["A15"].InlineString!.InnerText);
        Assert.Equal("Z", cells["A17"].InlineString!.InnerText);
        Assert.Equal(new[] { "10", "11", "12", "13", "14", "15" }, new[] { "D4", "E4", "F4", "H4", "I4", "J4" }.Select(address => cells[address].CellValue!.Text));
        Assert.Equal(CellValues.Number, cells["D4"].DataType!.Value);
        Assert.Equal("SUBTOTAL(9,D$4:D$7)", cells["D3"].CellFormula!.Text);
        Assert.Equal("SUBTOTAL(9,J$10:J$11)", cells["J12"].CellFormula!.Text);
        Assert.Equal("SUBTOTAL(9,J3:J12)", cells["J13"].CellFormula!.Text);
        Assert.Equal(new[] { "C1:F1", "G1:J1", "A4:A5", "A6:A7", "A10:A11" }, report.Descendants<MergeCell>().Select(merge => merge.Reference!.Value));
        var summary = workbook.WorksheetParts.Single(part => part.Worksheet != report).Worksheet!;
        Assert.Equal("SUM(Report!C13:J13)", Assert.Single(summary.Descendants<Cell>()).CellFormula!.Text);
        Assert.All(workbook.WorksheetParts.SelectMany(part => part.Worksheet!.Descendants<Cell>()).Where(cell => cell.CellFormula is not null), cell =>
        {
            Assert.Empty(cell.CellFormula!.GetAttributes()); Assert.Null(cell.CellValue); Assert.Null(cell.DataType);
        });
        Assert.Null(workbook.CalculationChainPart);
        Assert.Equal(CalculateModeValues.Auto, workbook.Workbook!.CalculationProperties!.CalculationMode!.Value);
        Assert.True(workbook.Workbook.CalculationProperties.FullCalculationOnLoad!.Value);
    }

}
