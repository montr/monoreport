using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012NestedTests
{
    [Theory]
    [InlineData(3, false)]
    [InlineData(3, true)]
    [InlineData(2, true)]
    public void Full_Parsed_nested_report_and_dedicated_two_by_two_intersections(int criteria, bool shared)
    {
        var source = FullNestedTemplate(shared); var original = source.ToArray();
        var json = FullNestedData(criteria);
        var result = RenderNested(source, json);
        var generated = _fixture.Merge(source, json);
        Assert.Equal(MonoReport.Models.DocumentFormat.Xlsx, generated.Format); Assert.Empty(generated.Warnings);
        result = result with { Document = generated.Document };
        Assert.Equal(original, source); Assert.Empty(result.Context.Errors); Assert.Empty(result.Context.Warnings);
        var main = result.Template.Sheets[0];
        Assert.Equal(new[] { (9, 12), (15, 15) }, main.Regions.Select(region => (region.FirstRow, region.LastRow)));
        Assert.Equal((10, 11), (main.Regions[0].Root.Children[0].Body!.FirstRow, main.Regions[0].Root.Children[0].Body!.LastRow));
        var arity = criteria * 2;
        Assert.Equal(5, result.Bound[0].Rows[0].Length);
        Assert.All(result.Bound[0].Rows[0], row => Assert.Equal(arity, row.Records.Length));
        Assert.Equal(arity * 5, result.Bound[0].Rows[0].Sum(row => row.Records.Length));
        Assert.All(result.Bound[0].Rows[1], row => Assert.Single(row.Records));
        Assert.Equal(new[] { "0", "0," + arity, "0," + arity * 2, "0", (arity * 3).ToString(), arity * 3 + "," + arity * 4, (arity * 3).ToString() },
            result.Schedules[0].Regions[0].Blocks.Select(block => string.Join(',', block.Owner.LogicalRows)));
        Assert.Equal(new[] { "0", "1", "2" }, result.Schedules[0].Regions[1].Blocks.Select(block => string.Join(',', block.Owner.LogicalRows)));

        var columns = criteria == 3 ? new[] { "D", "E", "F", "H", "I", "J" } : new[] { "D", "E", "G", "H" };
        var parentColumns = criteria == 3 ? new[] { "C", "G" } : new[] { "C", "F" };
        var lastColumn = criteria == 3 ? "K" : "I";
        var expected = new Dictionary<string, string>();
        foreach (var column in columns)
        {
            expected[column + "3"] = "SUBTOTAL(9," + column + "$4:" + column + "$7)";
            expected[column + "8"] = "SUBTOTAL(9," + column + "$4:" + column + "$7)";
            expected[column + "9"] = "SUBTOTAL(9," + column + "$10:" + column + "$11)";
            expected[column + "12"] = "SUBTOTAL(9," + column + "$10:" + column + "$11)";
            expected[column + "13"] = "SUBTOTAL(9," + column + "3:" + column + "12)";
        }
        var expectedTotals = criteria == 3 ? new[] { 63, 66, 69, 72, 75, 78 } : new[] { 63, 66, 69, 72 };
        var expectedFirstGroup = criteria == 3 ? new[] { 32, 34, 36, 38, 40, 42 } : new[] { 32, 34, 36, 38 };
        var expectedSecondGroup = criteria == 3 ? new[] { 31, 32, 33, 34, 35, 36 } : new[] { 31, 32, 33, 34 };
        AssertFullNestedPackage(result.Document, expected, lastColumn, criteria == 3 ? 109 : 83);
        using (var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false))
        {
            var workbook = package.WorkbookPart!;
            var worksheet = ((WorksheetPart)workbook.GetPartById(workbook.Workbook!.Sheets!.Elements<Sheet>().First().Id!.Value!)).Worksheet!;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal(new[] { "Alpha", "Beta" }, parentColumns.Select(column => cells[column + "1"].InlineString!.InnerText));
            Assert.Equal(2, cells.Values.Count(cell => cell.InlineString?.InnerText is "Alpha" or "Beta"));
            Assert.Equal(Enumerable.Range(0, 2).SelectMany(_ => Enumerable.Range(0, criteria).Select(index => "Criterion " + index)),
                columns.Select(column => cells[column + "2"].InlineString!.InnerText));
            Assert.Equal(new[] { "Lot 1", "A", "", "B", "", "Footer: Lot 1", "Lot 2", "C", "", "Footer: Lot 2", "Grand", "Separator", "X", "Y", "Z", "After second table" },
                Enumerable.Range(3, 16).Select(row => cells["A" + row].InlineString?.InnerText ?? ""));
            foreach (var (row, name, quantity, price) in new[] { (4, "A", 2, 10), (6, "B", 3, 20), (10, "C", 5, 30) })
            {
                Assert.Equal(quantity.ToString(), cells["B" + row].CellValue!.Text);
                Assert.Equal(CellValues.Number, cells["B" + row].DataType!.Value);
                // Parent-only columns use the first record, even above the second company's leaf tuples.
                Assert.All(parentColumns, column => Assert.Equal(name, cells[column + row].InlineString!.InnerText));
                Assert.Equal(row == 10 ? "Lot 2" : "Lot 1", cells["B" + (row + 1)].InlineString!.InnerText);
                Assert.Equal(Enumerable.Range(price, arity).Select(value => value.ToString()), columns.Select(column => cells[column + row].CellValue!.Text));
                Assert.All(columns, column => Assert.Equal("1", cells[column + (row + 1)].CellValue!.Text));
                Assert.Equal(1U, cells["A" + row].StyleIndex!.Value);
                Assert.Equal(2U, cells["A" + (row + 1)].StyleIndex!.Value);
                Assert.All(columns, column => Assert.Equal(3U, cells[column + row].StyleIndex!.Value));
            }
            // Independent numeric inputs establish expected SUBTOTAL results, without evaluating output formulas.
            Assert.Equal(expectedFirstGroup, columns.Select(column => new[] { 4, 5, 6, 7 }.Sum(row => int.Parse(cells[column + row].CellValue!.Text))));
            Assert.Equal(expectedSecondGroup, columns.Select(column => new[] { 10, 11 }.Sum(row => int.Parse(cells[column + row].CellValue!.Text))));
            Assert.Equal(expectedTotals, columns.Select(column => new[] { 4, 5, 6, 7, 10, 11 }.Sum(row => int.Parse(cells[column + row].CellValue!.Text))));
            Assert.Equal(criteria == 3 ? new[] { "C1:F1", "G1:J1", "A4:A5", "A6:A7", "A10:A11" }
                : new[] { "C1:E1", "F1:H1", "A4:A5", "A6:A7", "A10:A11" }, worksheet.Descendants<MergeCell>().Select(merge => merge.Reference!.Value));
            var zeroColumns = worksheet.Descendants<Column>().Where(column => column.Width?.Value == 0).ToArray();
            Assert.Equal(criteria == 3 ? new uint[] { 3, 7 } : new uint[] { 3, 6 }, zeroColumns.Select(column => column.Min!.Value));
            Assert.All(zeroColumns, column => { Assert.Equal(column.Min!.Value, column.Max!.Value); Assert.True(column.Hidden!.Value); });
            var separator = worksheet.Descendants<Row>().Single(row => row.RowIndex!.Value == 14);
            Assert.Equal(0d, separator.Height!.Value); Assert.True(separator.Hidden!.Value);
            Assert.Equal(new[] { 18d, 19d, 18d, 19d, 18d, 19d }, worksheet.Descendants<Row>()
                .Where(row => row.RowIndex!.Value is 4 or 5 or 6 or 7 or 10 or 11).Select(row => row.Height!.Value));
        }
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.True(analysis.Valid); Assert.Empty(analysis.Errors); Assert.Empty(analysis.Warnings);
        Assert.Equal(new[] { "Bids", "Criteria", "Items", "Other" }, analysis.Requirements.Select(requirement => requirement.Name));
        Assert.Equal(new[] { "Name", "Quantity", "Price", "Delta" }, analysis.Requirements[2].Fields!.Select(field => field.Name));
        Assert.Contains(analysis.Occurrences, occurrence => occurrence.Expression == "{../Name}" && occurrence.Path == "$.Items[].Name" && occurrence.Location!.EndsWith(", C11", StringComparison.Ordinal));
        Assert.Equal(original, input.ToArray());
        if (shared && criteria == 3)
        {
            using var roundTrip = new MemoryStream(); roundTrip.Write(result.Document); roundTrip.Position = 0;
            using (var package = SpreadsheetDocument.Open(roundTrip, true))
            {
                foreach (var part in package.WorkbookPart!.WorksheetParts) part.Worksheet!.Save();
                package.WorkbookPart.Workbook!.Save();
            }
            AssertFullNestedPackage(roundTrip.ToArray(), expected, lastColumn, 109);
        }
        if (Environment.GetEnvironmentVariable("MONOREPORT_XLSX_EVIDENCE") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory); var name = $"full-{criteria}-{(shared ? "shared" : "ordinary")}";
            File.WriteAllBytes(Path.Combine(directory, name + "-template.xlsx"), source);
            File.WriteAllText(Path.Combine(directory, name + "-data.json"), json);
            File.WriteAllBytes(Path.Combine(directory, name + "-generated.xlsx"), result.Document);
            File.WriteAllText(Path.Combine(directory, name + "-expected.json"), JsonSerializer.Serialize(new
            {
                Formulas = expected, GrandTotals = columns.Zip(expectedTotals).ToDictionary(pair => pair.First + "13", pair => pair.Second),
                FirstGroupTotals = columns.Zip(expectedFirstGroup).ToDictionary(pair => pair.First + "3", pair => pair.Second),
                SecondGroupTotals = columns.Zip(expectedSecondGroup).ToDictionary(pair => pair.First + "9", pair => pair.Second),
                SummaryFormula = "SUM(Report!C13:" + (criteria == 3 ? "J" : "H") + "13)", SummaryTotal = expectedTotals.Sum(),
                CalculationStatus = "Independent expectations only; numerical Excel result verification remains pending."
            }, new JsonSerializerOptions { WriteIndented = true }));
        }
    }

    [Theory]
    [InlineData("short-matrix", DiagnosticCodes.GenerationInvalidData)]
    [InlineData("reappearing-path", DiagnosticCodes.GenerationInvalidData)]
    [InlineData("no-first-children", DiagnosticCodes.XlsxUnsafeFormula)]
    public void Full_Invalid_matrix_or_missing_descendants_returns_no_nested_workbook(string scenario, string code)
    {
        var source = FullNestedTemplate(shared: true); var original = source.ToArray();
        using var data = JsonDocument.Parse(FullNestedData(2));
        var records = data.RootElement.GetProperty("Items").EnumerateArray().ToList();
        if (scenario == "short-matrix") records.RemoveAt(7);
        else if (scenario == "reappearing-path") records.AddRange(records.Take(4).ToArray());
        else records.RemoveRange(4, 8);
        var json = JsonSerializer.Serialize(new { Bids = data.RootElement.GetProperty("Bids"), Criteria = data.RootElement.GetProperty("Criteria"),
            Items = records, Other = data.RootElement.GetProperty("Other") });
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, json));
        Assert.Null(result); Assert.Equal(original, source);
        Assert.All(error.Errors, issue => Assert.Equal(code, issue.Code));
        Assert.Contains(error.Errors, issue => issue.Location!.Contains(scenario == "no-first-children" ? "E9" : "A9", StringComparison.Ordinal));
        if (scenario == "no-first-children")
            Assert.Contains(error.Errors, issue => issue.Message.Contains("E$10:E$11", StringComparison.Ordinal) && issue.Message.Contains("retained contiguous A1 range", StringComparison.Ordinal));
        else Assert.Contains(error.Errors, issue => issue.DataPath?.StartsWith("$.Items[", StringComparison.Ordinal) == true);
    }

    [Theory]
    [InlineData("cropped-static-formula", DiagnosticCodes.XlsxUnsafeFormula, "B4")]
    [InlineData("header-parent", DiagnosticCodes.TemplateInvalidContext, "E8")]
    [InlineData("affected-feature", DiagnosticCodes.XlsxUnsupportedFeature, "Report")]
    public void Full_Analyzer_keeps_nested_source_safety_checks(string scenario, string code, string location)
    {
        var source = FullNestedTemplate(shared: true, customize: (workbook, part) =>
        {
            if (scenario == "cropped-static-formula") workbook.WorksheetParts.Single(sheet => sheet != part).Worksheet!.Descendants<Cell>().Single().CellFormula!.Text = "SUM(Report!E6:E11)";
            else if (scenario == "header-parent") part.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference == "E8").InlineString!.Text!.Text = "{../../../Name}";
            else part.Worksheet!.Append(new ConditionalFormatting(new ConditionalFormattingRule { Type = ConditionalFormatValues.Expression, Priority = 1 }) { SequenceOfReferences = new() { InnerText = "E10:E11" } });
        });
        var original = source.ToArray(); using var services = _fixture.Services(); using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(analysis.Valid); Assert.Empty(analysis.Requirements); Assert.Empty(analysis.Occurrences);
        Assert.Contains(analysis.Errors, issue => issue.Code == code && issue.Location!.Contains(location, StringComparison.Ordinal));
        if (scenario == "affected-feature") Assert.Contains(analysis.Errors, issue => issue.Message.Contains("conditionalFormatting", StringComparison.Ordinal));
        Assert.All(analysis.Errors, issue => Assert.False(string.IsNullOrWhiteSpace(issue.Message)));
        Assert.Equal(original, input.ToArray());
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, FullNestedData(2)));
        Assert.Null(result); Assert.Equal(original, source);
        Assert.Contains(error.Errors, issue => issue.Code == code && issue.Location!.Contains(location, StringComparison.Ordinal));
    }

    private void AssertFullNestedPackage(byte[] bytes, Dictionary<string, string> expected, string lastColumn, int cellCount)
    {
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var workbook = package.WorkbookPart!;
        Assert.Equal(new[] { "Report", "Summary" }, workbook.Workbook!.Sheets!.Elements<Sheet>().Select(sheet => sheet.Name!.Value));
        var sheets = workbook.Workbook.Sheets!.Elements<Sheet>().Select(sheet => ((WorksheetPart)workbook.GetPartById(sheet.Id!.Value!)).Worksheet!).ToArray();
        Assert.Equal("A1:" + lastColumn + "18", sheets[0].GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal(Enumerable.Range(1, 18).Select(row => (uint)row), sheets[0].Descendants<Row>().Select(row => row.RowIndex!.Value));
        Assert.Equal(cellCount, sheets[0].Descendants<Cell>().Count());
        var formulas = sheets[0].Descendants<Cell>().Where(cell => cell.CellFormula is not null).ToDictionary(cell => cell.CellReference!.Value!, cell => cell.CellFormula!.Text);
        Assert.Equal(expected.OrderBy(pair => pair.Key), formulas.OrderBy(pair => pair.Key));
        Assert.Equal("A1:A1", sheets[1].GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal("SUM(Report!C13:" + (lastColumn == "K" ? "J" : "H") + "13)", Assert.Single(sheets[1].Descendants<Cell>()).CellFormula!.Text);
        Assert.Equal(new[] { "'Report'!$A$1:$" + lastColumn + "$18", "'Summary'!$A$1:$A$1" }, workbook.Workbook.DefinedNames!.Elements<DefinedName>().Select(name => name.Text));
        foreach (var cell in sheets.SelectMany(sheet => sheet.Descendants<Cell>()).Where(cell => cell.CellFormula is not null))
        {
            Assert.Empty(cell.CellFormula!.GetAttributes()); Assert.Null(cell.CellValue); Assert.Null(cell.InlineString); Assert.Null(cell.DataType);
        }
        Assert.Null(workbook.CalculationChainPart);
        Assert.Equal(CalculateModeValues.Auto, workbook.Workbook.CalculationProperties!.CalculationMode!.Value);
        Assert.True(workbook.Workbook.CalculationProperties.FullCalculationOnLoad!.Value);
        Assert.True(workbook.Workbook.CalculationProperties.ForceFullCalculation!.Value);
        Assert.All(sheets, sheet => { Assert.DoesNotContain("DISPOSABLE", sheet.InnerText); Assert.DoesNotContain("$l(", sheet.InnerText); Assert.DoesNotContain("$h(", sheet.InnerText); });
    }

    private byte[] FullNestedTemplate(bool shared, Action<WorkbookPart, WorksheetPart>? customize = null) =>
        new XlsxNestedFixture().Create(shared, customize);

    private static string FullNestedData(int criteria) => new XlsxNestedFixture().Data(criteria);
}
