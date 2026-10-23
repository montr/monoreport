using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012FormulaTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CrossSheet_Static_workbook_uses_target_crop_and_keeps_sheet_identity_and_qualifiers(bool reverse)
    {
        const string formula = "SUM(Alpha!$H8:I$9)+'O''Brien Q4'!$B$3+LEN(\"Alpha!H8\")+1E10";
        var source = SheetIdentityTemplate(formula, reverse, targetArea: "$H$8:$J$10");
        using var services = _fixture.Services();
        using var input = new MemoryStream(source);
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
        var result = _fixture.Merge(source, "{}");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Assert.Equal("SUM(Alpha!$A1:B$2)+'O''Brien Q4'!$A$1+LEN(\"Alpha!H8\")+1E10",
            CrossSheetCells(package, "Zulu")["A1"].CellFormula!.Text);
        var sheets = package.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToArray();
        Assert.Equal(reverse ? new uint[] { 41, 5, 73 } : new uint[] { 73, 5, 41 }, sheets.Select(sheet => sheet.SheetId!.Value));
        using var original = SpreadsheetDocument.Open(new MemoryStream(source), false);
        Assert.Equal(original.WorkbookPart!.Workbook!.Sheets!.OuterXml, package.WorkbookPart.Workbook.Sheets.OuterXml);
        Assert.Equal("'Alpha'!$A$1:$C$3", package.WorkbookPart.Workbook.DefinedNames!.Elements<DefinedName>().Single(name =>
            name.LocalSheetId!.Value == (reverse ? 0u : 2u)).Text);
        Assert.Empty(result.Warnings);
    }

    [Theory]
    [InlineData(1, 1, 1)]
    [InlineData(2, 3, 3)]
    [InlineData(3, 0, 2)]
    public void CrossSheet_Static_expanded_and_independent_expanded_sheets_use_final_target_geometry(int rows, int criteria, int consumers)
    {
        // With no child columns, author a range over the retained parent column explicitly.
        var source = CrossSheetTemplate(bodyLastColumn: criteria == 0 ? "D" : "E");
        var json = CrossSheetData(rows, criteria, consumers);
        var result = _fixture.Merge(source, json);
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var summary = CrossSheetCells(package, "O'Brien Summary");
        var report = CrossSheetCells(package, "Report");
        var consumer = CrossSheetCells(package, "Consumers");
        var lastColumn = criteria == 0 ? "D" : criteria == 1 ? "F" : "J";
        var tailColumn = criteria == 0 ? "E" : criteria == 1 ? "G" : "K";
        Assert.Equal($"SUM(Report!$C4:{lastColumn}${rows + 3})+Report!${tailColumn}${rows + 5}+LEN(\"Report!D10\")+1E10", summary["A2"].CellFormula!.Text);
        Assert.Equal($"SUBTOTAL(9,Report!C4:{lastColumn}{rows + 3})", summary["B2"].CellFormula!.Text);
        Assert.Equal("'O''Brien Summary'!$B$1", summary["C2"].CellFormula!.Text);
        for (var row = 4; row < 4 + rows; row++)
        {
            Assert.Equal($"'O''Brien Summary'!A1+'O''Brien Summary'!$B$1+'O''Brien Summary'!A$1+'O''Brien Summary'!$B1+Report!B{row}+Consumers!A{consumers + 5}",
                report[tailColumn + row].CellFormula!.Text);
            Assert.Equal((row - 3).ToString(), report["B" + row].CellValue!.Text);
        }
        for (var row = 2; row < 2 + consumers; row++)
        {
            Assert.Equal($"COUNT(Report!$C$4:${lastColumn}${rows + 3})+'O''Brien Summary'!$B1", consumer["B" + row].CellFormula!.Text);
            Assert.Equal($"Report!{tailColumn}{rows + 5}+Report!B3", consumer["C" + row].CellFormula!.Text);
        }
        Assert.Equal("consumer tail", consumer["A" + (consumers + 5)].InlineString!.InnerText);
        Assert.Equal("tail", report["A" + (rows + 5)].InlineString!.InnerText);
        Assert.Equal(new[] { "O'Brien Summary", "Consumers", "Report" }, package.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Select(sheet => sheet.Name!.Value));
        Assert.Equal(new uint[] { 90, 7, 31 }, package.WorkbookPart.Workbook.Sheets.Elements<Sheet>().Select(sheet => sheet.SheetId!.Value));
        Assert.All(package.WorkbookPart.WorksheetParts.SelectMany(part => part.Worksheet!.Descendants<Cell>()).Where(cell => cell.CellFormula is not null), cell =>
        {
            Assert.Null(cell.CellValue);
            Assert.Null(cell.CellFormula!.FormulaType);
            Assert.Null(cell.CellFormula.SharedIndex);
        });
        Assert.True(package.WorkbookPart.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);
        Assert.Empty(result.Warnings);
        if (rows == 2 && criteria == 3 && Environment.GetEnvironmentVariable("MONOREPORT_XLSX_EVIDENCE") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, "cross-sheet-template.xlsx"), source);
            File.WriteAllText(Path.Combine(directory, "cross-sheet-data.json"), json);
            File.WriteAllBytes(Path.Combine(directory, "cross-sheet-generated.xlsx"), result.Document);
        }
    }

    [Theory]
    [InlineData("Report!C10", 2, 3, false)]
    [InlineData("Report!$C$10", 2, 3, false)]
    [InlineData("Report!E10:E10", 2, 3, false)]
    [InlineData("Report!C10:C10", 0, 3, false)]
    [InlineData("Report!E10", 1, 0, false)]
    [InlineData("Report!A1", 2, 3, false)]
    [InlineData("Report!A9:C10", 2, 3, false)]
    [InlineData("Report!E7", 2, 0, true)]
    [InlineData("Report!D10:E10", 3, 0, false)]
    [InlineData("Report!E10:D10", 3, 0, false)]
    public void CrossSheet_Unsafe_target_images_reject_with_source_operand_and_remedy(string operand, int rows, int criteria, bool childFirst)
    {
        var source = CrossSheetTemplate("SUM(" + operand + ")", childFirst);
        using var services = _fixture.Services();
        using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        if (operand.Contains("A1", StringComparison.Ordinal) || operand.Contains("A9", StringComparison.Ordinal))
        {
            var issue = Assert.Single(analysis.Errors);
            Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
            Assert.Contains(operand, issue.Message);
        }
        else Assert.True(analysis.Valid); // Analysis cannot invent target repetition counts.
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, CrossSheetData(rows, criteria, 3)));
        Assert.Null(result);
        var failing = Assert.Single(error.Errors, issue => issue.Location!.EndsWith("H9", StringComparison.Ordinal));
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, failing.Code);
        Assert.Contains("sheet 'O'Brien Summary'", failing.Location);
        Assert.Contains(operand, failing.Message);
        Assert.Contains("empty, discarded, disjoint or ambiguous", failing.Message);
        Assert.Contains("retained contiguous A1 range", failing.Message);
        Assert.Contains("scalar operands must identify one cell", failing.Message);
        Assert.Equal(source, input.ToArray());
    }

    [Theory]
    [InlineData("Report!C10", 1, true)]
    [InlineData("Report!$C$10", 1, true)]
    [InlineData("Report!C10", 2, false)]
    [InlineData("Report!$C$10", 2, false)]
    public void CrossSheet_Independent_equal_owner_indices_never_select_an_arbitrary_repeated_scalar(string operand, int rows, bool valid)
    {
        var source = CrossSheetTemplate(consumerFormula: operand);
        if (valid)
        {
            var output = _fixture.Merge(source, CrossSheetData(rows, 1, 1));
            using var package = SpreadsheetDocument.Open(new MemoryStream(output.Document), false);
            Assert.Empty(new OpenXmlValidator().Validate(package));
            Assert.Equal(operand.Contains('$') ? "Report!$B$4" : "Report!B4", CrossSheetCells(package, "Consumers")["B2"].CellFormula!.Text);
            return;
        }
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, CrossSheetData(rows, 1, 1)));
        Assert.Null(result);
        var issue = Assert.Single(error.Errors);
        Assert.Contains("sheet 'Consumers', E5", issue.Location);
        Assert.Contains(operand, issue.Message);
        Assert.Contains("scalar operands must identify one cell", issue.Message);
    }

    [Fact]
    public void CrossSheet_Removed_source_formulas_do_not_require_missing_target_instances()
    {
        var source = CrossSheetTemplate("1", consumerFormula: "Report!C10");
        var result = _fixture.Merge(source, CrossSheetData(0, 1, 0));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Assert.DoesNotContain(CrossSheetCells(package, "Consumers").Values, cell => cell.CellFormula is not null);
        Assert.DoesNotContain(CrossSheetCells(package, "Report").Values, cell => cell.CellFormula is not null);
        Assert.Equal("1", CrossSheetCells(package, "O'Brien Summary")["A2"].CellFormula!.Text);
    }

    [Fact]
    public void CrossSheet_Retained_merge_anchors_follow_content_owners_when_their_physical_source_columns_disappear()
    {
        var source = CrossSheetTemplate("SUM(Report!D7:D7)", childFirst: true,
            consumerFormula: "COUNT(Report!$D$7:$D$7)+'O''Brien Summary'!$I8");
        var result = _fixture.Merge(source, CrossSheetData(2, 0, 1));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Assert.Equal("SUM(Report!C1:D1)", CrossSheetCells(package, "O'Brien Summary")["A2"].CellFormula!.Text);
        Assert.Equal("COUNT(Report!$C$1:$D$1)+'O''Brien Summary'!$B1", CrossSheetCells(package, "Consumers")["B2"].CellFormula!.Text);
        Assert.Equal("Alpha", CrossSheetCells(package, "Report")["C1"].InlineString!.InnerText);
        Assert.Equal("Beta", CrossSheetCells(package, "Report")["D1"].InlineString!.InnerText);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    public void CrossSheet_Qualified_self_leaf_uses_its_own_horizontal_parent_and_vertical_instance(int criteria)
    {
        var source = CrossSheetTemplate(leafFormula: "'O''Brien Summary'!H8+Report!$C10+'Report'!D7");
        var result = _fixture.Merge(source, CrossSheetData(2, criteria, 1));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var cells = CrossSheetCells(package, "Report");
        for (var row = 4; row <= 5; row++)
            for (var parent = 0; parent < 2; parent++)
                for (var child = 1; child <= criteria; child++)
                {
                    var anchor = (char)('C' + parent * (criteria + 1));
                    var column = (char)(anchor + child);
                    Assert.Equal($"'O''Brien Summary'!A1+Report!$B{row}+'Report'!{anchor}1", cells[column.ToString() + row].CellFormula!.Text);
                }
    }

    [Theory]
    [InlineData("C8:C13", true)]
    [InlineData("C10:C11", false)]
    public void CrossSheet_Static_totals_over_nested_targets_require_complete_contiguous_images(string range, bool valid)
    {
        var source = GroupTemplate(customize: (workbook, _) =>
        {
            var summary = workbook.AddNewPart<WorksheetPart>();
            summary.Worksheet = new Worksheet(new SheetData(new Row(GroupFormula("B3", "SUM(Report!" + range + ")")) { RowIndex = 3 }));
            workbook.Workbook!.Sheets!.PrependChild(new Sheet { Name = "Summary", SheetId = 7, Id = workbook.GetIdOfPart(summary) });
            workbook.Workbook.DefinedNames!.Elements<DefinedName>().Single().LocalSheetId = 1;
            workbook.Workbook.DefinedNames.Append(new DefinedName("'Summary'!$B$3:$B$3") { Name = "_xlnm.Print_Area", LocalSheetId = 0 });
        });
        if (valid)
        {
            var result = RenderGroups(source, GroupData());
            using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
            Assert.Empty(new OpenXmlValidator().Validate(package));
            Assert.Equal("SUM(Report!B2:B19)", CrossSheetCells(package, "Summary")["A1"].CellFormula!.Text);
        }
        else
        {
            var error = Assert.Throws<TemplateException>(() => RenderGroups(source, GroupData()));
            var issue = Assert.Single(error.Errors);
            Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
            Assert.Contains("sheet 'Summary', B3", issue.Location);
            Assert.Contains("Report!" + range, issue.Message);
            Assert.Contains("retained contiguous A1 range", issue.Message);
        }
    }

    [Fact]
    public void CrossSheet_Qualified_self_references_keep_descendant_group_and_leaf_scopes()
    {
        var source = GroupTemplate(customize: (_, part) =>
        {
            var cells = part.Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            cells["C9"].CellFormula!.Text = "SUM('Report'!$C10:C$11)";
            cells["C12"].CellFormula!.Text = "SUBTOTAL(9,Report!C10:C11)";
            cells["C11"].CellFormula!.Text = "Report!C10*2";
        });
        var result = RenderGroups(source, GroupData());
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var cells = CrossSheetCells(package, "Report");
        Assert.Equal("SUM('Report'!$B4:B$7)", cells["B3"].CellFormula!.Text);
        Assert.Equal("SUM('Report'!$B10:B$11)", cells["B9"].CellFormula!.Text);
        Assert.Equal("SUM('Report'!$B16:B$17)", cells["B15"].CellFormula!.Text);
        Assert.Equal("SUBTOTAL(9,Report!B4:B7)", cells["B8"].CellFormula!.Text);
        Assert.Equal("SUBTOTAL(9,Report!B10:B11)", cells["B12"].CellFormula!.Text);
        Assert.Equal("SUBTOTAL(9,Report!B16:B17)", cells["B18"].CellFormula!.Text);
        Assert.Equal(new[] { "Report!B4*2", "Report!B6*2", "Report!B10*2", "Report!B16*2" }, new[] { 5, 7, 11, 17 }.Select(row => cells["B" + row].CellFormula!.Text));
        using var publicPackage = SpreadsheetDocument.Open(new MemoryStream(_fixture.Merge(source, GroupData()).Document), false);
        var publicCells = CrossSheetCells(publicPackage, "Report");
        Assert.Equal(cells.Where(pair => pair.Value.CellFormula is not null).Select(pair => (pair.Key, pair.Value.CellFormula!.Text)).Order(),
            publicCells.Where(pair => pair.Value.CellFormula is not null).Select(pair => (pair.Key, pair.Value.CellFormula!.Text)).Order());
    }

    private byte[] CrossSheetTemplate(string? summaryFormula = null, bool childFirst = false, string? consumerFormula = null,
        string? leafFormula = null, string bodyLastColumn = "E") =>
        new XlsxTwoAxisFixture().Create((workbook, report) =>
        {
            var rows = report.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
            if (leafFormula is not null)
            {
                var cell = rows[10].Elements<Cell>().Last();
                cell.RemoveAllChildren(); cell.DataType = null; cell.CellFormula = new CellFormula(leafFormula);
            }
            rows[10].Append(GroupFormula("F10", "'O''Brien Summary'!H8+'O''Brien Summary'!$I$8+'O''Brien Summary'!H$8+'O''Brien Summary'!$I8+Report!C10+Consumers!D8"));
            if (childFirst)
            {
                rows[2].Elements<Cell>().Single().CellReference = "D2";
                rows[8].Elements<Cell>().Last().CellReference = "D8";
            }
            var consumers = workbook.AddNewPart<WorksheetPart>();
            consumers.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("D4", "consumer top")) { RowIndex = 4 },
                new Row(_fixture.Text("C5", "$l(70).Consumers"), _fixture.Text("D5", "{Name}"),
                    GroupFormula("E5", consumerFormula ?? $"COUNT(Report!$D$10:${bodyLastColumn}$10)+'O''Brien Summary'!$I8"), GroupFormula("F5", "Report!F12+Report!C9")) { RowIndex = 5 },
                new Row(_fixture.Text("D6", "separator")) { RowIndex = 6 },
                new Row(_fixture.Text("C7", "$l(90).Later"), _fixture.Text("D7", "{Name}")) { RowIndex = 7 },
                new Row(_fixture.Text("D8", "consumer tail")) { RowIndex = 8 }));
            var summary = workbook.AddNewPart<WorksheetPart>();
            summary.Worksheet = new Worksheet(new SheetData(
                new Row(new Cell { CellReference = "H8", CellValue = new CellValue("7") }, new Cell { CellReference = "I8", CellValue = new CellValue("11") }) { RowIndex = 8 },
                new Row(GroupFormula("H9", summaryFormula ?? $"SUM(Report!$D10:{bodyLastColumn}$10)+Report!$F$12+LEN(\"Report!D10\")+1E10"),
                    GroupFormula("I9", summaryFormula is null ? $"SUBTOTAL(9,Report!D10:{bodyLastColumn}10)" : "1"), GroupFormula("J9", "'O''Brien Summary'!$I$8")) { RowIndex = 9 }));
            workbook.Workbook!.Sheets!.RemoveAllChildren();
            workbook.Workbook.Sheets.Append(new Sheet { Name = "O'Brien Summary", SheetId = 90, Id = workbook.GetIdOfPart(summary) },
                new Sheet { Name = "Consumers", SheetId = 7, Id = workbook.GetIdOfPart(consumers) }, new Sheet { Name = "Report", SheetId = 31, Id = workbook.GetIdOfPart(report) });
            workbook.Workbook.DefinedNames!.RemoveAllChildren();
            workbook.Workbook.DefinedNames.Append(new DefinedName("'O''Brien Summary'!$H$8:$J$9") { Name = "_xlnm.Print_Area", LocalSheetId = 0 },
                new DefinedName("'Consumers'!$D$4:$F$8") { Name = "_xlnm.Print_Area", LocalSheetId = 1 },
                new DefinedName("'Report'!$B$7:$F$12") { Name = "_xlnm.Print_Area", LocalSheetId = 2 });
        });

    private string CrossSheetData(int rows, int criteria, int consumers)
    {
        var twoAxis = new XlsxTwoAxisFixture();
        using var json = JsonDocument.Parse(twoAxis.Data(criteria, rows));
        return JsonSerializer.Serialize(new
        {
            Bids = json.RootElement.GetProperty("Bids"), Criteria = json.RootElement.GetProperty("Criteria"), Items = json.RootElement.GetProperty("Items"),
            Consumers = Enumerable.Range(0, consumers).Select(index => new Dictionary<string, object> { ["$level"] = 70, ["Name"] = "consumer " + index }),
            Later = Enumerable.Range(0, 2).Select(index => new Dictionary<string, object> { ["$level"] = 90, ["Name"] = "later " + index })
        });
    }

    private Dictionary<string, Cell> CrossSheetCells(SpreadsheetDocument package, string name)
    {
        var sheet = package.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name == name);
        return ((WorksheetPart)package.WorkbookPart.GetPartById(sheet.Id!)).Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
    }
}
