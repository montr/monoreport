using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012NestedTests
{
    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)]
    public void Regions_Identical_runs_keep_duplicate_items_and_one_authored_separator(int count)
    {
        var source = ProjectionTemplate([40, 40, 40], (_, part) =>
        {
            var separator = part.Worksheet!.Descendants<Row>().Single(row => row.RowIndex!.Value == 4);
            separator.RemoveAllChildren(); separator.Append(_fixture.Text("B4", "separator"));
            separator.Height = 0; separator.CustomHeight = true; separator.Hidden = true;
        });
        var json = JsonSerializer.Serialize(new { Items = new[] { new Dictionary<string, object> { ["$level"] = 20, ["Name"] = "Department" } }
            .Concat(Enumerable.Range(0, count).Select(_ => new Dictionary<string, object> { ["$level"] = 40, ["Name"] = "Ivan" })) });
        var result = RenderNested(source, json);
        Assert.Equal(new[] { (0, 3, 3), (1, 5, 5) }, result.Template.Sheets[0].Regions.Select(region => (region.Id, region.FirstRow, region.LastRow)));
        foreach (var region in result.Schedules[0].Regions)
        {
            Assert.Equal(count, region.Blocks.Length);
            Assert.Equal(Enumerable.Range(1, count).Select(index => "0," + index), region.Blocks.Select(block => string.Join(',', block.Owner.LogicalRows)));
            Assert.All(region.Blocks, block => Assert.Equal(region.RegionId, block.Owner.RegionId));
        }
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal(Enumerable.Repeat("Ivan", count).Append("separator").Concat(Enumerable.Repeat("Ivan", count)),
            worksheet.Descendants<Cell>().Select(cell => cell.InlineString!.InnerText));
        var retained = Assert.Single(worksheet.Descendants<Row>(), row => row.Height?.Value == 0);
        Assert.Equal((uint)count + 1, retained.RowIndex!.Value); Assert.True(retained.Hidden!.Value); Assert.True(retained.CustomHeight!.Value);
        Assert.Equal("A1:A" + (count * 2 + 1), worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Empty(result.Context.Errors); Assert.Empty(result.Context.Warnings);
    }

    [Theory]
    [InlineData(2, 2, 2)] [InlineData(2, 2, 1)] [InlineData(2, 2, 0)]
    [InlineData(0, 2, 2)] [InlineData(2, 0, 2)]
    public void Regions_All_nested_runs_and_worksheets_share_final_placement_without_sharing_data_trees(int bids, int criteria, int parents)
    {
        var source = RegionsWorkbook(); var original = source.ToArray();
        var arity = Math.Max(1, bids * criteria);
        var items = Enumerable.Range(0, parents).SelectMany(parent =>
            new[] { (Level: 10, Name: parent == 0 ? "P" : "Q", Price: 0) }
                .Concat(Enumerable.Range(0, parent == 0 ? 2 : 1).Select(child => (40, "C" + (parent * 2 + child + 1), (parent * 2 + child + 1) * 100))))
            .SelectMany((item, index) => Enumerable.Range(0, arity).Select(tuple => new Dictionary<string, object>
            { ["$level"] = item.Item1, ["$path"] = "item" + index, ["Name"] = item.Item2, ["Quantity"] = index + 1, ["Price"] = item.Item3 + tuple + 1 }));
        var json = JsonSerializer.Serialize(new
        {
            Bids = Enumerable.Range(0, bids).Select(index => new { Company = "Company" + index }),
            Criteria = Enumerable.Range(0, criteria).Select(index => new { Name = "Criterion" + index }), Items = items,
            Other = new[] { new Dictionary<string, object> { ["$level"] = 5, ["Name"] = "Other parent" },
                new Dictionary<string, object> { ["$level"] = 90, ["Name"] = "Other child 1" }, new Dictionary<string, object> { ["$level"] = 90, ["Name"] = "Other child 2" } },
            Third = new[] { new Dictionary<string, object> { ["$level"] = 50, ["Name"] = "Third parent" },
                new Dictionary<string, object> { ["$level"] = 80, ["Name"] = "Third child", ["Price"] = 17 } },
            Value = 7
        });
        var result = RenderNested(source, json);
        Assert.Equal(original, source); Assert.Empty(result.Context.Errors); Assert.Empty(result.Context.Warnings);
        Assert.Equal(new uint[] { 1, 9, 4 }, result.Template.Sheets.Select(sheet => sheet.Source.Id));
        Assert.Equal(new[] { 3, 0, 1 }, result.Template.Sheets.Select(sheet => sheet.Regions.Length));
        Assert.Equal(new[] { "$.Items", "$.Items", "$.Other" }, result.Template.Sheets[0].Regions.Select(region => region.Collection));
        if (parents > 0) Assert.NotSame(result.Bound[0].Rows[0][0], result.Bound[0].Rows[2][0]);
        Assert.Equal(5, result.Bound[0].Rows[2][0].Level); Assert.Equal(50, result.Bound[2].Rows[0][0].Level);
        foreach (var region in result.Schedules[0].Regions.Take(2))
        {
            Assert.All(region.Blocks, block => Assert.Equal(region.RegionId, block.Owner.RegionId));
            Assert.All(result.Bound[0].Rows[region.RegionId], row => Assert.Equal(arity, row.Records.Length));
            var expectedOwners = parents == 0 ? Array.Empty<string>() : parents == 1
                ? new[] { "0", "0," + arity, "0," + arity * 2, "0" }
                : new[] { "0", "0," + arity, "0," + arity * 2, "0", (arity * 3).ToString(), arity * 3 + "," + arity * 4, (arity * 3).ToString() };
            Assert.Equal(expectedOwners, region.Blocks.Select(block => string.Join(',', block.Owner.LogicalRows)));
        }
        Assert.All(result.Bound[0].Rows[2], row => Assert.Single(row.Records)); // Outside the worksheet's horizontal strip.
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var workbook = package.WorkbookPart!;
        var worksheets = workbook.Workbook!.Sheets!.Elements<Sheet>().Select(sheet =>
            ((WorksheetPart)workbook.GetPartById(sheet.Id!.Value!)).Worksheet!).ToArray();
        var main = worksheets[0]; var cells = main.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        var runHeight = parents == 0 ? 0 : parents == 1 ? 6 : 10;
        var last = runHeight * 2 + 9;
        var endColumn = bids == 0 ? "C" : criteria == 0 ? "E" : "I";
        Assert.Equal("A1:" + endColumn + last, main.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal(new[] { (uint)(runHeight + 3), (uint)(runHeight * 2 + 4) },
            main.Descendants<Row>().Where(row => row.Height?.Value == 0).Select(row => row.RowIndex!.Value));
        Assert.All(main.Descendants<Row>().Where(row => row.Height?.Value == 0), row => Assert.True(row.Hidden!.Value));
        var names = parents == 0 ? Array.Empty<string>() : parents == 1 ? new[] { "H:P", "C1", "", "C2", "", "F:P" }
            : new[] { "H:P", "C1", "", "C2", "", "F:P", "H:Q", "C3", "", "F:Q" };
        Assert.Equal(names.Append("separator 1").Concat(names).Append("separator 2")
            .Concat(new[] { "H:Other parent", "Other child 1/Other parent", "Other child 2/Other parent", "F:Other parent", "tail" }),
            Enumerable.Range(3, last - 2).Select(row => cells.GetValueOrDefault("A" + row)?.InlineString?.InnerText ?? ""));
        Assert.Equal("SUBTOTAL(9,B3:B" + (runHeight * 2 + 3) + ")", cells[endColumn + "2"].CellFormula!.Text);
        Assert.Equal("COUNT(B3:B" + (last - 1) + ")", cells[endColumn + last].CellFormula!.Text);
        var childRows = parents == 0 ? Array.Empty<int>() : parents == 1 ? new[] { 4, 6 } : new[] { 4, 6, 10 };
        foreach (var offset in new[] { 0, runHeight + 1 })
            for (var index = 0; index < childRows.Length; index++)
            {
                var row = childRows[index] + offset;
                Assert.Equal(18d, main.Descendants<Row>().Single(item => item.RowIndex!.Value == row).Height!.Value);
                Assert.Equal(31d, main.Descendants<Row>().Single(item => item.RowIndex!.Value == row + 1).Height!.Value);
                if (bids * criteria > 0)
                    foreach (var physicalRow in new[] { row, row + 1 })
                        Assert.Equal(Enumerable.Range((index + 1) * 100 + 1, 4).Select(value => value.ToString()),
                            new[] { "D", "E", "G", "H" }.Select(column => cells[column + physicalRow].CellValue!.Text));
            }
        var merges = main.Descendants<MergeCell>().Select(merge => merge.Reference!.Value!).ToArray();
        var expectedMerges = new List<string>();
        if (bids * criteria > 0) expectedMerges.AddRange(new[] { "C1:E1", "F1:H1" });
        foreach (var offset in new[] { 0, runHeight + 1 })
        {
            foreach (var row in parents == 0 ? Array.Empty<int>() : parents == 1 ? new[] { 3 } : new[] { 3, 9 }) expectedMerges.Add($"A{row + offset}:B{row + offset}");
            foreach (var row in childRows) expectedMerges.Add($"A{row + offset}:A{row + offset + 1}");
            foreach (var row in parents == 0 ? Array.Empty<int>() : parents == 1 ? new[] { 8 } : new[] { 8, 12 }) expectedMerges.Add($"A{row + offset}:B{row + offset}");
        }
        Assert.Equal(expectedMerges, merges);
        Assert.All(main.Descendants<Column>(), column => Assert.Equal(0d, column.Width!.Value));
        Assert.Equal(bids, main.Descendants<Column>().Count());
        Assert.Equal("A1:B2", worksheets[1].GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal("7", worksheets[1].Descendants<Cell>().Single(cell => cell.CellReference == "A1").CellValue!.Text);
        Assert.Equal("A1*2", worksheets[1].Descendants<Cell>().Single(cell => cell.CellReference == "B2").CellFormula!.Text);
        Assert.Equal("A1:B4", worksheets[2].GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal(new[] { "H:Third parent", "Third child", "F:Third parent" },
            worksheets[2].Descendants<Cell>().Where(cell => cell.CellReference!.Value!.StartsWith('A')).Select(cell => cell.InlineString!.InnerText));
        Assert.Equal("B2*2", worksheets[2].Descendants<Cell>().Single(cell => cell.CellReference == "B3").CellFormula!.Text);
        Assert.Equal("17", worksheets[2].Descendants<Cell>().Single(cell => cell.CellReference == "B2").CellValue!.Text);
        Assert.Equal("A2:A3", Assert.Single(worksheets[2].Descendants<MergeCell>()).Reference!.Value);
        Assert.Equal(new[] { "'Third'!$A$1:$B$4", "'Report'!$A$1:$" + endColumn + "$" + last, "'Static'!$A$1:$B$2" },
            workbook.Workbook.DefinedNames!.Elements<DefinedName>().Select(name => name.Text));
        Assert.All(worksheets.SelectMany(sheet => sheet.Descendants<Cell>()), cell => Assert.DoesNotContain("$l(", cell.InnerText));
        if (bids == 2 && criteria == 2 && parents == 2 && Environment.GetEnvironmentVariable("MONOREPORT_XLSX_EVIDENCE") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory); File.WriteAllBytes(Path.Combine(directory, "regions-template.xlsx"), source);
            File.WriteAllText(Path.Combine(directory, "regions-data.json"), json); File.WriteAllBytes(Path.Combine(directory, "regions-generated.xlsx"), result.Document);
        }
    }

    [Fact]
    public void Regions_Conflicting_side_by_side_row_owners_reject_before_any_output()
    {
        var source = _fixture.Create((workbook, part) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$C$3:$E$5";
            part.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A3", "$l(10).Items"), _fixture.Text("B3", "$l(10).Other"), _fixture.Text("C3", "{Name}")) { RowIndex = 3 },
                new Row(_fixture.Text("A4", "$l(40).Items"), _fixture.Text("B4", "$l(90).Other"), _fixture.Text("C4", "{Name}")) { RowIndex = 4 },
                new Row(_fixture.Text("A5", "$l(10).Items"), _fixture.Text("B5", "$l(10).Other"), _fixture.Text("C5", "{Name}")) { RowIndex = 5 }));
        });
        var original = source.ToArray(); using var services = _fixture.Services();
        using var stream = new MemoryStream(source); var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(stream);
        var issue = Assert.Single(analysis.Errors); Assert.Equal(DiagnosticCodes.TemplateInvalidStructure, issue.Code);
        Assert.Contains("worksheet rows", issue.Message); Assert.Contains("different worksheets", issue.Message); Assert.Contains("A3", issue.Location);
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source, """{"Items":[],"Other":[]}"""));
        Assert.Equal(issue.Code, error.Code); Assert.Equal(issue.Message, error.Errors[0].Message); Assert.Equal(issue.Location, error.Errors[0].Location);
        NestedRenderResult? output = null;
        Assert.Throws<TemplateException>(() => output = RenderNested(source, """{"Items":[],"Other":[]}"""));
        Assert.Null(output); Assert.Equal(original, source);
    }

    private byte[] RegionsWorkbook()
    {
        var source = NestedTupleTemplate();
        using var stream = new MemoryStream(); stream.Write(source); stream.Position = 0;
        using (var package = SpreadsheetDocument.Open(stream, true))
        {
            var workbook = package.WorkbookPart!; var main = workbook.WorksheetParts.Single().Worksheet!;
            var rows = main.GetFirstChild<SheetData>()!;
            rows.Elements<Row>().Single(row => row.RowIndex!.Value == 8).Append(new Cell { CellReference = "F8", CellFormula = new CellFormula("SUBTOTAL(9,C9:C17)") });
            var repeat = rows.Elements<Row>().Where(row => row.RowIndex!.Value is >= 9 and <= 12).Select(row => (Row)row.CloneNode(true)).ToArray();
            rows.Append(new Row(_fixture.Text("B13", "separator 1")) { RowIndex = 13, Height = 0, CustomHeight = true, Hidden = true });
            foreach (var row in repeat)
            {
                row.RowIndex = row.RowIndex!.Value + 5;
                foreach (var cell in row.Elements<Cell>()) cell.CellReference = new string(cell.CellReference!.Value!.TakeWhile(char.IsLetter).ToArray()) + row.RowIndex.Value;
                rows.Append(row);
            }
            rows.Append(new Row(_fixture.Text("B18", "separator 2")) { RowIndex = 18, Height = 0, CustomHeight = true, Hidden = true },
                new Row(_fixture.Text("A19", "$l(5).Other"), _fixture.Text("B19", "H:{Name}")) { RowIndex = 19 },
                new Row(_fixture.Text("A20", "$l(90).Other"), _fixture.Text("B20", "{Name}/{../Name}")) { RowIndex = 20 },
                new Row(_fixture.Text("A21", "$l(5).Other"), _fixture.Text("B21", "F:{Name}")) { RowIndex = 21 },
                new Row(_fixture.Text("B22", "tail"), new Cell { CellReference = "F22", CellFormula = new CellFormula("COUNT(C9:C21)") }) { RowIndex = 22 });
            main.GetFirstChild<MergeCells>()!.Append(new MergeCell { Reference = "B14:C14" }, new MergeCell { Reference = "B15:B16" }, new MergeCell { Reference = "B17:C17" });
            var staticPart = workbook.AddNewPart<WorksheetPart>();
            staticPart.Worksheet = new Worksheet(new SheetData(new Row(_fixture.Text("C4", "{Value}")) { RowIndex = 4 },
                new Row(new Cell { CellReference = "D5", CellFormula = new CellFormula("C4*2") }) { RowIndex = 5 }));
            var third = workbook.AddNewPart<WorksheetPart>();
            third.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A5", "$l(50).Third"), _fixture.Text("D5", "H:{Name}")) { RowIndex = 5 },
                new Row(_fixture.Text("A6", "$l(80).Third"), _fixture.Text("D6", "{Name}"), _fixture.Text("E6", "{Price}")) { RowIndex = 6 },
                new Row(_fixture.Text("A7", "$l(80).Third"), new Cell { CellReference = "E7", CellFormula = new CellFormula("E6*2") }) { RowIndex = 7 },
                new Row(_fixture.Text("A8", "$l(50).Third"), _fixture.Text("D8", "F:{Name}")) { RowIndex = 8 }));
            third.Worksheet.Append(new MergeCells(new MergeCell { Reference = "D6:D7" }));
            workbook.Workbook!.Sheets!.Append(new Sheet { Name = "Static", SheetId = 9, Id = workbook.GetIdOfPart(staticPart) }, new Sheet { Name = "Third", SheetId = 4, Id = workbook.GetIdOfPart(third) });
            workbook.Workbook.DefinedNames!.RemoveAllChildren();
            workbook.Workbook.DefinedNames.Append(new DefinedName("'Third'!$D$5:$E$8") { Name = "_xlnm.Print_Area", LocalSheetId = 2 },
                new DefinedName("'Report'!$B$7:$F$22") { Name = "_xlnm.Print_Area", LocalSheetId = 0 },
                new DefinedName("'Static'!$C$4:$D$5") { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
        }
        return stream.ToArray();
    }
}
