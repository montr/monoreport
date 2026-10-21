using System.Collections.Immutable;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

internal sealed class XlsxTwoAxisFixture
{
    private readonly XlsxFixture _fixture = new();
    internal byte[] Create(Action<WorkbookPart, WorksheetPart>? customize = null) => _fixture.Create((workbook, worksheet) =>
    {
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$7:$F$12";
        worksheet.Worksheet = new Worksheet(new Columns(new Column { Min = 4, Max = 4, Width = 0, CustomWidth = true }),
            new SheetData(
                new Row(_fixture.Text("D1", "$h(1)"), _fixture.Text("E1", "$header(1)")) { RowIndex = 1 },
                new Row(_fixture.Text("E2", "$h(2)")) { RowIndex = 2 },
                new Row(_fixture.Text("A7", "$h(1).Bids"), _fixture.Text("D7", "{Company}"), new Cell { CellReference = "E7" }) { RowIndex = 7, Height = 24, CustomHeight = true },
                new Row(_fixture.Text("A8", "$header(2).Criteria"), _fixture.Text("E8", "{Name}")) { RowIndex = 8 },
                new Row(_fixture.Text("B9", "above")) { RowIndex = 9 },
                new Row(_fixture.Text("A10", "$l(10).Items"), _fixture.Text("B10", "{Name}"), _fixture.Text("C10", "{Quantity}"), _fixture.Text("D10", "{Name}"), _fixture.Text("E10", "{Price}")) { RowIndex = 10, Height = 18, CustomHeight = true },
                new Row(_fixture.Text("B11", "below")) { RowIndex = 11 },
                new Row(_fixture.Text("B12", "tail")) { RowIndex = 12 }),
            new MergeCells(new MergeCell { Reference = "D7:E7" }));
        customize?.Invoke(workbook, worksheet);
    });

    internal string Data(int criteria = 3, int rows = 2, int companies = 2) => JsonSerializer.Serialize(new
    {
        Bids = Enumerable.Range(0, companies).Select(index => new { Company = index == 0 ? "Alpha" : "Beta" }),
        Criteria = Enumerable.Range(0, criteria).Select(index => new { Name = "Criterion " + index }),
        Items = Enumerable.Range(0, rows).SelectMany(row => Enumerable.Range(0, Math.Max(1, companies * criteria)).Select(tuple =>
            new Dictionary<string, object> { ["$level"] = 10, ["$path"] = "row" + row, ["Name"] = tuple == 0 ? "Row " + row : "tuple " + tuple, ["Quantity"] = row + 1, ["Price"] = (row + 1) * 100 + tuple }))
    });

    internal XlsxTemplatePlan Plan(byte[] bytes)
    {
        using var document = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        using var services = _fixture.Services(); var context = new XlsxOperationContext(new MergeOptions());
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var plan = planner.Prepare(services.GetRequiredService<XlsxPackageReader>().Read(document, context), context);
        return plan with { Sheets = plan.Sheets.Select(planner.Structure).ToImmutableArray() };
    }
}
