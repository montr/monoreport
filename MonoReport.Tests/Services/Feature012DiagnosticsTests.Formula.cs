using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;
using MonoReport.Models.Xlsx;
using MonoReport.Services.Xlsx;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012DiagnosticsTests
{
    [Theory]
    [InlineData("{\"Items\":42}", "$.Items", "A3")]
    [InlineData("{\"Items\":[{\"$level\":10}]}\n", null, null)]
    public void Formula_Collections_reject_wrong_types_with_source_paths_and_missing_data_warnings_remain_ordered(string json, string? path, string? address)
    {
        var fixture = new XlsxFixture(); var bytes = fixture.CreateRepeated();
        if (path is not null)
        {
            var error = Assert.Throws<TemplateException>(() => fixture.Merge(bytes, json));
            Assert.Equal(DiagnosticCodes.GenerationInvalidData, error.Code); Assert.Contains(address!, error.Location);
            Assert.Equal(path, error.Errors[0].DataPath); return;
        }
        bytes = fixture.Create((_, part) =>
        {
            var data = part.Worksheet!.GetFirstChild<SheetData>()!; data.RemoveAllChildren();
            data.Append(new Row(fixture.Text("A3", "$l(10).Items"), fixture.Text("B3", "{Name}")) { RowIndex = 3 },
                new Row(fixture.Text("A5", "$l(10).Other"), fixture.Text("B5", "{Name}")) { RowIndex = 5 });
        });
        var result = fixture.Merge(bytes, "{\"Items\":[{\"$level\":10}],\"Other\":[{\"$level\":20}]}", new MergeOptions { MissingValues = MissingValueBehavior.Warning });
        var warning = Assert.Single(result.Warnings);
        Assert.Equal(DiagnosticCodes.XlsxMissingValue, warning.Code);
        Assert.Contains("B3", warning.Location);
    }
    [Fact]
    public void Formula_Repeated_field_diagnostics_keep_source_cell_before_input_instance_order()
    {
        var fixture = new XlsxFixture(); var twoAxis = new XlsxTwoAxisFixture();
        var bytes = twoAxis.Create((_, part) =>
        {
            part.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B10").InlineString!.Text!.Text = "{MissingA}";
            part.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "E10").InlineString!.Text!.Text = "{MissingB}";
        });
        var error = Assert.Throws<TemplateException>(() => fixture.Merge(bytes, twoAxis.Data(1, 2), new MergeOptions { MissingValues = MissingValueBehavior.Error }));
        Assert.Equal(new[] { "{MissingA}", "{MissingA}", "{MissingB}", "{MissingB}", "{MissingB}", "{MissingB}" }, error.Errors.Select(issue => issue.Field));
        Assert.Equal(new[] { "$.Items[0].MissingA", "$.Items[2].MissingA", "$.Items[0].MissingB", "$.Items[1].MissingB", "$.Items[2].MissingB", "$.Items[3].MissingB" }, error.Errors.Select(issue => issue.DataPath));
    }
    [Theory]
    [InlineData(99)] [InlineData(100)] [InlineData(101)]
    public void Formula_Repeated_value_preparation_stops_at_100_errors_without_a_101st_resolution(int count)
    {
        var fixture = new XlsxFixture();
        var bytes = fixture.Create((_, worksheet) =>
        {
            var row = worksheet.Worksheet!.Descendants<Row>().Single();
            row.PrependChild(fixture.Text("A3", "$l(10).Items"));
            row.Elements<Cell>().Single(cell => cell.CellReference!.Value == "B3").InlineString!.Text!.Text = "{Required}";
        });
        using var data = JsonDocument.Parse(JsonSerializer.Serialize(new { Items = Enumerable.Range(0, count).Select(_ => new Dictionary<string, object> { ["$level"] = 10 }) }));
        using var services = fixture.Services(); using var document = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions { MissingValues = MissingValueBehavior.Error });
        var planner = services.GetRequiredService<XlsxTemplatePlanner>(); var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
        var plan = planner.Plan(services.GetRequiredService<XlsxPackageReader>().Read(document, context), context);
        var bound = planner.BindSheet(plan.Sheets[0], data.RootElement, context);
        var output = geometry.PlanReports(plan, [bound], [planner.ScheduleSheet(plan.Sheets[0], bound, context)], context).Sheets[0];
        var error = Assert.Throws<TemplateException>(() => planner.BindInstances(plan.Sheets[0], output, bound, context));
        Assert.Equal(Math.Min(100, count), error.Errors.Count); Assert.Equal(Math.Min(100, count), context.ResolvedFields);
        Assert.Equal(count >= 100, error.LimitReached);
        Assert.Equal("$.Items[" + (Math.Min(100, count) - 1) + "].Required", error.Errors[^1].DataPath);
        Assert.Throws<TemplateException>(() => fixture.Merge(bytes, data.RootElement.GetRawText(), context.Options));
    }

    [Fact]
    public void Formula_Matrix_cardinality_and_unsafe_numeric_data_include_source_and_input_path()
    {
        var fixture = new XlsxFixture(); var twoAxis = new XlsxTwoAxisFixture(); var bytes = twoAxis.Create();
        var error = Assert.Throws<TemplateException>(() => fixture.Merge(bytes, "{\"Bids\":[{}],\"Criteria\":[{},{}],\"Items\":[{\"$level\":10,\"$path\":\"x\"}]}"));
        Assert.Equal(DiagnosticCodes.GenerationInvalidData, error.Code); Assert.Contains("A10", error.Location);
        Assert.Equal("$.Items[0]", error.Errors[0].DataPath); Assert.Contains("exactly 2", error.Errors[0].Message);
        var json = twoAxis.Data(1, 1, 1).Replace("\"Price\":100", "\"Price\":1234567890123456", StringComparison.Ordinal);
        error = Assert.Throws<TemplateException>(() => fixture.Merge(bytes, json));
        Assert.Equal(DiagnosticCodes.XlsxUnsafeNumber, error.Code); Assert.Contains("E10", error.Location);
        Assert.Equal("$.Items[0].Price", error.Errors[0].DataPath);
    }

    [Theory]
    [InlineData("INDIRECT(\"B3\")")]
    [InlineData("_xlfn.offset(B3,1,0)")]
    [InlineData("SUM(B:B)")]
    [InlineData("A1")]
    [InlineData("'Missing'!B3")]
    public void Formula_Unsafe_source_preflight_is_coded_actionable_and_shared_by_analyze_merge(string formula)
    {
        var fixture = new XlsxFixture();
        var bytes = fixture.Create((_, part) =>
        {
            var cell = part.Worksheet!.Descendants<Cell>().Single(); cell.RemoveAllChildren(); cell.DataType = null;
            cell.CellFormula = new CellFormula(formula);
        });
        using var services = fixture.Services(); using var input = new MemoryStream(bytes);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        var issue = Assert.Single(analysis.Errors); Assert.False(analysis.Valid);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code); Assert.Contains("B3", issue.Location);
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = fixture.Merge(bytes));
        Assert.Null(result); Assert.Equal(issue.Code, error.Code); Assert.Equal(issue.Message, error.Errors[0].Message);
        Assert.Equal(bytes, input.ToArray()); Assert.Contains("A1", issue.Message);
    }
}
