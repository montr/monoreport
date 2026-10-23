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
    [InlineData("depth", false)] [InlineData("depth", true)]
    [InlineData("children", false)] [InlineData("children", true)]
    [InlineData("header-rows", false)] [InlineData("header-rows", true)]
    public void Activation_Deferred_horizontal_shapes_reject_before_binding_but_keep_source_analysis(string shape, bool empty)
    {
        var source = ActivationTemplate(shape); var original = source.ToArray();
        using var services = _fixture.Services(); using var input = new MemoryStream(source);
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, ActivationData(empty)));
        Assert.Null(result); Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, error.Code);
        Assert.Contains(shape == "header-rows" ? "one physical horizontal header row" : "at most two horizontal levels", error.Message);
        Assert.Contains("move that layout", error.Message);
        Assert.Contains("Report", Assert.Single(error.Errors).Location);
        Assert.Equal(original, source); Assert.Equal(original, input.ToArray());
        // Null collections cannot bypass the source eligibility boundary either.
        Assert.Equal(error.Message, Assert.Throws<TemplateException>(() => _fixture.Merge(source, "{\"Items\":null}")).Message);
    }

    [Theory]
    [InlineData("depth")] [InlineData("children")] [InlineData("header-rows")]
    public void Activation_Deferred_horizontal_shapes_do_not_restrict_existing_flat_worksheets(string shape)
    {
        var source = ActivationTemplate(shape, nested: false);
        var result = _fixture.Merge(source, ActivationData(false).Replace("\"$level\":20", "\"$level\":10", StringComparison.Ordinal));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToArray();
        Assert.Contains(cells, cell => cell.InlineString?.InnerText == "parent");
        Assert.Contains(cells, cell => cell.InlineString?.InnerText == "child");
    }

    [Theory]
    [InlineData("rich-field")] [InlineData("pane")] [InlineData("break")]
    [InlineData("unsupported-formula")] [InlineData("partial-merge")]
    public void Activation_Public_nested_generation_preserves_format_geometry_and_formula_rejection(string scenario)
    {
        var source = FullNestedTemplate(true, (_, part) =>
        {
            var cells = part.Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            if (scenario == "rich-field") cells["B10"].InlineString = new InlineString(new Run(new Text("{Name}")));
            else if (scenario == "pane") part.Worksheet.PrependChild(new SheetViews(new SheetView(new Pane { TopLeftCell = "B7", State = PaneStateValues.Frozen, VerticalSplit = 1 }) { WorkbookViewId = 0 }));
            else if (scenario == "break") part.Worksheet.Append(new RowBreaks(new Break { Id = 10, Min = 0, Max = 16383, ManualPageBreak = true }) { Count = 1, ManualBreakCount = 1 });
            else if (scenario == "unsupported-formula") cells["E13"].CellFormula!.Text = "INDIRECT(\"E10\")";
            else part.Worksheet.GetFirstChild<MergeCells>()!.Append(new MergeCell { Reference = "F15:G15" });
        });
        var original = source.ToArray(); MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, FullNestedData(3)));
        Assert.Null(result); Assert.Equal(original, source);
        Assert.Equal(scenario == "unsupported-formula" ? DiagnosticCodes.XlsxUnsafeFormula : DiagnosticCodes.XlsxUnsupportedFeature, error.Code);
        Assert.Contains("Report", error.Errors[0].Location);
        Assert.Contains(scenario switch
        {
            "rich-field" => "Rich-text field substitution", "pane" => "pane", "break" => "manual page breaks",
            "unsupported-formula" => "INDIRECT/OFFSET", _ => "partial merge"
        }, error.Message);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void Activation_Eligibility_is_worksheet_local_and_a_later_ineligible_sheet_returns_no_workbook(bool ineligible)
    {
        var first = ActivationTemplate("depth", nested: false);
        var second = ineligible ? ActivationTemplate("depth") : Create("$l(10).Items", "$l(20).Items", "$l(10).Items");
        var source = _fixture.Create((workbook, part) =>
        {
            using var firstPackage = SpreadsheetDocument.Open(new MemoryStream(first), false);
            using var secondPackage = SpreadsheetDocument.Open(new MemoryStream(second), false);
            part.Worksheet = (Worksheet)firstPackage.WorkbookPart!.WorksheetParts.Single().Worksheet!.CloneNode(true);
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = firstPackage.WorkbookPart.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text;
            var other = workbook.AddNewPart<WorksheetPart>();
            other.Worksheet = (Worksheet)secondPackage.WorkbookPart!.WorksheetParts.Single().Worksheet!.CloneNode(true);
            workbook.Workbook.Sheets!.Append(new Sheet { Name = "Other", SheetId = 2, Id = workbook.GetIdOfPart(other) });
            workbook.Workbook.DefinedNames.Append(new DefinedName(secondPackage.WorkbookPart.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text.Replace("Report", "Other", StringComparison.Ordinal))
                { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
        });
        var original = source.ToArray();
        if (ineligible)
        {
            MergeResult? result = null;
            var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, ActivationData(false)));
            Assert.Null(result); Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, error.Code);
            Assert.Contains("Other", Assert.Single(error.Errors).Location);
        }
        else
        {
            var result = _fixture.Merge(source, ActivationData(false));
            using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
            Assert.Equal(2, package.WorkbookPart!.WorksheetParts.Count());
            Assert.Empty(new OpenXmlValidator().Validate(package));
            Assert.All(package.WorkbookPart.WorksheetParts, part => Assert.NotEmpty(part.Worksheet!.Descendants<Cell>()));
        }
        Assert.Equal(original, source);
    }

    [Theory]
    [InlineData(MissingValueBehavior.EmptyString)]
    [InlineData(MissingValueBehavior.Warning)]
    [InlineData(MissingValueBehavior.Error)]
    public void Activation_Public_nested_ancestor_policy_retains_warning_or_strict_error(MissingValueBehavior behavior)
    {
        var source = ProjectionTemplate([10, 40, 10]); var original = source.ToArray();
        const string json = "{\"Items\":[{\"$level\":40,\"Name\":\"orphan\"}]}";
        var options = new MergeOptions { MissingValues = behavior };
        if (behavior == MissingValueBehavior.Error)
        {
            MergeResult? result = null;
            var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, json, options));
            Assert.Null(result); Assert.Equal(DiagnosticCodes.GenerationInvalidData, error.Code);
            Assert.Contains("data ancestor", error.Errors[0].Message);
            Assert.Equal("$.Items[0]", error.Errors[0].DataPath);
        }
        else
        {
            var result = _fixture.Merge(source, json, options);
            Assert.Equal("orphan/", _fixture.ReadCellText(result.Document, "A1"));
            var warning = Assert.Single(result.Warnings);
            Assert.Equal(DiagnosticCodes.XlsxMissingValue, warning.Code); Assert.Equal("$.Items[0]", warning.DataPath);
            Assert.Contains("data ancestor", warning.Message);
        }
        Assert.Equal(original, source);
    }

    private string ActivationData(bool empty) => JsonSerializer.Serialize(new
    {
        Bids = new[] { new { Name = "bid" } }, Criteria = new[] { new { Name = "criterion" } }, Third = new[] { new { Name = "third" } },
        Items = empty ? [] : new[] { new Dictionary<string, object> { ["$level"] = 10, ["$path"] = "p", ["Name"] = "parent" },
            new Dictionary<string, object> { ["$level"] = 20, ["$path"] = "c", ["Name"] = "child" } }
    });

    private byte[] ActivationTemplate(string shape, bool nested = true) => _fixture.Create((workbook, part) =>
    {
        var rows = new SheetData(
            new Row(_fixture.Text("D1", "$h(1)"), _fixture.Text("E1", "$h(1)"), _fixture.Text("F1", "$h(1)")) { RowIndex = 1 },
            shape == "children"
                ? new Row(_fixture.Text("D2", "$h(2)"), _fixture.Text("F2", "$h(2)")) { RowIndex = 2 }
                : new Row(_fixture.Text("E2", "$h(2)"), _fixture.Text("F2", "$h(2)")) { RowIndex = 2 });
        if (shape == "depth") rows.Append(new Row(_fixture.Text("F3", "$h(3)")) { RowIndex = 3 });
        rows.Append(new Row(_fixture.Text("A7", "$h(1).Bids"), _fixture.Text("D7", "bid")) { RowIndex = 7 },
            new Row(_fixture.Text("A8", shape == "header-rows" ? "$h(1).Bids" : "$h(2).Criteria")) { RowIndex = 8 });
        if (shape == "depth" || shape == "header-rows") rows.Append(new Row(_fixture.Text("A9", shape == "depth" ? "$h(3).Third" : "$h(2).Criteria")) { RowIndex = 9 });
        rows.Append(new Row(_fixture.Text("A10", "$l(10).Items"), _fixture.Text("B10", "{Name}")) { RowIndex = 10 },
            new Row(_fixture.Text("A11", nested ? "$l(20).Items" : "$l(10).Items"), _fixture.Text("B11", "{Name}")) { RowIndex = 11 },
            new Row(_fixture.Text("A12", "$l(10).Items"), _fixture.Text("B12", "footer")) { RowIndex = 12 });
        part.Worksheet = new Worksheet(rows);
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$7:$G$12";
    });
}
