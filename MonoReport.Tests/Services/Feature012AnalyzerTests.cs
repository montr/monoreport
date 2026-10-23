using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed class Feature012AnalyzerTests
{
    private readonly XlsxFixture _fixture = new();

    [Fact]
    public void Analyzer_Prepared_marker_scopes_keep_axes_flat_parents_separate_and_occurrences_ordered()
    {
        var source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures/feature012/analyzer-scopes.template.xlsx"));
        using var package = SpreadsheetDocument.Open(new MemoryStream(source), false);
        var context = new XlsxOperationContext(new MergeOptions());
        using var services = _fixture.Services();
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Prepare(new XlsxPackageReader().Read(package, context), context);
        var builder = new TemplateRequirementBuilder();
        foreach (var dependency in plan.Dependencies) builder.Add(dependency);
        Assert.Empty(builder.Diagnostics);
        Assert.Equal(new[] { "Bids", "Criteria", "Items", "Title" }, builder.Requirements.Select(item => item.Name));
        Assert.Equal(new[] { "array", "array", "array", "value" }, builder.Requirements.Select(item => item.Type));
        Assert.Equal(new[] { "Company" }, builder.Requirements[0].Fields!.Select(item => item.Name));
        Assert.Equal(new[] { "Name" }, builder.Requirements[1].Fields!.Select(item => item.Name));
        Assert.Equal(new[] { "Name", "Price" }, builder.Requirements[2].Fields!.Select(item => item.Name));
        Assert.Equal(new[]
        {
            "$.Bids", "$.Bids[].Company", "$.Criteria", "$.Criteria[].Name", "$.Bids[].Company",
            "$.Items", "$.Items[].Name", "$.Title", "$.Items[].Name", "$.Items", "$.Items[].Name",
            "$.Items[].Price", "$.Items", "$.Items[].Name", "$.Title"
        }, builder.Occurrences.Select(item => item.Path));
        Assert.Equal(new[] { "$h(1).Bids", "$header(2).Criteria", "$l(10).Items", "$level(20).Items", "$l(10).Items" },
            builder.Occurrences.Where(item => item.Kind == "for").Select(item => item.Expression));
        Assert.Equal(3, plan.Sheets[0].Controls.Count(marker => marker.Kind == XlsxControlKind.Geometry));
        Assert.All(builder.Occurrences, item => Assert.Equal("/xl/worksheets/sheet1.xml", item.Part));
        Assert.Contains(builder.Occurrences, item => item.Expression == "{../Company}" && item.Path == "$.Bids[].Company" && item.Location!.Contains("C4"));
        Assert.Contains(builder.Occurrences, item => item.Expression == "{../Name}" && item.Path == "$.Items[].Name" && item.Location!.Contains("B6"));

        using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.True(analysis.Valid);
        Assert.Empty(analysis.Errors); Assert.Empty(analysis.Warnings);
        Assert.Equal(JsonSerializer.Serialize(builder.Requirements), JsonSerializer.Serialize(analysis.Requirements));
        Assert.Equal(JsonSerializer.Serialize(builder.Occurrences), JsonSerializer.Serialize(analysis.Occurrences));
        Assert.Equal(source, input.ToArray());
        var json = JsonSerializer.SerializeToElement(analysis, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
        Assert.Equal(new[] { "errors", "occurrences", "requirements", "valid", "warnings" }, json.EnumerateObject().Select(item => item.Name).Order());
        // Analysis is data-independent; public nested output is still separately gated.
        Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature,
            Assert.Throws<TemplateException>(() => _fixture.Merge(source, "{\"Items\":[]}")).Code);
    }

    [Fact]
    public void Analyzer_Static_public_schema_preserves_duplicates_root_paths_and_sheet_order()
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            worksheet.Worksheet!.Descendants<Text>().Single().Text = "{Name} {$.Title} {Name}";
            var other = workbook.AddNewPart<WorksheetPart>();
            other.Worksheet = new Worksheet(new SheetData(new Row(_fixture.Text("A1", "{Other}")) { RowIndex = 1 }));
            workbook.Workbook!.Sheets!.Append(new Sheet { Name = "A second", SheetId = 2, Id = workbook.GetIdOfPart(other) });
            workbook.Workbook.DefinedNames!.Append(new DefinedName("'A second'!$A$1") { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
        });
        using var input = new MemoryStream(source);
        using var services = _fixture.Services();
        var result = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.True(result.Valid);
        Assert.Equal(new[] { "Name", "Title", "Other" }, result.Requirements.Select(item => item.Name));
        Assert.Equal(new[] { "$.Name", "$.Title", "$.Name", "$.Other" }, result.Occurrences.Select(item => item.Path));
        var json = JsonSerializer.SerializeToElement(result, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
        Assert.Equal(new[] { "errors", "occurrences", "requirements", "valid", "warnings" }, json.EnumerateObject().Select(item => item.Name).Order());
        Assert.Contains("A second", result.Occurrences[^1].Location);
    }

    [Theory]
    [InlineData("$HEADER(1).Bids", "A3", DiagnosticCodes.TemplateInvalidStructure)]
    [InlineData("$h(1)", "B1", DiagnosticCodes.TemplateInvalidStructure)]
    [InlineData("$l(x).Items", "A3", DiagnosticCodes.TemplateInvalidStructure)]
    [InlineData("$l(10).Items", "C3", DiagnosticCodes.TemplateInvalidStructure)]
    public void Analyzer_Unsupported_or_malformed_controls_never_disappear_silently(string marker, string address, string code)
    {
        var source = _fixture.Create((_, worksheet) =>
        {
            var data = worksheet.Worksheet!.GetFirstChild<SheetData>()!;
            var rowIndex = address.EndsWith('1') ? 1U : 3U;
            var row = data.Elements<Row>().FirstOrDefault(item => item.RowIndex!.Value == rowIndex);
            if (row is null) { row = new Row { RowIndex = rowIndex }; data.PrependChild(row); }
            if (address.StartsWith('A')) row.PrependChild(_fixture.Text(address, marker));
            else row.Append(_fixture.Text(address, marker));
        });
        using var input = new MemoryStream(source);
        using var services = _fixture.Services();
        var result = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(result.Valid);
        Assert.Contains(result.Errors, issue => issue.Code == code && issue.Location!.Contains(address, StringComparison.Ordinal));
        Assert.Equal(source, input.ToArray());
    }
}
