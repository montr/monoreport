using System.IO;
using System.Linq;
using System.Numerics;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport;
using MonoReport.Core.Diagnostics;
using Xunit;

namespace MonoReport.Tests.Xlsx;

public sealed class XlsxSampleJsonGeneratorTests
{
    private readonly XlsxFixture _fixture = new();

    [Fact]
    public void Vertical_sample_uses_authored_levels_and_renders_in_record_order()
    {
        using var services = _fixture.Services();
        var template = _fixture.CreateRepeated();
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(2)});

        var items = sample.Data.GetProperty("Items");
        Assert.Equal(2, items.GetArrayLength());
        Assert.All(items.EnumerateArray(), item => Assert.Equal(10, item.GetProperty("$level").GetInt32()));
        Assert.All(items.EnumerateArray(), item => Assert.False(item.TryGetProperty("$path", out _)));
        var result = Render(services, template, sample.Data);
        Assert.Equal("Sample Name 1", _fixture.ReadCellText(result, "A1"));
        Assert.Equal("Sample Name 2", _fixture.ReadCellText(result, "A2"));
    }

    [Fact]
    public void Two_axis_sample_creates_complete_contiguous_tuple_groups_and_renders_leaf_order()
    {
        using var services = _fixture.Services();
        var template = new XlsxTwoAxisFixture().Create();
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(2)});

        Assert.Equal(2, sample.Data.GetProperty("Bids").GetArrayLength());
        Assert.Equal(2, sample.Data.GetProperty("Criteria").GetArrayLength());
        var records = sample.Data.GetProperty("Items").EnumerateArray().ToArray();
        Assert.Equal(8, records.Length);
        Assert.Equal(new[] {10, 10, 10, 10, 10, 10, 10, 10},
            records.Select(item => item.GetProperty("$level").GetInt32()));
        Assert.All(records.Chunk(4), group => Assert.Single(group.Select(item => item.GetProperty("$path").GetString()).Distinct()));
        Assert.Equal(records[0].GetProperty("$path").GetString(), records[3].GetProperty("$path").GetString());
        Assert.NotEqual(records[0].GetProperty("$path").GetString(), records[4].GetProperty("$path").GetString());

        var rendered = Render(services, template, sample.Data);
        using var package = SpreadsheetDocument.Open(new MemoryStream(rendered), false);
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        var prices = worksheet.Descendants<Cell>().Where(cell => cell.InlineString?.InnerText.StartsWith(
                "Sample Price ", System.StringComparison.Ordinal) == true)
            .Select(cell => cell.InlineString!.InnerText).ToArray();
        Assert.Equal(Enumerable.Range(1, 8).Select(index => "Sample Price " + index), prices);
    }

    [Fact]
    public void Footer_first_nested_layout_still_generates_authored_parent_before_child_levels()
    {
        using var services = _fixture.Services();
        var template = _fixture.Create((workbook, worksheet) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$D$5";
            worksheet.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A3", "$l(90).Items"), _fixture.Text("B3", "{Name}")) {RowIndex = 3},
                new Row(_fixture.Text("A4", "$l(40).Items"), _fixture.Text("B4", "{Name}")) {RowIndex = 4},
                new Row(_fixture.Text("A5", "$l(10).Items"), _fixture.Text("B5", "{Name}")) {RowIndex = 5}));
        });
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(1)});

        Assert.Equal(new[] {10, 40, 90}, sample.Data.GetProperty("Items").EnumerateArray()
            .Select(item => item.GetProperty("$level").GetInt32()));
        var rendered = Render(services, template, sample.Data);
        Assert.Equal(new[] {"Sample Name 3", "Sample Name 2", "Sample Name 1"},
            Enumerable.Range(1, 3).Select(row => _fixture.ReadCellText(rendered, "A" + row)));
    }

    [Fact]
    public void Nested_hierarchy_and_two_header_axes_keep_shared_record_groups_and_parent_values()
    {
        using var services = _fixture.Services();
        var template = new XlsxNestedFixture().Create(false);
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(2)});

        var records = sample.Data.GetProperty("Items").EnumerateArray().ToArray();
        Assert.Equal(24, records.Length);
        Assert.Equal(new[] {10, 20, 20, 10, 20, 20}, records.Chunk(4)
            .Select(group => group[0].GetProperty("$level").GetInt32()));
        Assert.All(records.Chunk(4), group =>
            Assert.Single(group.Select(item => item.GetProperty("$path").GetString()).Distinct()));
        Assert.Equal(6, records.Chunk(4).Select(group => group[0].GetProperty("$path").GetString()).Distinct().Count());

        var rendered = Render(services, template, sample.Data);
        using var package = SpreadsheetDocument.Open(new MemoryStream(rendered), false);
        var cells = package.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>().ToArray();
        Assert.Contains(cells, cell => cell.InlineString?.InnerText.Contains("Sample Name", System.StringComparison.Ordinal) == true);
        Assert.Contains(cells, cell => cell.InlineString?.InnerText.Contains("Sample Company", System.StringComparison.Ordinal) == true);
    }

    [Fact]
    public void Unknown_header_shapes_stay_empty_and_intersections_keep_one_descriptor()
    {
        using var services = _fixture.Services();
        var template = new XlsxTwoAxisFixture().Create((_, worksheet) =>
        {
            worksheet.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference == "D7")
                .InlineString!.Text!.Text = "Company";
            worksheet.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "E8")
                .InlineString!.Text!.Text = "Criterion";
        });
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(2)});

        Assert.Empty(sample.Data.GetProperty("Bids").EnumerateArray());
        Assert.Empty(sample.Data.GetProperty("Criteria").EnumerateArray());
        var records = sample.Data.GetProperty("Items").EnumerateArray().ToArray();
        Assert.Equal(2, records.Length);
        Assert.All(records, item => Assert.True(item.TryGetProperty("$path", out _)));
        Assert.Equal(2, sample.Warnings.Count);
        Assert.All(sample.Warnings, warning => Assert.Equal("sample.unresolved_shape", warning.Code));
        _ = Render(services, template, sample.Data);
    }

    [Fact]
    public void Compatible_collection_uses_across_worksheets_merge_into_one_array()
    {
        using var services = _fixture.Services();
        var template = SharedCollectionTemplate(false);
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(2)});

        var items = sample.Data.GetProperty("Items");
        Assert.Equal(2, items.GetArrayLength());
        Assert.Equal("Sample Name 1", items[0].GetProperty("Name").GetString());
        Assert.Equal("Sample Code 1", items[0].GetProperty("Code").GetString());
        var rendered = Render(services, template, sample.Data);
        using var package = SpreadsheetDocument.Open(new MemoryStream(rendered), false);
        var workbook = package.WorkbookPart!;
        Assert.Equal(2, workbook.WorksheetParts.Count());
        Assert.Contains(workbook.WorksheetParts, part => part.Worksheet!.Descendants<Cell>()
            .Any(cell => cell.InlineString?.InnerText == "Sample Code 1"));
    }

    [Fact]
    public void Shared_hierarchy_merges_compatible_level_prefixes()
    {
        using var services = _fixture.Services();
        var template = CompatibleSubsetSharedCollectionTemplate();
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(2)});

        var items = sample.Data.GetProperty("Items").EnumerateArray().ToArray();
        Assert.Equal(new[] {10, 20, 20, 10, 20, 20}, items.Select(item => item.GetProperty("$level").GetInt32()));
        Assert.Equal("Sample Name 1", items[0].GetProperty("Name").GetString());
        Assert.Equal("Sample Code 1", items[0].GetProperty("Code").GetString());
        Assert.Equal("Sample Value 2", items[1].GetProperty("Value").GetString());
    }

    [Fact]
    public void Horizontal_outer_header_parent_path_resolves_to_root()
    {
        using var services = _fixture.Services();
        var template = new XlsxTwoAxisFixture().Create((_, worksheet) =>
            worksheet.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference == "D7")
                .InlineString!.Text!.Text = "{../Title}");
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input);

        Assert.Equal("Sample Title 1", sample.Data.GetProperty("Title").GetString());
        Assert.DoesNotContain(sample.Warnings, warning => warning.Code == DiagnosticCodes.SampleUnavailableAncestor);
    }

    [Fact]
    public void Ancestor_leaf_evidence_is_stored_only_on_its_first_tuple()
    {
        using var services = _fixture.Services();
        var template = new XlsxNestedFixture().Create(false, (_, worksheet) =>
            worksheet.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference == "E11")
                .InlineString!.Text!.Text = "{../ParentCode}");
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(2)});

        var records = sample.Data.GetProperty("Items").EnumerateArray().ToArray();
        Assert.True(records[0].TryGetProperty("ParentCode", out _));
        Assert.All(records.Skip(1).Take(3), record => Assert.False(record.TryGetProperty("ParentCode", out _)));
        Assert.True(records[12].TryGetProperty("ParentCode", out _));
        Assert.All(records.Skip(13).Take(3), record => Assert.False(record.TryGetProperty("ParentCode", out _)));
    }

    [Fact]
    public void Semantic_conflicts_follow_worksheet_traversal_order()
    {
        using var services = _fixture.Services();
        var template = SharedSemanticConflictTemplate();
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input);

        Assert.Contains(sample.Warnings, warning => warning.Code == DiagnosticCodes.SampleSemanticConflict);
        Assert.Equal(42.5m, sample.Data.GetProperty("Items")[0].GetProperty("Value").GetDecimal());
    }

    [Fact]
    public void Conflicts_in_one_cell_keep_one_warning_per_property()
    {
        using var services = _fixture.Services();
        var template = _fixture.Create((_, worksheet) =>
            worksheet.Worksheet!.Descendants<Text>().Single().Text =
                "{First number \"0.0\"}/{First date \"yyyy-MM-dd\"}/{Second number \"0.0\"}/{Second date \"yyyy-MM-dd\"}");
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input);

        var warnings = sample.Warnings.Where(warning => warning.Code == DiagnosticCodes.SampleSemanticConflict).ToArray();
        Assert.Equal(2, warnings.Length);
        Assert.Equal(new[] {"First", "Second"}, warnings.Select(warning => warning.DataPath));
    }

    [Fact]
    public void Incompatible_shared_hierarchy_demands_fail_without_a_sample()
    {
        using var services = _fixture.Services();
        var template = SharedCollectionTemplate(true);
        using var input = new MemoryStream(template);

        var error = Assert.Throws<MonoReport.Core.Templates.TemplateException>(() =>
            services.GetRequiredService<SampleJsonGenerator>().Generate(input));

        Assert.Equal("sample.incompatible_demand", error.Code);
        Assert.NotNull(error.Location);
    }

    [Fact]
    public void Huge_nested_and_horizontal_counts_fail_during_preflight()
    {
        using var services = _fixture.Services();
        var template = new XlsxNestedFixture().Create(false);
        using var input = new MemoryStream(template);

        var error = Assert.Throws<MonoReport.Core.Templates.TemplateException>(() =>
            services.GetRequiredService<SampleJsonGenerator>().Generate(input,
                new SampleJsonOptions {ItemsPerCollection = BigInteger.Pow(10, 100_000)}));

        Assert.Equal(DiagnosticCodes.SampleElementLimit, error.Code);
        Assert.False(error.LimitReached);
    }

    [Fact]
    public void Count_above_aggregate_element_limit_fails_during_preflight()
    {
        using var services = _fixture.Services();
        var template = new XlsxNestedFixture().Create(false);
        using var input = new MemoryStream(template);

        var error = Assert.Throws<MonoReport.Core.Templates.TemplateException>(() =>
            services.GetRequiredService<SampleJsonGenerator>().Generate(input,
                new SampleJsonOptions {ItemsPerCollection = new BigInteger(SampleJsonLimits.PreflightSentinel)}));

        Assert.Equal(DiagnosticCodes.SampleElementLimit, error.Code);
        Assert.False(error.LimitReached);
    }

    [Fact]
    public void Same_canonical_property_can_use_different_semantic_evidence_at_distinct_levels()
    {
        using var services = _fixture.Services();
        var template = _fixture.Create((workbook, worksheet) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$B$5";
            worksheet.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A3", "$l(10).Items"), _fixture.Text("B3", "{Value number \"0.0\"}")) {RowIndex = 3},
                new Row(_fixture.Text("A4", "$l(20).Items"), _fixture.Text("B4", "{Value date \"yyyy-MM-dd\"}")) {RowIndex = 4},
                new Row(_fixture.Text("A5", "$l(10).Items"), _fixture.Text("B5", "{Value}")) {RowIndex = 5}));
        });
        using var input = new MemoryStream(template);

        var sample = services.GetRequiredService<SampleJsonGenerator>().Generate(input,
            new SampleJsonOptions {ItemsPerCollection = new BigInteger(1)});

        var items = sample.Data.GetProperty("Items");
        Assert.Equal(10, items[0].GetProperty("$level").GetInt32());
        Assert.Equal(42.5m, items[0].GetProperty("Value").GetDecimal());
        Assert.Equal(20, items[1].GetProperty("$level").GetInt32());
        Assert.Equal("2024-01-16", items[1].GetProperty("Value").GetString());
        Assert.Empty(sample.Warnings);
    }

    private static byte[] Render(ServiceProvider services, byte[] template, JsonElement sample)
    {
        using var data = JsonDocument.Parse(sample.GetRawText());
        using var input = new MemoryStream(template);
        return services.GetRequiredService<DocumentGenerator>().Generate(input, data).Document;
    }

    private byte[] SharedCollectionTemplate(bool incompatible)
    {
        return _fixture.Create((workbook, worksheet) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$B$3";
            worksheet.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A3", "$l(10).Items"), _fixture.Text("B3", "{Name}")) {RowIndex = 3}));
            var second = workbook.AddNewPart<WorksheetPart>();
            second.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A3", "$l(" + (incompatible ? "20" : "10") + ").Items"),
                    _fixture.Text("B3", "{Code}")) {RowIndex = 3}));
            workbook.Workbook.Sheets!.Append(new Sheet
                {Name = "Summary", SheetId = 2, Id = workbook.GetIdOfPart(second)});
            workbook.Workbook.DefinedNames.Append(new DefinedName("'Summary'!$B$3:$B$3")
                {Name = "_xlnm.Print_Area", LocalSheetId = 1});
        });
    }

    private byte[] CompatibleSubsetSharedCollectionTemplate()
    {
        return _fixture.Create((workbook, worksheet) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$B$4";
            worksheet.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A3", "$l(10).Items"), _fixture.Text("B3", "{Name}")) {RowIndex = 3},
                new Row(_fixture.Text("A4", "$l(20).Items"), _fixture.Text("B4", "{Value}")) {RowIndex = 4}));
            var second = workbook.AddNewPart<WorksheetPart>();
            second.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A3", "$l(10).Items"), _fixture.Text("B3", "{Code}")) {RowIndex = 3}));
            workbook.Workbook.Sheets!.Append(new Sheet
                {Name = "Summary", SheetId = 2, Id = workbook.GetIdOfPart(second)});
            workbook.Workbook.DefinedNames.Append(new DefinedName("'Summary'!$B$3:$B$3")
                {Name = "_xlnm.Print_Area", LocalSheetId = 1});
        });
    }

    private byte[] SharedSemanticConflictTemplate()
    {
        return _fixture.Create((workbook, worksheet) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$C$3";
            worksheet.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A3", "$l(10).Items"),
                    _fixture.Text("C3", "{Value number \"0.0\"}")) {RowIndex = 3}));
            var second = workbook.AddNewPart<WorksheetPart>();
            second.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A3", "$l(10).Items"),
                    _fixture.Text("B3", "{Value date \"yyyy-MM-dd\"}")) {RowIndex = 3}));
            workbook.Workbook.Sheets!.Append(new Sheet
                {Name = "Summary", SheetId = 2, Id = workbook.GetIdOfPart(second)});
            workbook.Workbook.DefinedNames.Append(new DefinedName("'Summary'!$B$3:$B$3")
                {Name = "_xlnm.Print_Area", LocalSheetId = 1});
        });
    }
}
