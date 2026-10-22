using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012BindingTests
{
    [Fact]
    public void Hierarchy_Uses_numeric_levels_and_first_ancestor_records_without_sorting_paths()
    {
        using var services = _fixture.Services();
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var template = HierarchyTemplate(horizontal: true);
        using var data = HierarchyData(2, 1);
        var context = new XlsxOperationContext(new MergeOptions());
        var bound = planner.BindSheet(template, data.RootElement, context);
        var rows = bound.Rows[0];
        Assert.Equal(new[] { "z/late", "unrelated", "standalone", "z/late/deeper", "a/early", "z/late/fake", "a/early/deep", "z/late/looks-child" }, rows.Select(row => row.Key));
        Assert.Equal(new[] { 0, 2, 4, 6, 8, 10, 12, 14 }, rows.Select(row => row.FirstIndex));
        Assert.Equal(new int?[] { null, 0, 2, 0, null, 8, 10, null }, rows.Select(row => row.Parent?.FirstIndex));
        Assert.Equal(new[] { "z/late", "a/early", "z/late/looks-child" }, rows.Where(row => row.Parent is null).Select(row => row.Key));
        Assert.Equal(new[] { 2, 6 }, rows[0].ChildFirstIndices);
        Assert.Equal(new[] { 4 }, rows[1].ChildFirstIndices);
        Assert.Equal(new[] { 10 }, rows[4].ChildFirstIndices);
        Assert.Equal(new[] { 12 }, rows[5].ChildFirstIndices);
        Assert.All(new[] { rows[2], rows[3], rows[6], rows[7] }, row => Assert.Empty(row.ChildFirstIndices));
        Assert.Same(rows[0], rows[1].Parent); Assert.Same(rows[0], rows[3].Parent);
        Assert.Same(rows[1], rows[2].Parent); Assert.Same(rows[4], rows[5].Parent);

        var scope = planner.RowScope(data.RootElement, rows[2], "$.Items", 1);
        Assert.Equal(new[] { "G-second", "CA-first", "P-first" },
            scope.Values.Select(value => value.GetProperty("Name").GetString()));
        Assert.Equal(new[] { "$.Items[5]", "$.Items[2]", "$.Items[0]" }, scope.Paths);
        var label = template.Cells.Single(cell => cell.Source.Id.Address == new XlsxAddress(7, 2));
        Assert.Equal("G-second/CA-first/P-first/ROOT", planner.PrepareValue(label, scope, context).Value);
        var rate = template.Cells.Single(cell => cell.Source.Id.Address == new XlsxAddress(7, 3));
        Assert.Equal(new XlsxPreparedValue("n", "10"), planner.PrepareValue(rate, scope, context));

        var resolver = services.GetRequiredService<DataResolver>();
        var language = services.GetRequiredService<TemplateLanguageParser>();
        var path = language.Parse("{../../Rate}", "test").Tokens.Single().Expression!.Path!;
        Assert.Equal("$.Items[0].Rate", resolver.DataPath(path, scope));
        Assert.Equal(10, resolver.Resolve(path, scope).Value.GetInt32());

        var other = planner.RowScope(data.RootElement, rows[6], "$.Items", 1);
        Assert.Equal("GC-second/CC-first/Q-first/ROOT", planner.PrepareValue(label, other, context).Value);
        Assert.Equal(new[] { "$.Items[13]", "$.Items[10]", "$.Items[8]" }, other.Paths);
        Assert.Equal("P-first", planner.RowScope(data.RootElement, rows[0], "$.Items", null).Values[0].GetProperty("Name").GetString());
        Assert.Empty(context.Errors); Assert.Empty(context.Warnings);
    }

    [Theory]
    [InlineData(2, 0)] [InlineData(0, 3)]
    public void Hierarchy_Empty_horizontal_product_keeps_descriptor_ancestors(int companies, int criteria)
    {
        using var services = _fixture.Services();
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var template = HierarchyTemplate(horizontal: true);
        using var data = HierarchyData(companies, criteria);
        var context = new XlsxOperationContext(new MergeOptions());
        var rows = planner.BindSheet(template, data.RootElement, context).Rows[0];
        Assert.Equal(Enumerable.Range(0, 8), rows.Select(row => row.FirstIndex));
        Assert.All(rows, row => Assert.Single(row.Records));
        Assert.Equal(new int?[] { null, 0, 1, 0, null, 4, 5, null }, rows.Select(row => row.Parent?.FirstIndex));
        Assert.Equal(new[] { 1, 3 }, rows[0].ChildFirstIndices);
        var scope = planner.RowScope(data.RootElement, rows[2], "$.Items", null);
        Assert.Equal(new[] { "$.Items[2]", "$.Items[1]", "$.Items[0]" }, scope.Paths);
        var label = template.Cells.Single(cell => cell.Source.Id.Address == new XlsxAddress(7, 2));
        Assert.Equal("G-first/CA-first/P-first/ROOT", planner.PrepareValue(label, scope, context).Value);
        Assert.Empty(context.Errors); Assert.Empty(context.Warnings);
    }

    [Fact]
    public void Hierarchy_Vertical_only_without_paths_uses_numeric_stack_across_siblings()
    {
        using var services = _fixture.Services();
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        using var data = JsonDocument.Parse("""
            {"Name":"ROOT","Items":[
              {"$level":10,"Name":"P"},{"$level":40,"Name":"CA"},
              {"$level":90,"Name":"G"},{"$level":40,"Name":"CB"},
              {"$level":10,"Name":"Q"},{"$level":40,"Name":"CC"}]}
            """);
        var rows = planner.BindSheet(HierarchyTemplate(), data.RootElement, new(new MergeOptions())).Rows[0];
        Assert.All(rows, row => Assert.Null(row.Key));
        Assert.Equal(new[] { "CB", "P" }, planner.RowScope(data.RootElement, rows[3], "$.Items", null)
            .Values.Select(value => value.GetProperty("Name").GetString()));
        Assert.Equal(new[] { "CC", "Q" }, planner.RowScope(data.RootElement, rows[5], "$.Items", null)
            .Values.Select(value => value.GetProperty("Name").GetString()));
    }

    [Theory]
    [InlineData("[]", 0)]
    [InlineData("[{\"$level\":40,\"Name\":\"only\"}]", 1)]
    [InlineData("[{\"$level\":40,\"Name\":\"first\"},{\"$level\":40,\"Name\":\"second\"}]", 2)]
    public void Hierarchy_Empty_single_and_equal_level_inputs_do_not_invent_ancestors(string items, int count)
    {
        using var services = _fixture.Services();
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        using var data = JsonDocument.Parse("{\"Name\":\"ROOT\",\"Items\":" + items + "}");
        var rows = planner.BindSheet(HierarchyTemplate(), data.RootElement, new(new MergeOptions())).Rows[0];
        Assert.Equal(count, rows.Length);
        Assert.All(rows, row => { Assert.Null(row.Parent); Assert.Empty(row.ChildFirstIndices); });
        foreach (var row in rows)
            Assert.Equal(new[] { "$.Items[" + row.FirstIndex + "]" },
                planner.RowScope(data.RootElement, row, "$.Items", null).Paths);
    }

    [Theory]
    [InlineData("{\"$level\":40.5}", "integer $level")]
    [InlineData("{\"$level\":2147483648}", "integer $level")]
    [InlineData("{\"$level\":null}", "integer $level")]
    [InlineData("{\"$level\":40,\"$path\":\"same\"}", "same integer $level")]
    public void Hierarchy_Rejects_invalid_or_inconsistent_levels_with_source_and_data_location(string child, string remedy)
    {
        using var services = _fixture.Services();
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        using var data = JsonDocument.Parse("{\"Items\":[{\"$level\":10,\"$path\":\"same\"}," + child + "]}");
        var error = Assert.Throws<TemplateException>(() => planner.BindSheet(HierarchyTemplate(), data.RootElement, new(new MergeOptions())));
        var issue = Assert.Single(error.Errors);
        Assert.Equal(DiagnosticCodes.GenerationInvalidData, issue.Code);
        Assert.Contains("sheet 'Report', A3", issue.Location);
        Assert.Equal("$.Items[1]", issue.DataPath);
        Assert.Contains(remedy, issue.Message);
    }

    [Fact]
    public void Hierarchy_Rejects_reappearing_groups_instead_of_reparenting_or_sorting_them()
    {
        using var services = _fixture.Services();
        using var data = JsonDocument.Parse("""
            {"Items":[{"$level":10,"$path":"p"},{"$level":40,"$path":"c"},{"$level":10,"$path":"p"}]}
            """);
        var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxTemplatePlanner>()
            .BindSheet(HierarchyTemplate(), data.RootElement, new(new MergeOptions())));
        Assert.Equal(DiagnosticCodes.GenerationInvalidData, error.Code);
        Assert.Equal("$.Items[2]", Assert.Single(error.Errors).DataPath);
        Assert.Contains("closed path must not reappear", error.Message);
    }

    [Theory]
    [InlineData(2, 1)] [InlineData(2, 0)]
    public void Hierarchy_Validates_nested_group_arity_before_exposing_bound_rows(int companies, int criteria)
    {
        using var services = _fixture.Services();
        var arity = criteria == 0 ? 1 : 2;
        using var data = JsonDocument.Parse(JsonSerializer.Serialize(new
        {
            Bids = Enumerable.Range(0, companies).Select(_ => new { Company = "bid" }),
            Criteria = Enumerable.Range(0, criteria).Select(_ => new { Name = "criterion" }),
            Items = Enumerable.Range(0, arity).Select(_ => new Dictionary<string, object> { ["$level"] = 10, ["$path"] = "p" })
                .Concat(Enumerable.Range(0, arity + 1).Select(_ => new Dictionary<string, object> { ["$level"] = 40, ["$path"] = "c" }))
        }));
        var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxTemplatePlanner>()
            .BindSheet(HierarchyTemplate(horizontal: true), data.RootElement, new(new MergeOptions())));
        Assert.Equal(DiagnosticCodes.GenerationInvalidData, error.Code);
        Assert.Equal("$.Items[" + arity + "]", Assert.Single(error.Errors).DataPath);
        Assert.Contains("exactly " + arity, error.Message);
    }

    [Fact]
    public void Hierarchy_Single_level_source_preserves_unrendered_data_ancestor_without_warnings()
    {
        var source = _fixture.Create((_, worksheet) =>
        {
            var row = worksheet.Worksheet!.Descendants<Row>().Single();
            row.PrependChild(_fixture.Text("A3", "$l(40).Items"));
            row.Elements<Cell>().Single(cell => cell.CellReference == "B3").InlineString!.Text!.Text = "{../Name}";
        });
        var result = _fixture.Merge(source, """
            {"Name":"ROOT","Items":[{"$level":10,"Name":"omitted"},{"$level":40,"Name":"child"}]}
            """);
        Assert.Equal("omitted", _fixture.ReadCellText(result.Document, "A1"));
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Hierarchy_Handles_integer_extremes_jumps_and_new_lower_roots_without_arithmetic_or_path_inference()
    {
        using var services = _fixture.Services();
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        using var data = JsonDocument.Parse("""
            {"Items":[{"$level":40,"$path":"a/b"},{"$level":2147483647,"$path":"z"},
              {"$level":-2147483648,"$path":"a/b/c"},{"$level":2147483647,"$path":"a"},
              {"$level":2147483647,"$path":"a/child"}]}
            """);
        var rows = planner.BindSheet(HierarchyTemplate(), data.RootElement, new(new MergeOptions())).Rows[0];
        Assert.Equal(new int?[] { null, 0, null, 2, 2 }, rows.Select(row => row.Parent?.FirstIndex));
        Assert.Equal(new[] { 1 }, rows[0].ChildFirstIndices);
        Assert.Equal(new[] { 3, 4 }, rows[2].ChildFirstIndices);
        Assert.Equal(new[] { "$.Items[4]", "$.Items[2]" }, planner.RowScope(data.RootElement, rows[4], "$.Items", null).Paths);
    }

    [Fact]
    public void Hierarchy_Bindings_are_independent_across_calls_on_the_same_planner()
    {
        using var services = _fixture.Services();
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var template = HierarchyTemplate();
        using var first = HierarchyData(0, 1);
        using var next = JsonDocument.Parse("""
            {"Name":"NEXT ROOT","Items":[{"$level":10,"Name":"NEXT"},{"$level":40,"Name":"next child"}]}
            """);
        var previous = planner.BindSheet(template, first.RootElement, new(new MergeOptions())).Rows[0];
        var current = planner.BindSheet(template, next.RootElement, new(new MergeOptions())).Rows[0];
        Assert.Equal(new[] { "next child", "NEXT" }, planner.RowScope(next.RootElement, current[1], "$.Items", null)
            .Values.Select(value => value.GetProperty("Name").GetString()));
        Assert.Equal(new[] { "G-first", "CA-first", "P-first" }, planner.RowScope(first.RootElement, previous[2], "$.Items", null)
            .Values.Select(value => value.GetProperty("Name").GetString()));
        Assert.Same(previous[0], previous[1].Parent); Assert.Same(current[0], current[1].Parent);
        Assert.NotSame(previous[0], current[1].Parent);
        Assert.Equal(new[] { 1, 3 }, previous[0].ChildFirstIndices);
        Assert.Equal(new[] { 1 }, current[0].ChildFirstIndices);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void Hierarchy_Missing_and_null_ancestor_fields_never_fall_back_to_another_tuple_or_root(bool explicitNull)
    {
        using var services = _fixture.Services();
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var template = HierarchyTemplate(horizontal: true);
        var first = explicitNull ? "{\"$level\":10,\"$path\":\"p\",\"Rate\":null}" : "{\"$level\":10,\"$path\":\"p\"}";
        using var data = JsonDocument.Parse("""
            {"Rate":777,"Bids":[{},{}],"Criteria":[{}],"Items":[
            """ + first + """
            ,{"$level":10,"$path":"p","Rate":999},
            {"$level":40,"$path":"c"},{"$level":40,"$path":"c"},
            {"$level":90,"$path":"g"},{"$level":90,"$path":"g"}]}
            """);
        var context = new XlsxOperationContext(new MergeOptions { MissingValues = MissingValueBehavior.Warning });
        var rows = planner.BindSheet(template, data.RootElement, context).Rows[0];
        var scope = planner.RowScope(data.RootElement, rows[2], "$.Items", 1);
        var rate = template.Cells.Single(cell => cell.Source.Id.Address == new XlsxAddress(7, 3));
        var path = rate.Tokens.Single().Expression!.Path!;
        var resolver = services.GetRequiredService<DataResolver>();
        Assert.Equal(explicitNull ? ResolutionState.Null : ResolutionState.Missing, resolver.Resolve(path, scope).State);
        Assert.Equal("$.Items[0].Rate", resolver.DataPath(path, scope));
        Assert.Equal(new XlsxPreparedValue(null, null), planner.PrepareValue(rate, scope, context));
        if (explicitNull) Assert.Empty(context.Warnings);
        else
        {
            var warning = Assert.Single(context.Warnings);
            Assert.Equal(DiagnosticCodes.XlsxMissingValue, warning.Code);
            Assert.Equal("$.Items[0].Rate", warning.DataPath);
            Assert.Contains("C7", warning.Location);
        }
        Assert.Empty(context.Errors);
    }

    private XlsxSheetTemplate HierarchyTemplate(bool horizontal = false)
    {
        var bytes = _fixture.Create((workbook, worksheet) =>
        {
            var data = worksheet.Worksheet!.GetFirstChild<SheetData>()!;
            data.RemoveAllChildren();
            if (horizontal)
                data.Append(new Row(_fixture.Text("D1", "$h(1)"), _fixture.Text("E1", "$h(1)")) { RowIndex = 1 },
                    new Row(_fixture.Text("E2", "$h(2)")) { RowIndex = 2 },
                    new Row(_fixture.Text("A3", "$h(1).Bids"), _fixture.Text("D3", "{Company}")) { RowIndex = 3 },
                    new Row(_fixture.Text("A4", "$h(2).Criteria"), _fixture.Text("E4", "{Name}")) { RowIndex = 4 });
            var start = horizontal ? 5 : 3;
            var levels = new[] { 10, 40, 90, 40, 10 };
            for (var index = 0; index < levels.Length; index++)
            {
                var row = start + index;
                var element = new Row(_fixture.Text("A" + row, "$l(" + levels[index] + ").Items"),
                    _fixture.Text("B" + row, index == 2 ? "{Name}/{../Name}/{../../Name}/{$.Name}" : "{Name}")) { RowIndex = (uint)row };
                if (index == 2)
                {
                    element.Append(_fixture.Text("C" + row, "{../../Rate}"));
                    if (horizontal) element.Append(_fixture.Text("E" + row, "{Price}"));
                }
                data.Append(element);
            }
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$" +
                (horizontal ? "F" : "D") + (start + 4);
        });
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        return services.GetRequiredService<XlsxTemplatePlanner>().Plan(
            services.GetRequiredService<XlsxPackageReader>().Read(package, context), context).Sheets.Single();
    }

    private JsonDocument HierarchyData(int companies, int criteria)
    {
        var arity = Math.Max(1, companies * criteria);
        var groups = new[] { (10, "z/late", "P"), (40, "unrelated", "CA"), (90, "standalone", "G"),
            (40, "z/late/deeper", "CB"), (10, "a/early", "Q"), (40, "z/late/fake", "CC"),
            (90, "a/early/deep", "GC"), (10, "z/late/looks-child", "R") };
        return JsonDocument.Parse(JsonSerializer.Serialize(new
        {
            Name = "ROOT",
            Bids = Enumerable.Range(0, companies).Select(_ => new { Company = "bid must not become an ancestor" }),
            Criteria = Enumerable.Range(0, criteria).Select(_ => new { Name = "criterion must not become an ancestor" }),
            Items = groups.SelectMany((group, index) => Enumerable.Range(0, arity).Select(tuple =>
                new Dictionary<string, object> { ["$level"] = group.Item1, ["$path"] = group.Item2,
                    ["Name"] = group.Item3 + (tuple == 0 ? "-first" : "-second"),
                    ["Rate"] = tuple == 0 ? (index + 1) * 10 : 999, ["Price"] = index * 100 + tuple }))
        }));
    }
}
