using System.Collections.Immutable;
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

public sealed partial class Feature012BindingTests
{
    [Fact]
    public void DataAuthority_Flat_source_keeps_multiple_unrendered_ancestors_and_duplicate_people_in_order()
    {
        var source = FlatDataTemplate("{Name}/{../Name}/{../../Name}/{../../../Name}/{../../../../Name}/{$.Name}");
        const string json = """
            {"Name":"ROOT","Items":[{"$level":5,"Name":"World"},{"$level":10,"Name":"Company A"},
              {"$level":20,"Name":"Department X"},{"$level":30,"Name":"Team T"},
              {"$level":40,"Name":"Ivan"},{"$level":40,"Name":"Anna"},
              {"$level":20,"Name":"Department Y"},{"$level":30,"Name":"Team U"},{"$level":40,"Name":"Ivan"}]}
            """;
        using var services = _fixture.Services(); using var data = JsonDocument.Parse(json);
        using var package = SpreadsheetDocument.Open(new MemoryStream(source), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var plan = planner.Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var sheet = Assert.Single(plan.Sheets);
        Assert.Equal(Enumerable.Repeat("$.Items[].Name", 5).Append("$.Name"),
            plan.Occurrences.Where(item => item.Kind == "field").Select(item => item.Path));
        var bound = planner.BindSheet(sheet, data.RootElement, context);
        var blocks = Assert.Single(planner.ScheduleSheet(sheet, bound, context).Regions).Blocks;
        Assert.Equal(new[] { "0,1,2,3,4", "0,1,2,3,5", "0,1,6,7,8" },
            blocks.Select(block => string.Join(',', block.Owner.LogicalRows)));
        Assert.Same(bound.Rows[0][3], bound.Rows[0][4].Parent);
        Assert.Same(bound.Rows[0][7], bound.Rows[0][8].Parent);
        var result = _fixture.Merge(source, json);
        Assert.Equal("Ivan/Team T/Department X/Company A/World/ROOT", _fixture.ReadCellText(result.Document, "A1"));
        Assert.Equal("Anna/Team T/Department X/Company A/World/ROOT", _fixture.ReadCellText(result.Document, "A2"));
        Assert.Equal("Ivan/Team U/Department Y/Company A/World/ROOT", _fixture.ReadCellText(result.Document, "A3"));
        Assert.Empty(result.Warnings); Assert.Empty(context.Warnings); Assert.Empty(context.Errors);
        using var input = new MemoryStream(source);
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
    }

    [Theory]
    [InlineData(MissingValueBehavior.EmptyString)]
    [InlineData(MissingValueBehavior.Warning)]
    [InlineData(MissingValueBehavior.Error)]
    public void DataAuthority_Unavailable_ancestors_are_distinct_from_missing_properties_and_never_use_root(MissingValueBehavior behavior)
    {
        var source = FlatDataTemplate("{../Name}", "prefix{../../Name}suffix", "{$.Name}");
        var options = new MergeOptions { MissingValues = behavior };
        const string json = """{"Name":"ROOT","Items":[{"$level":40,"Name":"Ivan"}]}""";
        if (behavior == MissingValueBehavior.Error)
        {
            MergeResult? result = null;
            var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, json, options));
            Assert.Null(result); Assert.Equal(2, error.Errors.Count);
            Assert.Equal(new[] { "{../Name}", "{../../Name}" }, error.Errors.Select(issue => issue.Field));
            Assert.All(error.Errors, issue =>
            {
                Assert.Equal(DiagnosticCodes.GenerationInvalidData, issue.Code);
                Assert.Equal("$.Items[0]", issue.DataPath);
                Assert.Contains("data ancestor", issue.Message); Assert.Contains("unavailable", issue.Message);
            });
            Assert.Contains("depth 1", error.Errors[0].Message); Assert.Contains("depth 2", error.Errors[1].Message);
            Assert.Contains("B3", error.Errors[0].Location); Assert.Contains("C3", error.Errors[1].Location);
            return;
        }
        var merged = _fixture.Merge(source, json, options);
        using var package = SpreadsheetDocument.Open(new MemoryStream(merged.Document), false);
        var native = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>()
            .Single(cell => cell.CellReference == "A1");
        Assert.Null(native.CellValue); Assert.Null(native.InlineString);
        Assert.Equal("prefixsuffix", _fixture.ReadCellText(merged.Document, "B1"));
        Assert.Equal("ROOT", _fixture.ReadCellText(merged.Document, "C1"));
        Assert.Equal(2, merged.Warnings.Count);
        Assert.All(merged.Warnings, warning =>
        {
            Assert.Equal(DiagnosticCodes.XlsxMissingValue, warning.Code); Assert.Equal("$.Items[0]", warning.DataPath);
            Assert.Contains("data ancestor", warning.Message); Assert.Contains("unavailable", warning.Message);
        });
        Assert.Contains("depth 1", merged.Warnings[0].Message); Assert.Contains("depth 2", merged.Warnings[1].Message);
    }

    [Theory]
    [InlineData(MissingValueBehavior.EmptyString, false)] [InlineData(MissingValueBehavior.Warning, false)]
    [InlineData(MissingValueBehavior.Error, false)] [InlineData(MissingValueBehavior.Error, true)]
    public void DataAuthority_Existing_ancestor_missing_or_null_property_retains_public_missing_values_policy(MissingValueBehavior behavior, bool explicitNull)
    {
        var source = FlatDataTemplate("{../Rate}");
        var json = "{\"Rate\":777,\"Items\":[{\"$level\":20" + (explicitNull ? ",\"Rate\":null" : "") + "},{\"$level\":40}]}";
        var options = new MergeOptions { MissingValues = behavior };
        if (behavior == MissingValueBehavior.Error && !explicitNull)
        {
            var issue = Assert.Single(Assert.Throws<TemplateException>(() => _fixture.Merge(source, json, options)).Errors);
            Assert.Equal("$.Items[0].Rate", issue.DataPath);
            Assert.DoesNotContain("ancestor", issue.Message);
            return;
        }
        var result = _fixture.Merge(source, json, options);
        if (behavior == MissingValueBehavior.Warning && !explicitNull)
        {
            var warning = Assert.Single(result.Warnings);
            Assert.Equal("$.Items[0].Rate", warning.DataPath);
            Assert.Contains("scalar field is missing", warning.Message); Assert.DoesNotContain("ancestor", warning.Message);
        }
        else Assert.Empty(result.Warnings);
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.All(package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>(), cell =>
        { Assert.Null(cell.CellValue); Assert.Null(cell.InlineString); });
    }

    [Fact]
    public void DataAuthority_Full_paths_have_structural_identity_in_cells_and_owner_collections()
    {
        var first = new VerticalInstancePath(1, [0, 2, 4]);
        var equal = new VerticalInstancePath(1, ImmutableArray.CreateRange(new[] { 0, 2, 4 }));
        Assert.False(first.LogicalRows.Equals(equal.LogicalRows));
        Assert.Equal(first, equal); Assert.Equal(first.GetHashCode(), equal.GetHashCode());
        var source = new SourceCellId(3, new(7, 5));
        var cells = new Dictionary<CellInstanceId, string> { [new(source, first, null)] = "formula or prepared value" };
        Assert.Equal("formula or prepared value", cells[new(source, equal, null)]);
        var owners = new HashSet<XlsxContentOwner> { new(source, first, null, new(1, 1)), new(source, equal, null, new(1, 1)) };
        Assert.Single(owners);
        Assert.True(owners.Remove(new(source, equal, null, new(1, 1)))); Assert.Empty(owners);
        foreach (var different in new VerticalInstancePath[] { new(2, [0, 2, 4]), new(1, [2, 0, 4]),
            new(1, [0, 2]), new(1, [0, 2, 5]) }) Assert.NotEqual(first, different);
        Assert.NotEqual(new CellInstanceId(source, first, null), new CellInstanceId(new(4, source.Address), equal, null));
    }

    [Fact]
    public void DataAuthority_Flat_multirow_merge_and_formula_keep_duplicate_instances_under_unrendered_parent()
    {
        var source = _fixture.Create((workbook, worksheet) =>
        {
            var data = worksheet.Worksheet!.GetFirstChild<SheetData>()!; data.RemoveAllChildren();
            data.Append(new Row(_fixture.Text("A3", "$l(40).Items"), _fixture.Text("B3", "{Name}"), _fixture.Text("C3", "{Price}")) { RowIndex = 3 },
                new Row(_fixture.Text("A4", "$l(40).Items"), new Cell { CellReference = "C4", CellFormula = new CellFormula("C3") }) { RowIndex = 4 });
            worksheet.Worksheet.Append(new MergeCells(new MergeCell { Reference = "B3:B4" }));
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$C$4";
        });
        var result = _fixture.Merge(source, """
            {"Items":[{"$level":20,"Name":"Department"},{"$level":40,"Name":"Ivan","Price":11},
              {"$level":40,"Name":"Ivan","Price":22}]}
            """);
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package)); Assert.Empty(result.Warnings);
        var sheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal(new[] { "A1:A2", "A3:A4" }, sheet.Descendants<MergeCell>().Select(merge => merge.Reference!.Value));
        Assert.Equal("Ivan", _fixture.ReadCellText(result.Document, "A1")); Assert.Equal("Ivan", _fixture.ReadCellText(result.Document, "A3"));
        var cells = sheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("11", cells["B1"].CellValue!.Text); Assert.Equal("22", cells["B3"].CellValue!.Text);
        Assert.Equal(CellValues.Number, cells["B1"].DataType!.Value); Assert.Equal(CellValues.Number, cells["B3"].DataType!.Value);
        Assert.Equal("B1", cells["B2"].CellFormula!.Text); Assert.Equal("B3", cells["B4"].CellFormula!.Text);
    }

    [Theory]
    [InlineData(2, 2)] [InlineData(2, 0)] [InlineData(0, 2)]
    public void DataAuthority_Flat_tuple_binding_uses_current_record_and_unrendered_ancestor_first_record(int companies, int criteria)
    {
        var fixture = new XlsxTwoAxisFixture();
        var source = fixture.Create((_, worksheet) =>
        {
            var row = worksheet.Worksheet!.Descendants<Row>().Single(item => item.RowIndex!.Value == 10);
            row.Elements<Cell>().Single(cell => cell.CellReference == "A10").InlineString!.Text!.Text = "$l(40).Items";
            row.Elements<Cell>().Single(cell => cell.CellReference == "C10").InlineString!.Text!.Text = "{../Quantity}";
        });
        var arity = Math.Max(1, companies * criteria);
        var json = JsonSerializer.Serialize(new
        {
            Bids = Enumerable.Range(0, companies).Select(_ => new { Company = "header parent" }),
            Criteria = Enumerable.Range(0, criteria).Select(_ => new { Name = "header leaf" }),
            Items = Enumerable.Range(0, 3).SelectMany(item => Enumerable.Range(0, arity).Select(tuple =>
                new Dictionary<string, object> { ["$level"] = item == 0 ? 20 : 40, ["$path"] = "group" + item,
                    ["Name"] = item == 0 ? "Department" : tuple == 0 ? "Ivan" : "tuple " + tuple,
                    ["Quantity"] = tuple == 0 ? 17 : 999, ["Price"] = item * 100 + tuple }))
        });
        var result = _fixture.Merge(source, json);
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package)); Assert.Empty(result.Warnings);
        var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>()
            .ToDictionary(cell => cell.CellReference!.Value!);
        foreach (var outputRow in new[] { 4, 5 })
        {
            Assert.Equal("Ivan", _fixture.ReadCellText(result.Document, "A" + outputRow));
            Assert.Equal("17", cells["B" + outputRow].CellValue!.Text);
        }
        if (companies > 0 && criteria > 0)
            foreach (var outputRow in new[] { 4, 5 })
                Assert.Equal(Enumerable.Range((outputRow - 3) * 100, arity).Select(value => value.ToString()),
                    new[] { "D", "E", "G", "H" }.Select(column => cells[column + outputRow].CellValue!.Text));
    }

    private byte[] FlatDataTemplate(params string[] fields) => _fixture.Create((workbook, worksheet) =>
    {
        var row = worksheet.Worksheet!.Descendants<Row>().Single(); row.RemoveAllChildren();
        row.Append(_fixture.Text("A3", "$l(40).Items"));
        for (var index = 0; index < fields.Length; index++) row.Append(_fixture.Text(((char)('B' + index)) + "3", fields[index]));
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$" + (char)('A' + fields.Length) + "$3";
    });
}
