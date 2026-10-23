using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012FormulaTests
{
    [Theory]
    [InlineData("SUM(C3:$D$4)", "SUM(E5:$D$4)")]
    [InlineData("SUM($D3:C$4)", "SUM($D5:E$4)")]
    [InlineData("SUM(D3:$F$5)", "SUM(F5:$F$5)")]
    public void EndpointOrder_Shared_follower_crosses_or_meets_fixed_endpoints(string master, string expected)
    {
        var bytes = SharedCopyWorkbook(master);
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(
            services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var follower = plan.Sheets[0].Cells.Single(cell => cell.Source.Id.Address == new XlsxAddress(5, 4)).Formula!;
        Assert.Equal(expected, follower.Text);
        Assert.True(Assert.Single(follower.References).IsRange);
    }

    [Fact]
    public void EndpointOrder_Expanded_geometry_restores_column_orientation_with_equal_source_rows()
    {
        var fixture = new XlsxTwoAxisFixture();
        var sheet = fixture.Plan(fixture.Create()).Sheets.Single();
        using var services = _fixture.Services();
        using var json = JsonDocument.Parse(fixture.Data());
        var context = new XlsxOperationContext(new MergeOptions());
        var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
        var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var bound = planner.BindSheet(sheet, json.RootElement, context);
        var positions = geometry.ExpansionGeometry(sheet, planner.ScheduleSheet(sheet, bound, context), geometry.HorizontalSpans(sheet, bound));
        var cell = geometry.EmitCells(sheet, positions).Rows.SelectMany(row => row.Cells)
            .Single(cell => cell.Source.Id.Address == new XlsxAddress(9, 2));
        var formula = services.GetRequiredService<XlsxFormulaReferences>().Read("SUM($E10:D$10)", "B9");
        Assert.Equal("SUM($J4:C$5)", services.GetRequiredService<XlsxFormulaPlanner>().BindFormula(formula, cell, positions));
    }

    [Fact]
    public void EndpointOrder_Name_definition_relocates_while_name_identity_and_text_survive()
    {
        var bytes = _fixture.Create((workbook, part) =>
        {
            var cell = part.Worksheet!.Descendants<Cell>().Single();
            cell.RemoveAllChildren(); cell.DataType = null;
            cell.CellFormula = new CellFormula("SUM(Area)+Report!Area+LEN(\"Area\")");
            workbook.Workbook!.DefinedNames!.Append(new DefinedName("Report!$D$5:$C$3") { Name = "Area" });
        });
        using var services = _fixture.Services();
        using var source = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(
            services.GetRequiredService<XlsxPackageReader>().Read(source, context), context);
        var parsed = plan.Sheets[0].Cells.Single().Formula!;
        Assert.All(parsed.References, reference => { Assert.Equal(1, reference.NameIndex); Assert.Equal(1u, reference.TargetSheetId); });
        Assert.Equal(new[] { "Area", "Report!Area" }, parsed.References.Select(reference => parsed.Text.Substring(reference.Start, reference.Length)));
        var result = _fixture.Merge(bytes, "{}");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Assert.Equal("Report!$C$3:$B$1", package.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "Area").Text);
        Assert.Equal(parsed.Text, package.WorkbookPart.WorksheetParts.Single().Worksheet!.Descendants<CellFormula>().Single().Text);
    }

    [Theory]
    [InlineData("A3:$C$3", 0, 4, "E3:$C$3", "F4:$C$3")]
    [InlineData("$C$3:E3", 0, -4, "$C$3:A3", "$C$3:B4")]
    [InlineData("C1:$C$3", 4, 0, "C5:$C$3", "D6:$C$3")]
    [InlineData("$C$3:C5", -4, 0, "$C$3:C1", "$C$3:D2")]
    [InlineData("A1:$C$3", 4, 4, "E5:$C$3", "F6:$C$3")]
    [InlineData("$C$3:E5", -4, -4, "$C$3:A1", "$C$3:B2")]
    [InlineData("E1:$C$3", 1, 1, "F2:$C$3", "G3:$C$3")]
    [InlineData("A3:$C$3", 0, 2, "C3:$C$3", "D4:$C$3")]
    [InlineData("$E5:C$3", -2, 2, "$E3:E$3", "$E4:F$3")]
    public void EndpointOrder_Copy_reparse_and_copy_again_preserve_endpoint_identity(string operand,
        int rows, int columns, string copied, string next)
    {
        using var services = _fixture.Services();
        var references = services.GetRequiredService<XlsxFormulaReferences>();
        var parsed = references.Read("SUM('Report'!" + operand + ")+LOG10(1E10)+\"A1\"", "F5");
        var text = references.Copy(parsed, rows, columns, "F5");
        Assert.Equal("SUM('Report'!" + copied + ")+LOG10(1E10)+\"A1\"", text);
        var reparsed = references.ResolveSheets(references.Read(text, "F5"), 1,
            new Dictionary<string, uint> { ["Report"] = 1 }, "F5");
        var token = Assert.Single(reparsed.References);
        Assert.True(token.IsRange);
        Assert.Equal("'Report'!" + copied, text.Substring(token.Start, token.Length));
        Assert.Equal(1u, token.TargetSheetId);
        Assert.Equal(parsed.References[0].First.RowAbsolute, token.First.RowAbsolute);
        Assert.Equal(parsed.References[0].First.ColumnAbsolute, token.First.ColumnAbsolute);
        Assert.Equal(parsed.References[0].Last.RowAbsolute, token.Last.RowAbsolute);
        Assert.Equal(parsed.References[0].Last.ColumnAbsolute, token.Last.ColumnAbsolute);
        Assert.Equal("SUM('Report'!" + next + ")+LOG10(1E10)+\"A1\"", references.Copy(reparsed, 1, 1, "F5"));
    }

    [Theory]
    [InlineData("E3:$C$5", "D1:$B$3")]
    [InlineData("C5:$E$3", "B3:$D$1")]
    [InlineData("E5:$C$3", "D3:$B$1")]
    [InlineData("$E5:C$3", "$D3:B$1")]
    [InlineData("C3:$E$5", "B1:$D$3")]
    [InlineData("C3:$C$3", "B1:$B$1")]
    public void EndpointOrder_Relocation_preserves_each_axis_flags_and_range_syntax(string operand, string expected)
    {
        var bytes = _fixture.Create((workbook, part) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "Report!$B$3:$F$5";
            var data = part.Worksheet!.GetFirstChild<SheetData>()!;
            data.RemoveAllChildren();
            data.Append(new Row(new Cell { CellReference = "F5",
                CellFormula = new CellFormula("SUM('Report'!" + operand + ")+LOG10(1E10)+\"C3\"") }) { RowIndex = 5 });
        });
        var result = _fixture.Merge(bytes, "{}");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var formula = Assert.Single(package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<CellFormula>());
        Assert.Equal("SUM('Report'!" + expected + ")+LOG10(1E10)+\"C3\"", formula.Text);
    }

    [Theory]
    [InlineData("E3:A3")]
    [InlineData("E5:C2")]
    public void EndpointOrder_Crop_still_rejects_descending_operands(string operand)
    {
        var bytes = _fixture.Create((_, part) =>
        {
            var cell = part.Worksheet!.Descendants<Cell>().Single();
            cell.RemoveAllChildren(); cell.DataType = null;
            cell.CellFormula = new CellFormula("SUM(" + operand + ")");
        });
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(bytes, "{}"));
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
        Assert.Contains(operand, error.Message);
    }
}
