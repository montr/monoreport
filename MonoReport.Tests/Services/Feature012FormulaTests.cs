using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services.Xlsx;
using MonoReport.Models.Xlsx;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012FormulaTests
{
    private readonly XlsxFixture _fixture = new();

    [Fact]
    public void Recalculation_Emits_scalar_formulas_without_caches_or_chain_and_requests_full_recalculation()
    {
        var fixture = new XlsxTwoAxisFixture();
        var bytes = fixture.Create((workbook, worksheet) =>
        {
            Cell Formula(string address, string text) => new() { CellReference = address, CellFormula = new CellFormula(text), CellValue = new CellValue("999999") };
            var rows = worksheet.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
            rows[9].Append(Formula("E9", "SUBTOTAL(9,E10:E10)"), Formula("F9", "SUM(D10:E10)"));
            rows[10].Append(Formula("F10", "SUM(D10:E10)"));
            rows[11].Append(Formula("F11", "COUNT(D10:E10)"));
            workbook.AddNewPart<CalculationChainPart>().CalculationChain = new CalculationChain(new CalculationCell { CellReference = "F9", SheetId = 1 });
        });
        var result = _fixture.Merge(bytes, fixture.Data());
        using var document = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var workbook = document.WorkbookPart!; Assert.Null(workbook.CalculationChainPart);
        Assert.Equal(CalculateModeValues.Auto, workbook.Workbook!.CalculationProperties!.CalculationMode!.Value);
        Assert.True(workbook.Workbook.CalculationProperties!.ForceFullCalculation!.Value);
        Assert.True(workbook.Workbook.CalculationProperties.FullCalculationOnLoad!.Value);
        var cells = workbook.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Where(cell => cell.CellFormula is not null).ToDictionary(cell => cell.CellReference!.Value!, cell => cell);
        Assert.Equal("SUM(C4:J5)", cells["K3"].CellFormula!.Text);
        Assert.Equal("SUM(C4:J4)", cells["K4"].CellFormula!.Text); Assert.Equal("SUM(C5:J5)", cells["K5"].CellFormula!.Text);
        Assert.Equal("COUNT(C4:J5)", cells["K6"].CellFormula!.Text);
        Assert.Equal("SUBTOTAL(9,D4:D5)", cells["D3"].CellFormula!.Text); Assert.Equal("SUBTOTAL(9,J4:J5)", cells["J3"].CellFormula!.Text);
        Assert.All(cells.Values, cell => { Assert.Null(cell.CellValue); Assert.Null(cell.CellFormula!.FormulaType); Assert.Null(cell.CellFormula.SharedIndex); });
        if (Environment.GetEnvironmentVariable("FEATURE012_EVIDENCE_DIRECTORY") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, "m2-template.xlsx"), bytes);
            File.WriteAllText(Path.Combine(directory, "m2-data.json"), fixture.Data());
            File.WriteAllBytes(Path.Combine(directory, "m2-generated.xlsx"), result.Document);
            File.WriteAllText(Path.Combine(directory, "m2-expected-values.json"), JsonSerializer.Serialize(new Dictionary<string, double>
            { ["K3"] = 1830, ["K4"] = 615, ["K5"] = 1215, ["K6"] = 12, ["D3"] = 300, ["J3"] = 310 }));
        }
    }

    [Theory]
    [InlineData("$E$10", 3, true)] [InlineData("E$10", 3, true)] [InlineData("$E10", 3, true)]
    [InlineData("SUM(E10:E10)", 3, false)] [InlineData("E10", 0, false)]
    [InlineData("A1", 3, false)] [InlineData("'Other'!E10", 3, true)]
    [InlineData("SUM($D10:E10)", 3, true)]
    public void Safety_Rejects_ambiguous_absolute_disjoint_empty_discarded_or_unbound_targets(string formula, int criteria, bool leaf)
    {
        var fixture = new XlsxTwoAxisFixture(); var sheet = fixture.Plan(fixture.Create()).Sheets.Single();
        using var services = _fixture.Services(); using var json = JsonDocument.Parse(fixture.Data(criteria));
        var context = new XlsxOperationContext(new MergeOptions()); var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
        var data = services.GetRequiredService<XlsxTemplatePlanner>().BindSheet(sheet, json.RootElement, context);
        var positions = geometry.ExpansionGeometry(sheet, services.GetRequiredService<XlsxTemplatePlanner>().ScheduleSheet(sheet, data, context), geometry.HorizontalSpans(sheet, data));
        var cell = geometry.EmitCells(sheet, positions).Rows.SelectMany(row => row.Cells).First(cell => cell.Source.Id.Address == (leaf ? new XlsxAddress(10, 5) : new XlsxAddress(9, 2)));
        var parsed = services.GetRequiredService<XlsxFormulaReferences>().Read(formula, "Report B9");
        var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxFormulaPlanner>().BindFormula(parsed, cell, positions));
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code); Assert.Contains("sheet 'Report'", error.Location);
        Assert.Contains("A1", error.Message);
    }

    [Theory]
    [InlineData(9)] [InlineData(11)]
    public void Global_Above_and_below_functions_use_the_complete_contiguous_region(int sourceRow)
    {
        var fixture = new XlsxTwoAxisFixture(); var sheet = fixture.Plan(fixture.Create()).Sheets.Single();
        using var services = _fixture.Services(); using var json = JsonDocument.Parse(fixture.Data());
        var context = new XlsxOperationContext(new MergeOptions()); var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
        var data = services.GetRequiredService<XlsxTemplatePlanner>().BindSheet(sheet, json.RootElement, context);
        var positions = geometry.ExpansionGeometry(sheet, services.GetRequiredService<XlsxTemplatePlanner>().ScheduleSheet(sheet, data, context), geometry.HorizontalSpans(sheet, data));
        var cell = geometry.EmitCells(sheet, positions).Rows.SelectMany(row => row.Cells).Single(cell => cell.Source.Id.Address == new XlsxAddress(sourceRow, 2));
        var parsed = services.GetRequiredService<XlsxFormulaReferences>().Read("SUM(D10:E10)+COUNT(D10:E10)+SUBTOTAL(9,D10:E10)+_xlfn.FUTURE(C10:C10)", "B" + sourceRow);
        Assert.Equal("SUM(C4:J5)+COUNT(C4:J5)+SUBTOTAL(9,C4:J5)+_xlfn.FUTURE(B4:B5)", services.GetRequiredService<XlsxFormulaPlanner>().BindFormula(parsed, cell, positions));
    }

    [Fact]
    public void Leaf_Selects_corresponding_tuple_and_row_then_crops_once()
    {
        var fixture = new XlsxTwoAxisFixture(); var sheet = fixture.Plan(fixture.Create()).Sheets.Single();
        using var services = _fixture.Services(); using var json = JsonDocument.Parse(fixture.Data());
        var context = new XlsxOperationContext(new MergeOptions()); var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
        var data = services.GetRequiredService<XlsxTemplatePlanner>().BindSheet(sheet, json.RootElement, context);
        var positions = geometry.ExpansionGeometry(sheet, services.GetRequiredService<XlsxTemplatePlanner>().ScheduleSheet(sheet, data, context), geometry.HorizontalSpans(sheet, data));
        var output = geometry.EmitCells(sheet, positions);
        var references = services.GetRequiredService<XlsxFormulaReferences>(); var formulas = services.GetRequiredService<XlsxFormulaPlanner>();
        var parsed = references.Read("E10+$C10+E$9+$B$9", "E10");
        var cells = output.Rows.SelectMany(row => row.Cells).Where(cell => cell.Source.Id.Address == new XlsxAddress(10, 5)).ToArray();
        Assert.Equal(new[] { "D4+$B4+D$3+$A$3", "E4+$B4+E$3+$A$3", "F4+$B4+F$3+$A$3", "H4+$B4+H$3+$A$3", "I4+$B4+I$3+$A$3", "J4+$B4+J$3+$A$3", "D5+$B5+D$3+$A$3", "E5+$B5+E$3+$A$3", "F5+$B5+F$3+$A$3", "H5+$B5+H$3+$A$3", "I5+$B5+I$3+$A$3", "J5+$B5+J$3+$A$3" },
            cells.Select(cell => formulas.BindFormula(parsed, cell, positions)));
    }

    [Theory]
    [InlineData(2, 3, "H12+$E$10+$E12+H$10+SUM(H12:$E$10)+\"E10\"")]
    [InlineData(-1, -1, "D9+$E$10+$E9+D$10+SUM(D9:$E$10)+\"E10\"")]
    public void Copy_Honors_each_dollar_flag_without_rewriting_literals(int rows, int columns, string expected)
    {
        using var services = _fixture.Services(); var references = services.GetRequiredService<XlsxFormulaReferences>();
        Assert.Equal(expected, references.Copy(references.Read("E10+$E$10+$E10+E$10+SUM(E10:$E$10)+\"E10\"", "Report F10"), rows, columns, "Report F10"));
    }

    [Fact]
    public void Copy_Rejects_reference_overflow_and_underflow()
    {
        using var services = _fixture.Services(); var references = services.GetRequiredService<XlsxFormulaReferences>();
        Assert.Throws<TemplateException>(() => references.Copy(references.Read("A1", "B1"), -1, 0, "B1"));
        Assert.Throws<TemplateException>(() => references.Copy(references.Read("XFD1048576", "B1"), 0, 1, "B1"));
    }

    [Theory]
    [InlineData("LOG10(E10)+1E10+SUM($D$7:E10)+\"E10\"", 2)]
    [InlineData("'A ''quoted'''!D10+Sheet1!$E$10", 2)]
    [InlineData("_xlfn.FUTURE(E10,\"a\"\"D8\")+SUBTOTAL(9,D10:E10)", 2)]
    public void Lexer_Preserves_nonreference_text_and_distinguishes_numbers_functions_and_strings(string formula, int count)
    {
        using var services = _fixture.Services(); var references = services.GetRequiredService<XlsxFormulaReferences>();
        var parsed = references.Read(formula, "Report E10"); Assert.Equal(count, parsed.References.Length);
        Assert.Equal(formula, references.Rewrite(parsed, token => formula.Substring(token.Start, token.Length)));
        Assert.All(parsed.References, token => Assert.Contains("10", formula.Substring(token.Start, token.Length)));
    }

    [Theory]
    [InlineData("SUM(A:A)")] [InlineData("Table1[Price]")] [InlineData("[Book]Sheet!A1")]
    [InlineData("A1#")] [InlineData("@A1")] [InlineData("SUM(A1")]
    [InlineData("\"unterminated")] [InlineData("1E+")] [InlineData("1.2.3")]
    [InlineData("_xlfn.INDIRECT(\"A1\")")] [InlineData("offset(A1,1,0)")]
    public void Lexer_Rejects_unsafe_or_malformed_operand_forms(string formula)
    {
        using var services = _fixture.Services();
        var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxFormulaReferences>().Read(formula, "Report E10"));
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code); Assert.Contains("E10", error.Location);
    }
}
