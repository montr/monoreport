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

public sealed partial class Feature012FormulaTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Names_Local_scope_shadows_workbook_scope_using_actual_sheet_order_and_preserves_operands(bool reverse)
    {
        const string formula = "tax+'O''Brien Q4'!tAX+OnlyGlobal+Alpha!OnlyGlobal+LEN(\"Tax\")+LOG10(1E10)";
        var source = NameScopeWorkbook(formula, reverse);
        using var services = _fixture.Services();
        using var input = new MemoryStream(source);
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
        using var original = SpreadsheetDocument.Open(new MemoryStream(source), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(original, context), context);
        var parsed = plan.Sheets.Single(sheet => sheet.Source.Id == 73).Cells.Single(cell => cell.Formula is not null).Formula!;
        Assert.Equal(new uint?[] { 5, 41, 73, 73 }, parsed.References.Select(reference => reference.TargetSheetId));
        var alpha = plan.Sheets.Single(sheet => sheet.Source.Id == 41).Cells.Single(cell => cell.Formula is not null).Formula!;
        Assert.Equal(new uint?[] { 41, 73, 41 }, alpha.References.Select(reference => reference.TargetSheetId));
        Assert.All(parsed.References, reference => Assert.NotNull(reference.NameIndex));
        var result = _fixture.Merge(source, "{}");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Assert.Equal(formula, CrossSheetCells(package, "Zulu")["A1"].CellFormula!.Text);
        Assert.Equal("Tax+OnlyGlobal+'O''Brien Q4'!OtherOnly", CrossSheetCells(package, "Alpha")["A1"].CellFormula!.Text);
        Assert.Equal("Tax", CrossSheetCells(package, "O'Brien Q4")["A1"].CellFormula!.Text);
        var names = package.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().ToArray();
        Assert.Equal("'Alpha'!$B$1", names.Single(name => name.Name == "Tax" && name.LocalSheetId is null).Text);
        Assert.Equal("'O''Brien Q4'!$B$1", names.Single(name => name.Name == "Tax" && name.LocalSheetId?.Value == (reverse ? 2u : 0u)).Text);
        Assert.Equal("'Alpha'!$C$2", names.Single(name => name.Name == "Tax" && name.LocalSheetId?.Value == 1).Text);
        Assert.Equal("'Zulu'!$C$3", names.Single(name => name.Name == "OnlyGlobal").Text);
        Assert.Equal("'Alpha'!$A$2:$C$3", names.Single(name => name.Name == "Unused").Text);
        var originals = original.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().ToArray();
        Assert.Equal(originals.Length, names.Length);
        for (var i = 0; i < names.Length; i++)
            Assert.Equal(originals[i].GetAttributes(), names[i].GetAttributes());
        Assert.True(names.Single(name => name.Name == "Unused").Hidden!.Value);
        Assert.Equal("retained comment", names.Single(name => name.Name == "Unused").Comment!.Value);
        Assert.Equal(original.WorkbookPart.Workbook.Sheets!.OuterXml, package.WorkbookPart.Workbook.Sheets!.OuterXml);
        Assert.Empty(result.Warnings);
    }

    [Theory]
    [InlineData("A1Suffix")]
    [InlineData("Ω.rate")]
    [InlineData("_Rate")]
    public void Names_Identifiers_remain_distinct_from_A1_functions_strings_and_copy_displacement(string name)
    {
        var source = _fixture.Create((workbook, part) =>
        {
            var cell = part.Worksheet!.Descendants<Cell>().Single();
            cell.RemoveAllChildren(); cell.DataType = null;
            cell.CellFormula = new CellFormula(name + "+C3+LOG10(1E10)+LEN(\"" + name + "\")");
            workbook.Workbook!.DefinedNames!.Append(new DefinedName("Report!$D$5") { Name = name });
        });
        var result = _fixture.Merge(source, "{}");
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Assert.Equal(name + "+B1+LOG10(1E10)+LEN(\"" + name + "\")", CrossSheetCells(package, "Report")["A1"].CellFormula!.Text);
        using var services = _fixture.Services();
        using var original = SpreadsheetDocument.Open(new MemoryStream(source), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(original, context), context);
        var formula = plan.Sheets[0].Cells.Single().Formula!;
        Assert.Equal(name + "+D4+LOG10(1E10)+LEN(\"" + name + "\")", services.GetRequiredService<XlsxFormulaReferences>().Copy(formula, 1, 1, "Report B3"));
    }

    [Theory]
    [InlineData(1, 1, 1)]
    [InlineData(2, 3, 3)]
    [InlineData(3, 0, 0)]
    public void Names_Definitions_follow_final_geometry_while_repeated_and_cross_sheet_uses_keep_global_meaning(int rows, int criteria, int consumers)
    {
        var source = EditNamesWorkbook(CrossSheetTemplate("SUM(Body)+Tail+Report!F12", consumerFormula: "SUM(Body)+Rate", leafFormula: "Rate+Report!$C10"), workbook =>
            // Empty child collections require an explicitly retained parent-only source range.
            workbook.Workbook!.DefinedNames!.Append(new DefinedName(criteria == 0 ? "Report!$D$10:$D$10" : "Report!$D$10:$E$10") { Name = "Body" },
                new DefinedName("Consumers!$D$8") { Name = "Tail" }, new DefinedName("'O''Brien Summary'!$I$8") { Name = "Rate" }));
        var json = CrossSheetData(rows, criteria, consumers);
        var result = _fixture.Merge(source, json);
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var lastColumn = criteria == 0 ? "D" : criteria == 1 ? "F" : "J";
        var tailColumn = criteria == 0 ? "E" : criteria == 1 ? "G" : "K";
        var names = package.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().Where(name => name.Name != "_xlnm.Print_Area").ToDictionary(name => name.Name!.Value!);
        Assert.Equal($"SUM(Body)+Tail+Report!{tailColumn}{rows + 5}", CrossSheetCells(package, "O'Brien Summary")["A2"].CellFormula!.Text);
        Assert.Equal($"Report!$C$4:${lastColumn}${rows + 3}", names["Body"].Text);
        Assert.Equal($"Consumers!$A${consumers + 5}", names["Tail"].Text);
        Assert.Equal("'O''Brien Summary'!$B$1", names["Rate"].Text);
        var report = CrossSheetCells(package, "Report");
        for (var row = 4; row < rows + 4; row++)
            for (var parent = 0; parent < 2; parent++)
                for (var child = 1; child <= criteria; child++)
                    Assert.Equal($"Rate+Report!$B{row}", report[((char)('C' + parent * (criteria + 1) + child)).ToString() + row].CellFormula!.Text);
        var consumer = CrossSheetCells(package, "Consumers");
        for (var row = 2; row < consumers + 2; row++) Assert.Equal("SUM(Body)+Rate", consumer["B" + row].CellFormula!.Text);
        Assert.All(package.WorkbookPart.WorksheetParts.SelectMany(part => part.Worksheet!.Descendants<Cell>()).Where(cell => cell.CellFormula is not null), cell => Assert.Null(cell.CellValue));
        Assert.Empty(result.Warnings);
        if (rows == 2 && Environment.GetEnvironmentVariable("MONOREPORT_XLSX_EVIDENCE") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, "names-template.xlsx"), source);
            File.WriteAllText(Path.Combine(directory, "names-data.json"), json);
            File.WriteAllBytes(Path.Combine(directory, "names-generated.xlsx"), result.Document);
        }
    }

    [Theory]
    [InlineData("SUM(Report!$B$3:$C$3)")]
    [InlineData("Budget")]
    [InlineData("OtherName")]
    [InlineData("Report!$B$3+1")]
    [InlineData("Report!B3")]
    [InlineData("Report!$B3")]
    [InlineData("Report!B$3")]
    [InlineData("$B$3")]
    [InlineData("Report!$B$3,Report!$C$3")]
    [InlineData("Report!$B:$D")]
    [InlineData("'[Other.xlsx]Report'!$B$3")]
    [InlineData("Report:Other!$B$3")]
    [InlineData("'Missing'!$B$3")]
    [InlineData("#REF!")]
    [InlineData("42")]
    [InlineData("")]
    [InlineData("INDIRECT(\"B3\")")]
    public void Names_Unsupported_definitions_fail_before_output_even_when_unused(string definition)
    {
        var source = _fixture.Create((workbook, _) => workbook.Workbook!.DefinedNames!.Append(new DefinedName(definition) { Name = "Budget" }));
        using var services = _fixture.Services();
        using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        var issue = Assert.Single(analysis.Errors);
        Assert.False(analysis.Valid);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
        Assert.Contains("defined name 'Budget'", issue.Location);
        Assert.Equal("/xl/workbook.xml", issue.Part);
        if (definition.Length > 0) Assert.Contains(definition, issue.Message);
        Assert.Contains(definition.Contains("Missing", StringComparison.Ordinal) ? "Correct the worksheet name" : "absolute A1", issue.Message);
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source));
        Assert.Null(result);
        Assert.Equal(issue.Message, Assert.Single(error.Errors).Message);
        Assert.Equal(source, input.ToArray());
    }

    [Theory]
    [InlineData("Report!$C$10", 2, 3, false)]
    [InlineData("Report!$E$10:$E$10", 2, 3, false)]
    [InlineData("Report!$C$10:$C$10", 0, 3, false)]
    [InlineData("Report!$E$10", 1, 0, false)]
    [InlineData("Report!$A$1", 2, 3, false)]
    [InlineData("Report!$A$9:$C$10", 2, 3, false)]
    [InlineData("Report!$E$7", 2, 0, true)]
    [InlineData("Report!$D$10:$E$10", 3, 0, false)]
    [InlineData("Report!$E$10:$D$10", 3, 0, false)]
    public void Names_Unsafe_final_images_reject_instead_of_selecting_one_owner_or_filling_gaps(string definition, int rows, int criteria, bool childFirst)
    {
        var source = EditNamesWorkbook(CrossSheetTemplate("1", childFirst), workbook =>
            workbook.Workbook!.DefinedNames!.Append(new DefinedName(definition) { Name = "UnsafeTarget" }));
        using var services = _fixture.Services();
        using var input = new MemoryStream(source);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.Equal(!definition.Contains("$A$", StringComparison.Ordinal), analysis.Valid);
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, CrossSheetData(rows, criteria, 0)));
        Assert.Null(result);
        var issue = Assert.Single(error.Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
        Assert.Contains("defined name 'UnsafeTarget'", issue.Location);
        Assert.Contains(definition, issue.Message);
        Assert.Contains("retained contiguous A1 range", issue.Message);
        Assert.Equal(source, input.ToArray());
    }

    [Theory]
    [InlineData("NoSuchName")]
    [InlineData("Alpha!OtherOnly")]
    [InlineData("'Missing'!Tax")]
    public void Names_Unresolved_operands_keep_formula_source_and_do_not_use_another_sheets_local_name(string formula)
    {
        var source = NameScopeWorkbook(formula, false);
        using var services = _fixture.Services();
        using var input = new MemoryStream(source);
        var issue = Assert.Single(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
        Assert.Contains("sheet 'Zulu', B3", issue.Location);
        Assert.Contains(formula, issue.Message);
        Assert.Contains("A1", issue.Message);
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source, "{}"));
        Assert.Equal(issue.Message, Assert.Single(error.Errors).Message);
    }

    [Theory]
    [InlineData("duplicate")]
    [InlineData("bad-scope")]
    [InlineData("macro")]
    [InlineData("A1")]
    [InlineData("R1C1")]
    public void Names_Invalid_metadata_or_duplicate_scope_is_actionable(string scenario)
    {
        var source = _fixture.Create((workbook, _) =>
        {
            var name = new DefinedName("Report!$B$3") { Name = scenario is "A1" or "R1C1" ? scenario : "Budget" };
            if (scenario == "bad-scope") name.LocalSheetId = 99;
            if (scenario == "macro") name.Function = true;
            workbook.Workbook!.DefinedNames!.Append(name);
            if (scenario == "duplicate") workbook.Workbook.DefinedNames.Append(new DefinedName("Report!$C$3") { Name = "bUDGET" });
        });
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source));
        var issue = Assert.Single(error.Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
        Assert.Contains("defined name", issue.Location);
        Assert.Contains(scenario == "bad-scope" ? "zero-based workbook position" : scenario == "duplicate" ? "same scope" : scenario == "macro" ? "ordinary cell/range name" : "valid simple identifier", issue.Message);
    }

    [Fact]
    public void Names_Cell_definition_follows_the_unique_repeated_instance_and_keeps_its_operand()
    {
        var source = EditNamesWorkbook(CrossSheetTemplate("Budget+Report!C10"), workbook =>
            workbook.Workbook!.DefinedNames!.Append(new DefinedName("Report!$C$10") { Name = "Budget" }));
        var result = _fixture.Merge(source, CrossSheetData(1, 1, 0));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Assert.Equal("Report!$B$4", package.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "Budget").Text);
        Assert.Equal("Budget+Report!B4", CrossSheetCells(package, "O'Brien Summary")["A2"].CellFormula!.Text);
    }

    [Theory]
    [InlineData("Report!$C$8:$C$13", true)]
    [InlineData("Report!$C$10:$C$11", false)]
    public void Names_Nested_group_uses_cannot_rebind_a_fixed_name_to_the_current_group(string definition, bool valid)
    {
        var source = GroupTemplate(customize: (workbook, part) =>
        {
            workbook.Workbook!.DefinedNames!.Append(new DefinedName(definition) { Name = "AllData" });
            part.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference == "C9").CellFormula!.Text = "SUM(AllData)";
        });
        if (valid)
        {
            var result = _fixture.Merge(source, GroupData());
            using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
            Assert.Empty(new OpenXmlValidator().Validate(package));
            Assert.Equal("Report!$B$2:$B$19", package.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "AllData").Text);
            var cells = CrossSheetCells(package, "Report");
            foreach (var row in new[] { 3, 9, 15 }) Assert.Equal("SUM(AllData)", cells["B" + row].CellFormula!.Text);
        }
        else
        {
            MergeResult? result = null; var original = source.ToArray();
            var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, GroupData()));
            Assert.Null(result); Assert.Equal(original, source);
            Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
            var issue = Assert.Single(error.Errors);
            Assert.Contains("defined name 'AllData'", issue.Location);
            Assert.Contains("retained contiguous A1 range", issue.Message);
        }
    }

    private byte[] NameScopeWorkbook(string formula, bool reverse) => EditNamesWorkbook(SheetIdentityTemplate(formula, reverse, otherFormula: "Tax"), workbook =>
    {
        var sheets = workbook.Workbook!.Sheets!.Elements<Sheet>().ToArray();
        uint Scope(string name) => (uint)Array.FindIndex(sheets, sheet => sheet.Name == name);
        var alpha = (WorksheetPart)workbook.GetPartById(sheets.Single(sheet => sheet.Name == "Alpha").Id!);
        alpha.Worksheet!.Descendants<Cell>().Single().CellValue = null;
        alpha.Worksheet.Descendants<Cell>().Single().CellFormula = new CellFormula("Tax+OnlyGlobal+'O''Brien Q4'!OtherOnly");
        var names = workbook.Workbook.DefinedNames!;
        names.PrependChild(new DefinedName("'Alpha'!$C$3") { Name = "Tax" });
        names.InsertAt(new DefinedName("'O''Brien Q4'!$C$3") { Name = "Tax", LocalSheetId = Scope("Zulu") }, 2);
        names.Append(new DefinedName("'Alpha'!$D$4") { Name = "Tax", LocalSheetId = Scope("O'Brien Q4") },
            new DefinedName("'Zulu'!$D$5") { Name = "OnlyGlobal" }, new DefinedName("'Alpha'!$D$5") { Name = "OtherOnly", LocalSheetId = Scope("O'Brien Q4") },
            new DefinedName("'Alpha'!$B$4:$D$5") { Name = "Unused", Hidden = true, Comment = "retained comment" });
    });

    private byte[] EditNamesWorkbook(byte[] source, Action<WorkbookPart> edit)
    {
        using var stream = new MemoryStream(); stream.Write(source); stream.Position = 0;
        using (var package = SpreadsheetDocument.Open(stream, true)) edit(package.WorkbookPart!);
        return stream.ToArray();
    }
}
