using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
    public void SheetIdentity_Resolves_actual_ids_in_workbook_order_without_changing_source_operands(bool reverse)
    {
        const string text = "SUM(Alpha!$C3:D$5)+'O''Brien Q4'!$B$3+zULu!B3+B3+\"Ghost!A1\"";
        var bytes = SheetIdentityTemplate(text, reverse, otherFormula: "Zulu!B3+B3");
        using var services = _fixture.Services();
        using var document = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var source = services.GetRequiredService<XlsxPackageReader>().Read(document, context);
        var originals = source.Sheets.Select(sheet => sheet.Xml.ToString()).ToArray();
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(source, context);
        Assert.Equal(reverse ? new[] { "Alpha", "O'Brien Q4", "Zulu" } : new[] { "Zulu", "O'Brien Q4", "Alpha" },
            plan.Sheets.Select(sheet => sheet.Source.Name));
        Assert.Equal(reverse ? new uint[] { 41, 5, 73 } : new uint[] { 73, 5, 41 }, plan.Sheets.Select(sheet => sheet.Source.Id));
        Assert.Equal(new[] { 0, 1, 2 }, plan.Sheets.Select(sheet => sheet.Source.PrintAreaNameIndex));
        var cell = plan.Sheets.Single(sheet => sheet.Source.Id == 73).Cells.Single(cell => cell.Formula is not null);
        var formula = cell.Formula!;
        Assert.Equal(new uint?[] { 41, 5, 73, 73 }, formula.References.Select(reference => reference.TargetSheetId));
        Assert.Equal(new[] { "Alpha!", "'O''Brien Q4'!", "zULu!", "" }, formula.References.Select(reference => reference.Prefix));
        Assert.Equal(new[] { "Alpha!$C3:D$5", "'O''Brien Q4'!$B$3", "zULu!B3", "B3" },
            formula.References.Select(reference => formula.Text.Substring(reference.Start, reference.Length)));
        var range = formula.References[0];
        Assert.True(range.IsRange);
        Assert.Equal(new XlsxFormulaPoint(new(3, 3), false, true), range.First);
        Assert.Equal(new XlsxFormulaPoint(new(5, 4), true, false), range.Last);
        Assert.False(formula.References[1].IsRange);
        Assert.True(formula.References[1].First.RowAbsolute);
        Assert.True(formula.References[1].First.ColumnAbsolute);
        Assert.Equal(text, formula.Text);
        var other = plan.Sheets.Single(sheet => sheet.Source.Id == 5).Cells.Single(cell => cell.Formula is not null).Formula!;
        Assert.Equal(new uint?[] { 73, 5 }, other.References.Select(reference => reference.TargetSheetId));
        services.GetRequiredService<XlsxFormulaPlanner>().ValidateStaticReferences(plan, context);
        Assert.Equal(originals, source.Sheets.Select(sheet => sheet.Xml.ToString()));
        Assert.All(formula.References, reference =>
        {
            var target = source.Sheets.Single(sheet => sheet.Id == reference.TargetSheetId);
            var workbookSheet = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.SheetId!.Value == target.Id);
            Assert.Equal(target.Part, document.WorkbookPart.GetPartById(workbookSheet.Id!).Uri.ToString());
        });
        Assert.Empty(context.Errors);
    }

    [Theory]
    [InlineData("'Plan ''A'''!$B3", "Plan 'A'")]
    [InlineData("'Bob''s ''Q4'''!B3", "Bob's 'Q4'")]
    [InlineData("'Данные 2026'!B3", "Данные 2026")]
    [InlineData("Данные!B3", "Данные")]
    [InlineData("'2026'!B3", "2026")]
    [InlineData("'A!B'!B3", "A!B")]
    [InlineData("LOG10!B3", "LOG10")]
    [InlineData("TRUE!B3", "TRUE")]
    public void SheetIdentity_Decodes_sheet_quoting_once_and_preserves_authored_prefix(string operand, string name)
    {
        var bytes = SheetIdentityTemplate(operand, targetName: name);
        using var services = _fixture.Services();
        using var document = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var source = services.GetRequiredService<XlsxPackageReader>().Read(document, context);
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(source, context);
        var formula = plan.Sheets[0].Cells.Single(cell => cell.Formula is not null).Formula!;
        var reference = Assert.Single(formula.References);
        Assert.Equal((uint)41, reference.TargetSheetId);
        Assert.Equal(name, source.Sheets.Single(sheet => sheet.Id == reference.TargetSheetId).Name);
        Assert.Equal(operand[..(operand.LastIndexOf('!') + 1)], reference.Prefix);
        Assert.Equal(operand, formula.Text);
        var copied = services.GetRequiredService<XlsxFormulaReferences>().Copy(formula, 1, 1, "Zulu B3");
        Assert.Equal(reference.Prefix + (reference.First.ColumnAbsolute ? "$B4" : "C4"), copied);
    }

    [Theory]
    [InlineData("Alpha!H8", true)]
    [InlineData("Alpha!B3", false)]
    public void SheetIdentity_Static_validation_uses_the_target_sheets_retained_area(string operand, bool valid)
    {
        var bytes = SheetIdentityTemplate(operand, targetArea: "$H$8:$J$10");
        using var services = _fixture.Services();
        using var document = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(
            services.GetRequiredService<XlsxPackageReader>().Read(document, context), context);
        if (valid)
        {
            services.GetRequiredService<XlsxFormulaPlanner>().ValidateStaticReferences(plan, context);
            Assert.Empty(context.Errors);
        }
        else
        {
            var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxFormulaPlanner>().ValidateStaticReferences(plan, context));
            var issue = Assert.Single(error.Errors);
            Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
            Assert.Contains("sheet 'Zulu', B3", issue.Location);
            Assert.Contains(operand, issue.Message);
            Assert.Contains("retained contiguous A1 range", issue.Message);
        }
    }

    [Theory]
    [InlineData("Missing!B3", "unresolved worksheet", "Correct the worksheet name")]
    [InlineData("'O''''Brien Q4'!B3", "unresolved worksheet", "Correct the worksheet name")]
    [InlineData("[Other.xlsx]Alpha!B3", "external workbook", "Move the required values")]
    [InlineData("'[Other.xlsx]Alpha'!B3", "external workbook", "Move the required values")]
    [InlineData("'C:\\Reports\\[Other.xlsx]Alpha'!B3", "external workbook", "Move the required values")]
    [InlineData("Zulu:Alpha!B3", "3-D", "explicit same-workbook A1 operands")]
    [InlineData("'Zulu:Alpha'!B3", "3-D", "explicit same-workbook A1 operands")]
    [InlineData("'Zulu':'O''Brien Q4'!B3", "3-D", "explicit same-workbook A1 operands")]
    [InlineData("Zulu:'O''Brien Q4'!B3", "3-D", "explicit same-workbook A1 operands")]
    [InlineData("'O''Brien Q4':Alpha!B3", "3-D", "explicit same-workbook A1 operands")]
    public void SheetIdentity_Unsupported_references_have_source_operand_and_remedy_before_output(string operand, string reason, string remedy)
    {
        var bytes = SheetIdentityTemplate("SUM(" + operand + ")");
        using var services = _fixture.Services();
        using var input = new MemoryStream(bytes);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(analysis.Valid);
        var issue = Assert.Single(analysis.Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
        Assert.Contains("sheet 'Zulu', B3", issue.Location);
        Assert.Contains("/xl/worksheets/", issue.Part);
        Assert.Contains(operand, issue.Message);
        Assert.Contains(reason, issue.Message);
        Assert.Contains(remedy, issue.Message);
        MergeResult? output = null;
        var error = Assert.Throws<TemplateException>(() => output = _fixture.Merge(bytes, "{}"));
        Assert.Null(output);
        Assert.Equal(issue.Message, Assert.Single(error.Errors).Message);
        Assert.Equal(bytes, input.ToArray());
    }

    [Fact]
    public void SheetIdentity_Resolved_qualification_cannot_fall_back_to_the_origins_geometry()
    {
        var bytes = SheetIdentityTemplate("Alpha!B3");
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(
            services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var source = plan.Sheets.Single(sheet => sheet.Source.Id == 73).Cells.Single(cell => cell.Formula is not null);
        var workbook = services.GetRequiredService<XlsxGeometryPlanner>().Preflight(plan, context)!;
        var origin = workbook.Sheets.Single(sheet => sheet.Geometry.Source.Id == 73);
        var placement = origin.Rows.SelectMany(row => row.Cells).Single(cell => cell.Source.Id == source.Source.Id);
        var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxFormulaPlanner>()
            .BindFormula(source.Formula!, placement, origin.Geometry));
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
        Assert.Contains("Alpha!B3", error.Message);
        Assert.Contains("no retained worksheet geometry", error.Message);
        Assert.Contains("retained A1 operand", error.Message);
        Assert.Contains("sheet 'Zulu', B3", error.Location);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SheetIdentity_Diagnostics_follow_workbook_order_instead_of_ids_names_or_relationships(bool reverse)
    {
        var bytes = SheetIdentityTemplate("FirstMissing!B3", reverse, otherFormula: "SecondMissing!B3");
        using var services = _fixture.Services();
        using var input = new MemoryStream(bytes);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.Equal(2, analysis.Errors.Count);
        Assert.Contains(reverse ? "SecondMissing!B3" : "FirstMissing!B3", analysis.Errors[0].Message);
        Assert.Contains(reverse ? "FirstMissing!B3" : "SecondMissing!B3", analysis.Errors[1].Message);
        Assert.All(analysis.Errors, issue =>
        {
            Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
            Assert.Contains("B3", issue.Location);
            Assert.Contains("Correct the worksheet name", issue.Message);
        });
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(bytes, "{}"));
        Assert.Equal(analysis.Errors.Select(issue => issue.Message), error.Errors.Select(issue => issue.Message));
        Assert.Equal(analysis.Errors.Select(issue => issue.Location), error.Errors.Select(issue => issue.Location));
    }

    private byte[] SheetIdentityTemplate(string formula, bool reverse = false, string targetName = "Alpha", string targetArea = "$B$3:$D$5", string? otherFormula = null) =>
        _fixture.Create((workbook, first) =>
        {
            var alpha = workbook.AddNewPart<WorksheetPart>();
            var quoted = workbook.AddNewPart<WorksheetPart>();
            var sheets = new[] { (Name: "Zulu", Id: 73u, Part: first), (Name: "O'Brien Q4", Id: 5u, Part: quoted), (Name: targetName, Id: 41u, Part: alpha) };
            // Part creation order, workbook order, IDs and alphabetical order deliberately differ.
            foreach (var sheet in sheets)
                sheet.Part.Worksheet = new Worksheet(new SheetData(new Row(new Cell
                {
                    CellReference = sheet.Id == 41 && targetArea != "$B$3:$D$5" ? "H8" : "B3",
                    CellFormula = sheet.Id == 73 ? new CellFormula(formula) : sheet.Id == 5 && otherFormula is not null ? new CellFormula(otherFormula) : null,
                    CellValue = new CellValue(sheet.Id.ToString(System.Globalization.CultureInfo.InvariantCulture))
                }) { RowIndex = sheet.Id == 41 && targetArea != "$B$3:$D$5" ? 8u : 3u }));
            if (reverse) Array.Reverse(sheets);
            workbook.Workbook!.Sheets!.RemoveAllChildren();
            workbook.Workbook.DefinedNames!.RemoveAllChildren();
            foreach (var (sheet, index) in sheets.Select((sheet, index) => (sheet, index)))
            {
                workbook.Workbook.Sheets.Append(new Sheet { Name = sheet.Name, SheetId = sheet.Id, Id = workbook.GetIdOfPart(sheet.Part) });
                workbook.Workbook.DefinedNames.Append(new DefinedName("'" + sheet.Name.Replace("'", "''", StringComparison.Ordinal) + "'!" + (sheet.Id == 41 ? targetArea : "$B$3:$D$5"))
                { Name = "_xlnm.Print_Area", LocalSheetId = (uint)index });
            }
        });
}
