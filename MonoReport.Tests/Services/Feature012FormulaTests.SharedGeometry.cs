using System.Collections.Immutable;
using System.IO.Compression;
using System.Xml.Linq;
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
    [Fact]
    public void SharedGeometry_Later_master_overrides_holes_and_redundant_text_emit_scalar_package_and_roundtrip()
    {
        var source = SharedGeometryStaticWorkbook();
        using var services = _fixture.Services();
        using var input = new MemoryStream(source);
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
        var output = _fixture.Merge(source, "{}").Document;
        var expected = new Dictionary<string, string>
        {
            ["D1"] = "A1+$A$1", ["E1"] = "1+1", ["F1"] = "C1+$A$1",
            ["E2"] = "B2+$A$1", ["F2"] = "C2+$A$1"
        };
        AssertSharedGeometryPackage(output, expected, 11, "A1:F3");
        using var roundtrip = new MemoryStream(); roundtrip.Write(output); roundtrip.Position = 0;
        using (var package = SpreadsheetDocument.Open(roundtrip, true))
        {
            package.WorkbookPart!.Workbook!.Save();
            foreach (var sheet in package.WorkbookPart!.WorksheetParts) sheet.Worksheet!.Save();
        }
        AssertSharedGeometryPackage(roundtrip.ToArray(), expected, 11, "A1:F3");
        Assert.Equal(source, input.ToArray());
        ExportSharedGeometry("static", source, "{}", output);
        ExportSharedGeometry("roundtrip", source, "{}", roundtrip.ToArray());
    }

    [Theory]
    [InlineData(true)] [InlineData(false)]
    public void SharedGeometry_Cropped_master_or_follower_needs_no_replacement_master(bool croppedMaster)
    {
        var source = _fixture.Create((_, part) =>
        {
            part.Worksheet = new Worksheet(new SheetData(
                new Row(SharedInventoryCell("E2", croppedMaster ? "C3+$B$3" : "ignored", croppedMaster ? "E2:E4" : null, 7)) { RowIndex = 2 },
                new Row(new Cell { CellReference = "B3", CellValue = new CellValue("2") },
                    new Cell { CellReference = "C3", CellValue = new CellValue("3") },
                    SharedInventoryCell("E3", croppedMaster ? "" : "C4+$B$3", croppedMaster ? null : "E2:E4", 7)) { RowIndex = 3 },
                new Row(new Cell { CellReference = "C4", CellValue = new CellValue("4") }, SharedInventoryCell("E4", "", null, 7)) { RowIndex = 4 },
                new Row(new Cell { CellReference = "C5", CellValue = new CellValue("5") }) { RowIndex = 5 }));
        });
        // Keep E3:E4, while E2 is disposable. The same retained formulas result in either master position.
        source = _fixture.RewritePart(source, "xl/workbook.xml", xml => xml.Replace("$D$5", "$E$5", StringComparison.Ordinal));
        var output = _fixture.Merge(source, "{}").Document;
        AssertSharedGeometryPackage(output, new() { ["D1"] = "B2+$A$1", ["D2"] = "B3+$A$1" }, 6, "A1:D3");
        ExportSharedGeometry(croppedMaster ? "cropped-master" : "cropped-follower", source, "{}", output);
    }

    [Theory]
    [InlineData(0, 0)] [InlineData(0, 2)] [InlineData(1, 1)] [InlineData(3, 2)]
    public void SharedGeometry_One_group_spans_independent_runs_and_static_scope(int items, int other)
    {
        var source = _fixture.Create((workbook, part) =>
        {
            var rows = new SheetData();
            for (var row = 3; row <= 7; row++)
            {
                var cells = new List<Cell>();
                if (row != 5) cells.Add(_fixture.Text("A" + row, "$l(10)." + (row < 5 ? "Items" : "Other")));
                cells.Add(new Cell { CellReference = "C" + row, CellValue = new CellValue(row.ToString()) });
                cells.Add(SharedInventoryCell("E" + row, row == 4 ? "C4*2" : "ignored", row == 4 ? "E3:E7" : null, 7));
                rows.Append(new Row(cells) { RowIndex = (uint)row });
            }
            part.Worksheet = new Worksheet(rows);
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "Report!$B$3:$E$7";
        });
        var json = "{\"Items\":[" + string.Join(',', Enumerable.Repeat("{\"$level\":10}", items)) +
            "],\"Other\":[" + string.Join(',', Enumerable.Repeat("{\"$level\":10}", other)) + "]}";
        var result = RenderGroups(source, json);
        // Each emitted source row owns its own C operand; source and formula column crop to B and D.
        var count = items * 2 + 1 + other * 2;
        var expected = Enumerable.Range(1, count).ToDictionary(row => "D" + row, row => "B" + row + "*2");
        AssertSharedGeometryPackage(result.Document, expected, count * 2, "A1:D" + count);
        var formulas = result.Output.Sheets[0].Rows.SelectMany(row => row.Cells).Where(cell => cell.Formula is not null).ToArray();
        Assert.Equal(items * 2, formulas.Count(cell => cell.Id.Vertical?.RegionId == 0));
        Assert.Equal(other * 2, formulas.Count(cell => cell.Id.Vertical?.RegionId == 1));
        Assert.Single(formulas, cell => cell.Id.Vertical is null);
        Assert.Equal(items + other, formulas.Where(cell => cell.Id.Vertical is not null).Select(cell => cell.Id.Vertical).Distinct().Count());
        Assert.Empty(result.Context.Warnings);
        // Also prove the public single-level path consumes the same normalized formulas.
        AssertSharedGeometryPackage(_fixture.Merge(source, json).Document, expected, count * 2, "A1:D" + count);
        ExportSharedGeometry($"runs-{items}-{other}", source, json, result.Document);
    }

    [Fact]
    public void SharedGeometry_Nested_members_use_their_own_group_descendants_and_full_paths()
    {
        var source = GroupTemplate(customize: (_, part) =>
        {
            var rows = part.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
            rows[9].Elements<Cell>().Single(cell => cell.CellReference == "D9").Remove();
            rows[9].Append(SharedInventoryCell("D9", "ignored", null, 7));
            rows[12].Append(SharedInventoryCell("D12", "SUM($C$10:$C$11)", "D9:D12", 7));
        });
        var result = RenderGroups(source, GroupData());
        using var package = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var formulas = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>()
            .Where(cell => cell.CellFormula is not null).ToDictionary(cell => cell.CellReference!.Value!, cell => cell.CellFormula!.Text);
        Assert.Equal(22, formulas.Count);
        var expected = new Dictionary<string, string> { ["C3"] = "SUM($B$4:$B$7)", ["C8"] = "SUM($B$4:$B$7)",
            ["C9"] = "SUM($B$10:$B$11)", ["C12"] = "SUM($B$10:$B$11)", ["C15"] = "SUM($B$16:$B$17)", ["C18"] = "SUM($B$16:$B$17)" };
        foreach (var (address, formula) in expected) Assert.Equal(formula, formulas[address]);
        var placements = result.Output.Sheets[0].Rows.SelectMany(row => row.Cells).Where(cell => expected.ContainsKey(cell.Output.ToString())).ToArray();
        Assert.Equal(new[] { 0, 1, 2 }, placements[0].Id.Vertical!.LogicalRows);
        Assert.Equal(placements[0].Id.Vertical, placements[1].Id.Vertical);
        Assert.Equal(new[] { 0, 5, 6 }, placements[2].Id.Vertical!.LogicalRows);
        Assert.Equal(new[] { 8, 9, 10 }, placements[4].Id.Vertical!.LogicalRows);
        var allExpected = new Dictionary<string, string>(expected)
        {
            ["B1"] = "COUNT(B2:B19)", ["B2"] = "SUBTOTAL(9,B3:B12)", ["B3"] = "SUM(B4:B7)",
            ["B5"] = "B4*2", ["B7"] = "B6*2", ["B8"] = "SUBTOTAL(9,B4:B7)", ["B9"] = "SUM(B10:B11)",
            ["B11"] = "B10*2", ["B12"] = "SUBTOTAL(9,B10:B11)", ["B13"] = "SUBTOTAL(9,B3:B12)",
            ["B14"] = "SUBTOTAL(9,B15:B18)", ["B15"] = "SUM(B16:B17)", ["B17"] = "B16*2",
            ["B18"] = "SUBTOTAL(9,B16:B17)", ["B19"] = "SUBTOTAL(9,B15:B18)", ["B20"] = "SUM(B2:B19)"
        };
        AssertSharedGeometryPackage(result.Document, allExpected, 42, "A1:C20");
        // Public nested activation remains a separate boundary.
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(source, GroupData()));
        Assert.Contains("nested", error.Message, StringComparison.OrdinalIgnoreCase);
        ExportSharedGeometry("nested", source, GroupData(), result.Document);
    }

    [Theory]
    [InlineData("SUM(E$10:$D$10)", "SUM(J$4:$C$5)", 2)] // Equal row axis expands; columns descend.
    [InlineData("SUM(D$11:$E$9)", "SUM(C$6:$J$3)", 2)] // Rows descend independently.
    [InlineData("SUM(E$11:$D$9)", "SUM(J$6:$C$3)", 2)] // Both axes descend, mixed column flags preserved.
    [InlineData("SUM(D$9:$E$11)", "SUM(C$3:$J$6)", 2)]
    [InlineData("SUM(C$10:C$10)", "SUM(B$4:B$5)", 2)] // Equal endpoints remain a range and expand on the row axis.
    [InlineData("SUM(E$10:E$10)", "SUM(D$4:F$5)", 1)] // Equal axes expand independently within one contiguous horizontal parent.
    public void SharedGeometry_Output_preserves_endpoint_orientation_and_equal_axis_expansion(string formula, string expected, int companies)
    {
        var fixture = new XlsxTwoAxisFixture();
        var source = fixture.Create((_, part) =>
        {
            var rows = part.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
            rows[9].Append(SharedInventoryCell("F9", "redundant", null, 7));
            rows[11].Append(SharedInventoryCell("F11", formula, "F9:F11", 7));
        });
        var json = fixture.Data(companies: companies);
        var output = _fixture.Merge(source, json).Document;
        // Header 8 + criterion 6 + top 2 + details 20 + bottom 2 + tail 1 = 39 authored cells.
        var column = companies == 2 ? "K" : "G";
        AssertSharedGeometryPackage(output, new() { [column + "3"] = expected, [column + "6"] = expected }, companies == 2 ? 39 : 24, "A1:" + column + "7");
        ExportSharedGeometry("endpoints-" + formula.Replace('$', 'x').Replace(':', '-'), source, json, output);
    }

    [Theory]
    [InlineData("A$1", 3)] // Cropped.
    [InlineData("SUM(D$10:E$10)", 0)] // Incomplete source coverage with surviving parent columns.
    [InlineData("SUM(E$10:E$10)", 3)] // Disjoint leaf columns.
    [InlineData("SUM(E$10:E$10)", 0)] // Empty removed operand.
    [InlineData("$E$10", 3)] // Ambiguous scalar.
    [InlineData("E$7", 0)] // Displaced nonanchor identity after the parent merge collapses.
    public void SharedGeometry_Unsafe_shared_operands_keep_existing_diagnostics_and_no_output(string operand, int criteria)
    {
        var fixture = new XlsxTwoAxisFixture();
        var source = fixture.Create((_, part) =>
        {
            var rows = part.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
            rows[9].Append(SharedInventoryCell("F9", operand, "F9:F11", 7));
            rows[11].Append(SharedInventoryCell("F11", "", null, 7));
        });
        var original = source.ToArray(); MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, fixture.Data(criteria)));
        Assert.Null(result); Assert.Equal(original, source);
        Assert.All(error.Errors, issue => Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code));
        var issue = error.Errors.First(issue => issue.Location!.EndsWith(", F9", StringComparison.Ordinal));
        var reference = operand.StartsWith("SUM(", StringComparison.Ordinal) ? operand[4..^1] : operand;
        Assert.Contains(reference, issue.Message); Assert.Contains("retained contiguous A1 range", issue.Message);
    }

    [Fact]
    public void SharedGeometry_Copy_crossing_and_mixed_row_flags_survive_saved_output()
    {
        var source = _fixture.RewritePart(SharedGeometryStaticWorkbook(), "xl/worksheets/sheet1.xml", text => text.Replace("C4+$B$3", "SUM($C$3:C4)", StringComparison.Ordinal));
        var output = _fixture.Merge(source, "{}").Document;
        AssertSharedGeometryPackage(output, new() { ["D1"] = "SUM($B$1:A1)", ["E1"] = "1+1", ["F1"] = "SUM($B$1:C1)",
            ["E2"] = "SUM($B$1:B2)", ["F2"] = "SUM($B$1:C2)" }, 11, "A1:F3");
        ExportSharedGeometry("crossing", source, "{}", output);
    }

    [Fact]
    public void SharedGeometry_Horizontal_members_select_each_tuple_and_current_vertical_owner()
    {
        var fixture = new XlsxTwoAxisFixture();
        var source = fixture.Create((_, part) =>
        {
            var rows = part.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
            rows[9].Append(SharedInventoryCell("E9", "ignored", null, 7));
            rows[10].Elements<Cell>().Single(cell => cell.CellReference == "E10").Remove();
            rows[10].Append(SharedInventoryCell("E10", "COUNTA(E$8:E$8)", "E9:E10", 7));
        });
        var result = RenderGroups(source, fixture.Data());
        var expected = new Dictionary<string, string>();
        foreach (var column in new[] { "D", "E", "F", "H", "I", "J" })
            foreach (var row in new[] { 3, 4, 5 }) expected.Add(column + row, $"COUNTA({column}$2:{column}$2)");
        AssertSharedGeometryPackage(result.Document, expected, 43, "A1:K7");
        var body = result.Output.Sheets[0].Rows.SelectMany(row => row.Cells).Where(cell => cell.Source.Id.Address == new XlsxAddress(10, 5)).ToArray();
        Assert.Equal(12, body.Length); Assert.Equal(12, body.Select(cell => cell.Id).Distinct().Count());
        Assert.Equal(6, body.Select(cell => cell.Id.Horizontal).Distinct().Count());
        Assert.Equal(2, body.Select(cell => cell.Id.Vertical).Distinct().Count());
        AssertSharedGeometryPackage(_fixture.Merge(source, fixture.Data()).Document, expected, 43, "A1:K7");
        ExportSharedGeometry("tuples", source, fixture.Data(), result.Document);
    }

    [Fact]
    public void SharedGeometry_Mixed_endpoint_flags_cannot_select_incompatible_horizontal_owners()
    {
        var fixture = new XlsxTwoAxisFixture();
        var source = fixture.Create((_, part) =>
        {
            var row = part.Worksheet!.Descendants<Row>().Single(row => row.RowIndex!.Value == 10);
            row.Elements<Cell>().Single(cell => cell.CellReference == "E10").Remove();
            row.Append(SharedInventoryCell("E10", "SUM($D10:E10)", "E10", 7));
        });
        var original = source.ToArray(); MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(source, fixture.Data()));
        Assert.Null(result); Assert.Equal(original, source);
        Assert.All(error.Errors, issue =>
        {
            Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code); Assert.EndsWith(", E10", issue.Location);
            Assert.Contains("Mixed endpoint copy flags select incompatible repeated owners", issue.Message);
            Assert.Contains("Re-author", issue.Message);
        });
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void SharedGeometry_Incomplete_nested_source_coverage_rejects_for_plain_and_content_owners(bool merged)
    {
        var source = OwnerRangeBodyWorkbook("B2:B3");
        source = _fixture.RewritePart(source, "xl/worksheets/sheet1.xml", text =>
        {
            var xml = XDocument.Parse(text); var ns = xml.Root!.Name.Namespace;
            var formula = xml.Descendants(ns + "c").Single(cell => (string?)cell.Attribute("r") == "C3").Element(ns + "f")!;
            formula.SetAttributeValue("t", "shared"); formula.SetAttributeValue("si", 7); formula.SetAttributeValue("ref", "C3");
            if (merged) xml.Root.Add(new XElement(ns + "mergeCells", new XElement(ns + "mergeCell", new XAttribute("ref", "B2:C2"))));
            return xml.ToString();
        });
        AssertOwnerRangeUnsafe(source, """{"Items":[{"$level":10,"Name":"P"},{"$level":40,"Price":11}]}""", "C3", "B2:B3");
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void SharedGeometry_Missing_normalization_rejects_even_without_surviving_instances(bool finalization)
    {
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(SharedGeometryStaticWorkbook()), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var template = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        template = template with { Sheets = template.Sheets.Select(sheet => sheet with { Cells = sheet.Cells.Select(cell =>
            cell.SharedFormulaIndex is not null ? cell with { Formula = null } : cell).ToImmutableArray() }).ToImmutableArray() };
        var formulas = services.GetRequiredService<XlsxFormulaPlanner>();
        var error = Assert.Throws<TemplateException>(() =>
        {
            if (finalization) formulas.Finalize(template, new([], []), context);
            else formulas.ValidateStaticReferences(template, context);
        });
        Assert.Equal(4, error.Errors.Count);
        Assert.All(error.Errors, issue => { Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code); Assert.Contains("no normalized parsed formula", issue.Message); });
    }

    [Theory]
    [InlineData(false, false)] [InlineData(true, false)] [InlineData(false, true)]
    public void SharedGeometry_Missing_finalized_or_parsed_scalar_text_cannot_fall_back_to_source_XML(bool suppress, bool missingParsed)
    {
        var source = _fixture.Create((_, part) => part.Worksheet = new Worksheet(new SheetData(new Row(GroupFormula("B3", "1+1")) { RowIndex = 3 })));
        using var services = _fixture.Services();
        using var stream = new MemoryStream(); stream.Write(source); stream.Position = 0;
        using var package = SpreadsheetDocument.Open(stream, true);
        var context = new XlsxOperationContext(new MergeOptions());
        var template = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var output = services.GetRequiredService<XlsxGeometryPlanner>().Plan(template, ImmutableDictionary<SourceCellId, XlsxPreparedValue>.Empty, context);
        if (suppress) output = output with { Sheets = output.Sheets.Select(sheet => sheet with { Rows = sheet.Rows.Select(row => row with
            { Cells = row.Cells.Select(cell => cell with { Suppressed = true }).ToImmutableArray() }).ToImmutableArray() }).ToImmutableArray() };
        string ReadXml()
        {
            using var xmlStream = package.WorkbookPart!.WorksheetParts.Single().GetStream();
            using var reader = new StreamReader(xmlStream); return reader.ReadToEnd();
        }
        var original = ReadXml();
        if (missingParsed)
        {
            template = template with { Sheets = template.Sheets.Select(sheet => sheet with { Cells = sheet.Cells.Select(cell => cell with { Formula = null }).ToImmutableArray() }).ToImmutableArray() };
            var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxFormulaPlanner>().Finalize(template, output, context));
            Assert.Contains("no normalized parsed formula", error.Message);
        }
        else if (suppress)
        {
            services.GetRequiredService<XlsxDocumentRenderer>().Render(package, output, context);
            using var xmlStream = package.WorkbookPart!.WorksheetParts.Single().GetStream();
            Assert.Empty(XDocument.Load(xmlStream).Descendants(XName.Get("f", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
        }
        else
        {
            var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxDocumentRenderer>().Render(package, output, context));
            Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code); Assert.Contains("no finalized text", error.Message);
        }
        if (!suppress) Assert.Equal(original, ReadXml());
    }

    [Theory]
    [InlineData(null)] [InlineData(" ")]
    public void SharedGeometry_Renderer_checks_all_sheets_before_writing_any_sheet(string? missingText)
    {
        var source = _fixture.Create((workbook, part) =>
        {
            part.Worksheet = new Worksheet(new SheetData(new Row(GroupFormula("B3", "1+1")) { RowIndex = 3 }));
            var second = workbook.AddNewPart<WorksheetPart>();
            second.Worksheet = new Worksheet(new SheetData(new Row(GroupFormula("B3", "2+2")) { RowIndex = 3 }));
            workbook.Workbook!.Sheets!.Append(new Sheet { Name = "Other", SheetId = 7, Id = workbook.GetIdOfPart(second) });
            workbook.Workbook.DefinedNames!.Append(new DefinedName("Other!$B$3:$D$5") { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
        });
        using var stream = new MemoryStream(); stream.Write(source); stream.Position = 0;
        using var package = SpreadsheetDocument.Open(stream, true);
        using var services = _fixture.Services(); var context = new XlsxOperationContext(new MergeOptions());
        var template = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var output = services.GetRequiredService<XlsxGeometryPlanner>().Plan(template, ImmutableDictionary<SourceCellId, XlsxPreparedValue>.Empty, context);
        output = services.GetRequiredService<XlsxFormulaPlanner>().Finalize(template, output, context);
        output = output with { Sheets = output.Sheets.Select(sheet => sheet.Geometry.Source.Id != 7 ? sheet : sheet with { Rows = sheet.Rows.Select(row => row with
            { Cells = row.Cells.Select(cell => cell with { Formula = missingText }).ToImmutableArray() }).ToImmutableArray() }).ToImmutableArray() };
        string[] ReadXml() => package.WorkbookPart!.WorksheetParts.OrderBy(part => part.Uri.ToString()).Select(part =>
        {
            using var xml = part.GetStream(); using var reader = new StreamReader(xml); return reader.ReadToEnd();
        }).ToArray();
        var original = ReadXml();
        var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxDocumentRenderer>().Render(package, output, context));
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code); Assert.Contains("sheet 'Other', B3", error.Location);
        Assert.Equal(original, ReadXml());
    }

    private byte[] SharedGeometryStaticWorkbook() => _fixture.Create((workbook, part) =>
    {
        part.Worksheet = new Worksheet(new SheetData(
            new Row(new Cell { CellReference = "B3", CellValue = new CellValue("2") }, new Cell { CellReference = "C3", CellValue = new CellValue("3") },
                new Cell { CellReference = "D3", CellValue = new CellValue("4") }, SharedInventoryCell("E3", "redundant invalid text", null, 7),
                GroupFormula("F3", "1+1"), SharedInventoryCell("G3", "", null, 7)) { RowIndex = 3 },
            new Row(new Cell { CellReference = "B4", CellValue = new CellValue("5") }, new Cell { CellReference = "C4", CellValue = new CellValue("6") },
                new Cell { CellReference = "D4", CellValue = new CellValue("7") }, SharedInventoryCell("F4", "C4+$B$3", "E3:G4", 7),
                SharedInventoryCell("G4", "", null, 7)) { RowIndex = 4 }));
        workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "Report!$B$3:$G$5";
        workbook.AddNewPart<CalculationChainPart>().CalculationChain = new CalculationChain(new CalculationCell { CellReference = "F4", SheetId = 1 });
    });

    private void AssertSharedGeometryPackage(byte[] bytes, Dictionary<string, string> expected, int cells, string dimension)
    {
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var workbook = package.WorkbookPart!; Assert.Null(workbook.CalculationChainPart);
        var calculation = workbook.Workbook!.CalculationProperties!;
        Assert.Equal(CalculateModeValues.Auto, calculation.CalculationMode!.Value);
        Assert.True(calculation.FullCalculationOnLoad!.Value); Assert.True(calculation.ForceFullCalculation!.Value);
        var worksheet = workbook.WorksheetParts.Single();
        Assert.Equal(dimension, worksheet.Worksheet!.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal(cells, worksheet.Worksheet.Descendants<Cell>().Count());
        var bounds = dimension.Split(':');
        Assert.Equal("'Report'!$" + bounds[0][0] + "$" + bounds[0][1..] + ":$" + bounds[1][0] + "$" + bounds[1][1..],
            workbook.Workbook.DefinedNames!.Elements<DefinedName>().Single().Text);
        using var stream = worksheet.GetStream(); var xml = XDocument.Load(stream); var ns = xml.Root!.Name.Namespace;
        var formulas = xml.Descendants(ns + "c").Where(cell => cell.Element(ns + "f") is not null).ToArray();
        Assert.Equal(expected.OrderBy(pair => pair.Key), formulas.ToDictionary(cell => (string)cell.Attribute("r")!, cell => cell.Element(ns + "f")!.Value).OrderBy(pair => pair.Key));
        Assert.All(formulas, cell =>
        {
            var formula = Assert.Single(cell.Elements(ns + "f"));
            Assert.False(string.IsNullOrWhiteSpace(formula.Value)); Assert.Empty(formula.Attributes());
            Assert.Null(cell.Element(ns + "v")); Assert.Null(cell.Element(ns + "is")); Assert.Null(cell.Attribute("t"));
        });
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        Assert.DoesNotContain(zip.Entries, entry => entry.FullName.Contains("calcChain", StringComparison.OrdinalIgnoreCase));
        foreach (var entry in zip.Entries.Where(entry => entry.FullName.EndsWith(".rels", StringComparison.Ordinal)))
        {
            using var relationships = entry.Open(); var document = XDocument.Load(relationships);
            Assert.DoesNotContain(document.Root!.Elements(), relation => ((string?)relation.Attribute("Type"))?.EndsWith("/calcChain", StringComparison.Ordinal) == true);
        }
    }

    private void ExportSharedGeometry(string label, byte[] source, string json, byte[] output)
    {
        if (Environment.GetEnvironmentVariable("MONOREPORT_XLSX_EVIDENCE") is not { Length: > 0 } directory) return;
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, "shared-" + label + "-template.xlsx"), source);
        File.WriteAllText(Path.Combine(directory, "shared-" + label + "-data.json"), json);
        File.WriteAllBytes(Path.Combine(directory, "shared-" + label + "-generated.xlsx"), output);
    }
}
