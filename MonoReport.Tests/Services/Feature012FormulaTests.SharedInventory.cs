using System.Collections.Immutable;
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
    public void SharedInventory_Followers_before_master_keep_source_identity_flags_and_ordinary_overrides()
    {
        var bytes = SharedInventoryWorkbook();
        using var input = new MemoryStream(bytes);
        using var package = SpreadsheetDocument.Open(input, false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        using var services = _fixture.Services();
        var context = new XlsxOperationContext(new MergeOptions());
        var source = services.GetRequiredService<XlsxPackageReader>().Read(package, context);
        var xml = source.Sheets[0].Xml.ToString();
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(source, context);
        var group = Assert.Single(plan.SharedFormulas);
        Assert.Equal(7u, group.Index);
        Assert.Equal(new SourceCellId(1, new(4, 4)), group.Master);
        Assert.Equal(new XlsxRectangle(new(3, 2), new(4, 4)), group.Range);
        Assert.Equal("D4+$D$4+$D4+D$4+SUM(D4:E4)+\"B3\"", group.Formula);
        Assert.Equal(new[] { new SourceCellId(1, new(3, 2)), new SourceCellId(1, new(3, 4)), new SourceCellId(1, new(4, 4)) }, group.Members);
        var cells = plan.Sheets[0].Cells;
        Assert.All(cells.Where(cell => cell.SharedFormulaIndex is not null), cell =>
        {
            Assert.Equal(7u, cell.SharedFormulaIndex);
            Assert.NotNull(cell.Formula); // M3.12 normalizes explicit members without mutating source XML.
            Assert.Empty(cell.Tokens);
        });
        var ordinary = cells.Single(cell => cell.Source.Id.Address == new XlsxAddress(3, 3));
        Assert.Null(ordinary.SharedFormulaIndex);
        Assert.Equal("1+1", ordinary.Formula!.Text);
        Assert.DoesNotContain(ordinary.Source.Id, group.Members);
        Assert.Equal("ignored follower text", cells[0].Source.Xml.Element(cells[0].Source.Xml.Name.Namespace + "f")!.Value);
        Assert.Equal(4, context.SourceFormulas);
        Assert.Equal(xml, source.Sheets[0].Xml.ToString());
        Assert.Equal(bytes, input.ToArray());
        if (Environment.GetEnvironmentVariable("MONOREPORT_XLSX_EVIDENCE") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, "shared-inventory-template.xlsx"), bytes);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedInventory_Indices_are_worksheet_local_and_master_lookup_does_not_cross_sheets(bool reverse)
    {
        var bytes = SharedInventoryWorkbook((workbook, _) =>
        {
            var part = workbook.AddNewPart<WorksheetPart>();
            part.Worksheet = new Worksheet(new SheetData(new Row(
                SharedInventoryCell("B3", "1", "B3", 7)) { RowIndex = 3 }));
            var sheets = workbook.Workbook!.Sheets!;
            sheets.Append(new Sheet { Name = "Other", SheetId = 41, Id = workbook.GetIdOfPart(part) });
            workbook.Workbook.DefinedNames!.Append(new DefinedName("Other!$B$3:$D$5") { Name = "_xlnm.Print_Area", LocalSheetId = 1 });
            if (!reverse) return;
            var last = sheets.LastChild!; last.Remove(); sheets.PrependChild(last);
            foreach (var name in workbook.Workbook.DefinedNames.Elements<DefinedName>()) name.LocalSheetId = 1 - name.LocalSheetId!.Value;
        });
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        Assert.Equal(reverse ? new uint[] { 41, 1 } : [1, 41], plan.SharedFormulas.Select(group => group.Master.SheetId));
        Assert.Equal(new XlsxAddress(4, 4), plan.SharedFormulas.Single(group => group.Master.SheetId == 1).Master.Address);
        Assert.Equal("1", plan.SharedFormulas.Single(group => group.Master.SheetId == 41).Formula);
        Assert.All(plan.SharedFormulas, group => Assert.All(group.Members, member => Assert.Equal(group.Master.SheetId, member.SheetId)));

        var orphan = _fixture.RewritePart(bytes, "xl/worksheets/sheet2.xml", xml => xml.Replace("ref=\"B3\"", "", StringComparison.Ordinal));
        AssertSharedInventoryFailure(orphan, "no master", "B3", "Other");
    }

    [Theory]
    [InlineData("missing-master", "no master", "B3")]
    [InlineData("conflicting-masters", "conflicting masters", "D4")]
    [InlineData("empty-master", "no scalar formula text", "D4")]
    [InlineData("whitespace-master", "no scalar formula text", "D4")]
    [InlineData("follower-outside", "outside its master ref", "B3")]
    [InlineData("master-outside", "outside its master ref", "D4")]
    [InlineData("overlap", "ranges overlap", "C4")]
    [InlineData("duplicate-formula", "exactly one formula", "B3")]
    [InlineData("special-metadata", "unsupported shared-formula metadata", "D4")]
    [InlineData("array", "t=shared", "D4")]
    [InlineData("normal", "t=shared", "D4")]
    [InlineData("child-element", "unsupported shared-formula metadata", "D4")]
    [InlineData("invalid-ca", "unsupported shared-formula metadata", "D4")]
    public void SharedInventory_Invalid_or_ambiguous_metadata_rejects_before_output(string mutation, string message, string address)
    {
        var bytes = _fixture.RewritePart(SharedInventoryWorkbook(), "xl/worksheets/sheet1.xml", xml =>
        {
            var document = XDocument.Parse(xml); var ns = document.Root!.Name.Namespace;
            var master = document.Descendants(ns + "c").Single(cell => (string?)cell.Attribute("r") == "D4").Element(ns + "f")!;
            var follower = document.Descendants(ns + "c").First().Element(ns + "f")!;
            switch (mutation)
            {
                case "missing-master": master.Attribute("ref")!.Remove(); break;
                case "conflicting-masters": follower.SetAttributeValue("ref", "B3:D4"); break;
                case "empty-master": master.Value = ""; break;
                case "whitespace-master": master.Value = " \t "; break;
                case "follower-outside": master.SetAttributeValue("ref", "C3:D4"); break;
                case "master-outside": master.SetAttributeValue("ref", "B3:D3"); break;
                case "overlap": master.Parent!.AddBeforeSelf(new XElement(ns + "c", new XAttribute("r", "C4"),
                    new XElement(ns + "f", new XAttribute("t", "shared"), new XAttribute("si", 8), new XAttribute("ref", "C4:D5"), "1"))); break;
                case "duplicate-formula": follower.AddAfterSelf(new XElement(follower)); break;
                case "special-metadata": master.SetAttributeValue("bx", "1"); break;
                case "array": master.SetAttributeValue("t", "array"); break;
                case "normal": master.SetAttributeValue("t", "normal"); break;
                case "child-element": master.Add(new XElement(ns + "x")); break;
                case "invalid-ca": master.SetAttributeValue("ca", "yes"); break;
                default: throw new InvalidOperationException(mutation);
            }
            return document.ToString();
        });
        AssertSharedInventoryFailure(bytes, message, address);
    }

    [Theory]
    [InlineData(null)] [InlineData("")] [InlineData("-1")] [InlineData("4294967296")] [InlineData("seven")]
    public void SharedInventory_Missing_or_invalid_unsigned_index_rejects(string? index)
    {
        var bytes = SharedInventoryWorkbook((_, part) => part.Worksheet!.Descendants<CellFormula>().First()
            .SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("si", "", index ?? "")));
        if (index is null) bytes = _fixture.RewritePart(bytes, "xl/worksheets/sheet1.xml", xml => xml.Replace("si=\"\"", "", StringComparison.Ordinal));
        AssertSharedInventoryFailure(bytes, "valid unsigned si", "B3");
    }

    [Theory]
    [InlineData("")] [InlineData("D4:B3")] [InlineData("B4:D3")] [InlineData("B:B")]
    [InlineData("Report!B3:D4")] [InlineData("B3:D4,F5")] [InlineData("A0:D4")]
    [InlineData("B3:XFE4")] [InlineData("B3:D1048577")] [InlineData("B3:D4suffix")]
    public void SharedInventory_Invalid_master_range_rejects(string range)
    {
        var bytes = SharedInventoryWorkbook((_, part) => part.Worksheet!.Descendants<CellFormula>().Single(formula => formula.Reference is not null).Reference = range);
        AssertSharedInventoryFailure(bytes, "A1", "D4");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SharedInventory_Crop_does_not_hide_the_master_or_an_invalid_follower(bool masterOutside)
    {
        var bytes = _fixture.Create((_, part) =>
        {
            var data = part.Worksheet!.GetFirstChild<SheetData>()!; data.RemoveAllChildren();
            data.Append(new Row(SharedInventoryCell("B2", masterOutside ? "SUM(B3:$B$3)" : "", masterOutside ? "B2:B3" : null, 0)) { RowIndex = 2 },
                new Row(SharedInventoryCell("B3", masterOutside ? "" : "SUM(B3:$B$3)", masterOutside ? null : "B2:B3", 0)) { RowIndex = 3 });
        });
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var group = Assert.Single(plan.SharedFormulas);
        Assert.Equal(new XlsxAddress(masterOutside ? 2 : 3, 2), group.Master.Address);
        Assert.Equal(2, group.Members.Length);
        Assert.Equal(0u, Assert.Single(plan.Sheets[0].Cells).SharedFormulaIndex);
        Assert.Equal(masterOutside ? "SUM(B4:$B$3)" : "SUM(B3:$B$3)", Assert.Single(plan.Sheets[0].Cells).Formula!.Text);
        AssertSharedInventoryFailure(bytes, "requires shared-formula relocation", "B3");
        var invalid = _fixture.RewritePart(bytes, "xl/worksheets/sheet1.xml", xml => xml.Replace("B2:B3", "B3", StringComparison.Ordinal));
        AssertSharedInventoryFailure(invalid, "outside its master ref", "B2");
    }

    [Fact]
    public void SharedInventory_Valid_retained_groups_are_guarded_in_preflight_and_direct_formula_finalization()
    {
        var bytes = SharedInventoryWorkbook();
        AssertSharedInventoryFailure(bytes, "requires shared-formula relocation", "B3", errorCount: 3);
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        var output = services.GetRequiredService<XlsxGeometryPlanner>().Plan(plan, ImmutableDictionary<SourceCellId, XlsxPreparedValue>.Empty, context);
        var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxFormulaPlanner>().Finalize(plan, output, context));
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
        Assert.All(error.Errors, issue => Assert.Contains("ordinary per-cell scalar A1", issue.Message));
        Assert.Equal(3, error.Errors.Count);
    }

    [Fact]
    public void SharedInventory_Distinct_nonoverlapping_groups_on_one_sheet_keep_their_own_members()
    {
        var bytes = SharedInventoryWorkbook((_, part) => part.Worksheet!.GetFirstChild<SheetData>()!.Append(
            new Row(SharedInventoryCell("B5", "42", "B5", 8)) { RowIndex = 5 }));
        using var services = _fixture.Services();
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(package, context), context);
        Assert.Equal(new uint[] { 7, 8 }, plan.SharedFormulas.Select(group => group.Index));
        var single = plan.SharedFormulas[1];
        Assert.Equal("42", single.Formula);
        Assert.Equal(new SourceCellId(1, new(5, 2)), Assert.Single(single.Members));
        Assert.Equal(single.Master, single.Members[0]);
        Assert.Equal(new XlsxRectangle(new(5, 2), new(5, 2)), single.Range);
        Assert.Equal(8u, plan.Sheets[0].Cells.Single(cell => cell.Source.Id == single.Master).SharedFormulaIndex);
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(3)]
    public void SharedInventory_Repeated_groups_reject_before_data_counts_can_mask_unnormalized_formulas(int count)
    {
        var bytes = SharedInventoryWorkbook((_, part) =>
        {
            foreach (var row in part.Worksheet!.Descendants<Row>())
                row.PrependChild(_fixture.Text("A" + row.RowIndex!.Value, "$l(10).Items"));
        });
        var json = "{\"Items\":[" + string.Join(',', Enumerable.Repeat("{\"$level\":10}", count)) + "]}";
        AssertSharedInventoryFailure(bytes, "requires shared-formula relocation", "B3", errorCount: 3, json: json);
    }

    [Fact]
    public void SharedInventory_Wholly_cropped_groups_do_not_affect_ordinary_formula_output_or_allocate_their_range()
    {
        var bytes = _fixture.Create((_, part) =>
        {
            var data = part.Worksheet!.GetFirstChild<SheetData>()!; data.RemoveAllChildren();
            data.Append(new Row(SharedInventoryCell("A1", "1", "A1:XFD2", uint.MaxValue)) { RowIndex = 1 },
                new Row(SharedInventoryCell("A2", "", null, uint.MaxValue)) { RowIndex = 2 },
                new Row(new Cell { CellReference = "B3", CellFormula = new CellFormula("C3+$D$3+\"A1\""), CellValue = new CellValue("9999") },
                    new Cell { CellReference = "C3", CellValue = new CellValue("3") }, new Cell { CellReference = "D3", CellValue = new CellValue("4") }) { RowIndex = 3 });
        });
        using var services = _fixture.Services();
        using var input = new MemoryStream(bytes);
        Assert.True(services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
        using var original = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var context = new XlsxOperationContext(new MergeOptions());
        var plan = services.GetRequiredService<XlsxTemplatePlanner>().Plan(services.GetRequiredService<XlsxPackageReader>().Read(original, context), context);
        Assert.Equal(2, Assert.Single(plan.SharedFormulas).Members.Length);
        Assert.Equal(uint.MaxValue, plan.SharedFormulas[0].Index);
        var result = _fixture.Merge(bytes, "{}");
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(output));
        var cells = output.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToArray();
        Assert.Equal(new[] { "A1", "B1", "C1" }, cells.Select(cell => cell.CellReference!.Value));
        Assert.Equal("B1+$C$1+\"A1\"", cells[0].CellFormula!.Text);
        Assert.Null(cells[0].CellValue);
        Assert.Empty(cells[0].CellFormula!.GetAttributes());
        Assert.Equal(new[] { "3", "4" }, cells.Skip(1).Select(cell => cell.CellValue!.Text));
        Assert.True(output.WorkbookPart.Workbook!.CalculationProperties!.ForceFullCalculation!.Value);
        Assert.Empty(result.Warnings);
        if (Environment.GetEnvironmentVariable("MONOREPORT_XLSX_EVIDENCE") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, "shared-cropped-template.xlsx"), bytes);
            File.WriteAllBytes(Path.Combine(directory, "shared-cropped-generated.xlsx"), result.Document);
        }
    }

    [Theory]
    [InlineData(99)] [InlineData(100)] [InlineData(101)]
    public void SharedInventory_Malformed_groups_keep_source_order_and_the_existing_error_cap(int count)
    {
        var bytes = _fixture.Create((_, part) =>
        {
            var data = part.Worksheet!.GetFirstChild<SheetData>()!; data.RemoveAllChildren();
            data.Append(new Row(Enumerable.Range(0, count).Select(index => SharedInventoryCell(new XlsxAddress(3, index + 2).ToString(), "", null, (uint)index))) { RowIndex = 3 });
        });
        using var services = _fixture.Services();
        using var input = new MemoryStream(bytes);
        var issues = services.GetRequiredService<TemplateAnalyzer>().Analyze(input).Errors;
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(bytes, "{}"));
        Assert.Equal(Math.Min(100, count), issues.Count);
        Assert.Equal(count >= 100, error.LimitReached);
        Assert.Equal(issues.Select(issue => issue.Message), error.Errors.Select(issue => issue.Message));
        Assert.Equal(issues.Select(issue => issue.Location), error.Errors.Select(issue => issue.Location));
        for (var index = 0; index < issues.Count; index++)
        {
            Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issues[index].Code);
            Assert.EndsWith(", " + new XlsxAddress(3, index + 2), issues[index].Location);
            Assert.Contains("group " + index + " has no master", issues[index].Message);
        }
        Assert.Equal(bytes, input.ToArray());
    }

    private byte[] SharedInventoryWorkbook(Action<WorkbookPart, WorksheetPart>? customize = null) => _fixture.Create((workbook, part) =>
    {
        var data = part.Worksheet!.GetFirstChild<SheetData>()!; data.RemoveAllChildren();
        var master = SharedInventoryCell("D4", "D4+$D$4+$D4+D$4+SUM(D4:E4)+\"B3\"", "B3:D4", 7);
        master.CellFormula!.CalculateCell = true;
        master.CellFormula.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
        data.Append(new Row(SharedInventoryCell("B3", "ignored follower text", null, 7),
                new Cell { CellReference = "C3", CellFormula = new CellFormula("1+1") }, SharedInventoryCell("D3", "", null, 7)) { RowIndex = 3 },
            new Row(master) { RowIndex = 4 });
        customize?.Invoke(workbook, part);
    });

    private Cell SharedInventoryCell(string address, string formula, string? range, uint index)
    {
        var element = new CellFormula(formula) { FormulaType = CellFormulaValues.Shared, SharedIndex = index };
        if (range is not null) element.Reference = range;
        return new Cell { CellReference = address, CellFormula = element, CellValue = new CellValue("9999") };
    }

    private void AssertSharedInventoryFailure(byte[] bytes, string message, string address, string sheet = "Report", int errorCount = 1, string json = "{}")
    {
        var unchanged = bytes.ToArray();
        using var services = _fixture.Services(); using var input = new MemoryStream(bytes);
        var analysis = services.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(analysis.Valid);
        Assert.Equal(errorCount, analysis.Errors.Count);
        var issue = analysis.Errors[0];
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
        Assert.Contains("sheet '" + sheet + "', " + address, issue.Location);
        Assert.Contains(message, issue.Message);
        Assert.Contains("A1", issue.Message);
        MergeResult? result = null;
        var error = Assert.Throws<TemplateException>(() => result = _fixture.Merge(bytes, json));
        Assert.Null(result);
        Assert.Equal(analysis.Errors.Select(item => item.Message), error.Errors.Select(item => item.Message));
        Assert.Equal(analysis.Errors.Select(item => item.Location), error.Errors.Select(item => item.Location));
        Assert.Equal(unchanged, bytes);
        Assert.Equal(unchanged, input.ToArray());
    }
}
