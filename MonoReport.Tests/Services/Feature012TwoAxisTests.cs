using System.Collections.Immutable;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services.Xlsx;
using DocumentFormat.OpenXml.Validation;
using System.IO.Compression;

namespace MonoReport.Tests.Services;

public sealed class Feature012TwoAxisTests
{
    private readonly XlsxFixture _fixture = new();
    private readonly XlsxTwoAxisFixture _twoAxis = new();

    [Fact]
    public void Matrix_Public_analysis_and_header_parent_root_scopes_follow_the_same_prepared_plan()
    {
        var bytes = _twoAxis.Create((_, part) => part.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "E8")
            .InlineString!.Text!.Text = "{Name}/{../Company}/{$.Title}");
        using var services = _fixture.Services(); using var input = new MemoryStream(bytes);
        var analysis = services.GetRequiredService<MonoReport.Services.TemplateAnalyzer>().Analyze(input);
        Assert.True(analysis.Valid);
        Assert.Equal(new[] { "Bids", "Criteria", "Title", "Items" }, analysis.Requirements.Select(item => item.Name));
        Assert.Contains(analysis.Occurrences, item => item.Path == "$.Bids[].Company" && item.Expression == "{../Company}");
        Assert.DoesNotContain(analysis.Occurrences, item => item.Expression is "$h(1)" or "$header(1)" or "$h(2)");
        var result = _fixture.Merge(bytes, _twoAxis.Data(1, 1)[..^1] + ",\"Title\":\"Root\"}");
        Assert.Equal("Criterion 0/Alpha/Root", _fixture.ReadCellText(result.Document, "D2"));
        Assert.Equal("Criterion 0/Beta/Root", _fixture.ReadCellText(result.Document, "F2"));
    }

    [Fact]
    public void Matrix_Disjoint_selection_ranges_reject_with_a_remedy_instead_of_null_geometry()
    {
        var bytes = _twoAxis.Create((_, part) => part.Worksheet!.PrependChild(new SheetViews(new SheetView(new Selection { SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "E10:E10" } }) { WorkbookViewId = 0 })));
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(bytes, _twoAxis.Data()));
        Assert.Equal(DiagnosticCodes.XlsxUnsupportedFeature, error.Code); Assert.Contains("selection", error.Errors[0].Message);
    }

    [Fact]
    public void Empty_Removed_leaf_formula_disappears_but_a_surviving_empty_reference_rejects()
    {
        var bytes = _twoAxis.Create((_, part) =>
        {
            var cell = part.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "E10");
            cell.RemoveAllChildren(); cell.DataType = null; cell.CellFormula = new CellFormula("A1");
        });
        var result = _fixture.Merge(bytes, _twoAxis.Data(0));
        using (var document = SpreadsheetDocument.Open(new MemoryStream(result.Document), false))
            Assert.Empty(document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<CellFormula>());
        bytes = _twoAxis.Create((_, part) => part.Worksheet!.Descendants<Row>().Single(row => row.RowIndex!.Value == 11)
            .Append(new Cell { CellReference = "F11", CellFormula = new CellFormula("SUM(E10:E10)") }));
        var error = Assert.Throws<TemplateException>(() => _fixture.Merge(bytes, _twoAxis.Data(0)));
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code); Assert.Contains("F11", error.Location); Assert.Contains("empty", error.Errors[0].Message);
    }

    [Fact]
    public void Matrix_Sibling_leaf_blocks_reuse_tuple_indices_and_bind_corresponding_references()
    {
        var bytes = _twoAxis.Create((workbook, part) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$7:$H$12";
            var rows = part.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
            rows[1].Append(_fixture.Text("F1", "$h(1)"), _fixture.Text("G1", "$h(1)"));
            rows[2].Append(_fixture.Text("G2", "$h(2)"));
            rows[8].Append(_fixture.Text("G8", "{Name}")); rows[10].Append(_fixture.Text("G10", "{Price}"));
            rows[11].Append(new Cell { CellReference = "G11", CellFormula = new CellFormula("E10") });
            part.Worksheet.GetFirstChild<MergeCells>()!.Elements<MergeCell>().Single().Reference = "D7:G7";
        });
        var result = _fixture.Merge(bytes, _twoAxis.Data(3, 1));
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(output));
        var cells = output.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal(new[] { "100", "101", "102" }, new[] { "H4", "I4", "J4" }.Select(address => cells[address].CellValue!.Text));
        Assert.Equal(new[] { "D4", "E4", "F4", "L4", "M4", "N4" }, new[] { "H5", "I5", "J5", "P5", "Q5", "R5" }.Select(address => cells[address].CellFormula!.Text));
    }

    [Fact]
    public void Matrix_Parent_heading_inside_child_columns_has_one_owner_and_one_content_copy()
    {
        var bytes = _twoAxis.Create((_, part) =>
        {
            var row = part.Worksheet!.Descendants<Row>().Single(row => row.RowIndex!.Value == 7);
            row.Elements<Cell>().Single(cell => cell.CellReference!.Value == "D7").Remove();
            var anchor = row.Elements<Cell>().Single(cell => cell.CellReference!.Value == "E7");
            anchor.DataType = CellValues.InlineString; anchor.InlineString = new InlineString(new Text("{Company}"));
            part.Worksheet.RemoveAllChildren<MergeCells>();
        });
        var result = _fixture.Merge(bytes, _twoAxis.Data());
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Equal(2, output.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Count(cell => cell.InlineString?.InnerText is "Alpha" or "Beta"));
        Assert.Equal("Alpha", _fixture.ReadCellText(result.Document, "D1")); Assert.Equal("Beta", _fixture.ReadCellText(result.Document, "H1"));
    }

    [Theory]
    [InlineData(0, 2, 2)] [InlineData(1, 1, 2)] [InlineData(3, 3, 2)]
    [InlineData(7, 2, 2)] [InlineData(3, 0, 2)] [InlineData(3, 2, 0)]
    public void Matrix_Public_slice_has_exact_geometry_native_values_and_empty_semantics(int criteria, int rows, int companies)
    {
        var result = _fixture.Merge(_twoAxis.Create(), _twoAxis.Data(criteria, rows, companies));
        using var document = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var sheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        var width = 3 + companies * (1 + criteria);
        Assert.Equal(new XlsxAddress(5 + rows, width).ToString(), sheet.GetFirstChild<SheetDimension>()!.Reference!.Value!.Split(':')[1]);
        var cells = sheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal(criteria == 0 ? 0 : companies, sheet.Descendants<MergeCell>().Count());
        for (var row = 0; row < rows; row++)
        {
            Assert.Equal("Row " + row, _fixture.ReadCellText(result.Document, "A" + (row + 4)));
            Assert.Equal((row + 1).ToString(), cells["B" + (row + 4)].CellValue!.Text);
            for (var company = 0; company < companies; company++)
                for (var criterion = 0; criterion < criteria; criterion++)
                {
                    var address = new XlsxAddress(row + 4, 4 + company * (1 + criteria) + criterion).ToString();
                    Assert.Equal(((row + 1) * 100 + company * criteria + criterion).ToString(), cells[address].CellValue!.Text);
                    Assert.Equal(CellValues.Number, cells[address].DataType!.Value);
                }
        }
        Assert.Equal("tail", _fixture.ReadCellText(result.Document, "A" + (rows + 5)));
        Assert.Equal(companies, sheet.Descendants<Cell>().Count(cell => cell.InlineString?.InnerText is "Alpha" or "Beta"));
    }

    [Fact]
    public void Matrix_Vertical_only_empty_report_and_static_scalar_formula_remain_valid()
    {
        var bytes = _fixture.Create((workbook, part) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$3:$D$3";
            part.Worksheet!.Descendants<Row>().Single().PrependChild(_fixture.Text("A3", "$l(10).Items"));
        });
        var empty = _fixture.Merge(bytes, "{\"Items\":[]}");
        using (var document = SpreadsheetDocument.Open(new MemoryStream(empty.Document), false))
        {
            Assert.Empty(document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>());
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        bytes = _fixture.Create((_, part) =>
        {
            var cell = part.Worksheet!.Descendants<Cell>().Single(); cell.RemoveAllChildren(); cell.DataType = null;
            cell.CellFormula = new CellFormula("SUM(B4:B5)"); cell.CellValue = new CellValue("9999");
        });
        var result = _fixture.Merge(bytes);
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        var formula = Assert.Single(output.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>());
        Assert.Equal("SUM(A2:A3)", formula.CellFormula!.Text); Assert.Null(formula.CellValue);
    }

    [Fact]
    public void Matrix_Two_row_block_repeats_as_a_unit_with_one_logical_tuple_set()
    {
        var bytes = _twoAxis.Create((_, part) =>
        {
            var row = part.Worksheet!.Descendants<Row>().Single(item => item.RowIndex!.Value == 11);
            row.RemoveAllChildren(); row.Append(_fixture.Text("A11", "$level(10).Items"), _fixture.Text("B11", "{Name}"),
                new Cell { CellReference = "E11", CellFormula = new CellFormula("E10*2") });
        });
        var result = _fixture.Merge(bytes, _twoAxis.Data(1, 2));
        Assert.Equal("Row 0", _fixture.ReadCellText(result.Document, "A4")); Assert.Equal("Row 0", _fixture.ReadCellText(result.Document, "A5"));
        Assert.Equal("Row 1", _fixture.ReadCellText(result.Document, "A6")); Assert.Equal("Row 1", _fixture.ReadCellText(result.Document, "A7"));
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        var cells = output.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("D4*2", cells["D5"].CellFormula!.Text); Assert.Equal("F6*2", cells["F7"].CellFormula!.Text);
    }

    [Fact]
    public void Matrix_Independent_table_outside_the_horizontal_strip_uses_one_record_per_row()
    {
        var bytes = _twoAxis.Create((_, part) =>
        {
            var row = part.Worksheet!.Descendants<Row>().Single(item => item.RowIndex!.Value == 12);
            row.RemoveAllChildren(); row.Append(_fixture.Text("A12", "$l(10).Other"), _fixture.Text("B12", "{Name}"));
        });
        using var original = JsonDocument.Parse(_twoAxis.Data(3, 1));
        var json = original.RootElement.GetRawText()[..^1] + ",\"Other\":[{\"$level\":10,\"Name\":\"independent 1\"},{\"$level\":10,\"Name\":\"independent 2\"}]}";
        var result = _fixture.Merge(bytes, json);
        Assert.Equal("independent 1", _fixture.ReadCellText(result.Document, "A6"));
        Assert.Equal("independent 2", _fixture.ReadCellText(result.Document, "A7"));
    }

    [Fact]
    public void Boundaries_Same_markers_on_both_sides_of_zero_height_separator_are_independent()
    {
        var bytes = _twoAxis.Create((workbook, worksheet) =>
        {
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$7:$F$13";
            var rows = worksheet.Worksheet!.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
            rows[11].Height = 0; rows[11].CustomHeight = true;
            rows[11].Append(new Cell { CellReference = "C11", CellFormula = new CellFormula("SUBTOTAL(9,C10:C10)") });
            rows[12].PrependChild(_fixture.Text("A12", "$level(10).Items"));
            rows[12].Elements<Cell>().Single(cell => cell.CellReference!.Value == "B12").InlineString!.Text!.Text = "{Name}";
            rows[12].Append(_fixture.Text("D12", "{Name}"), _fixture.Text("E12", "{Price}"), new Cell { CellReference = "F12", CellFormula = new CellFormula("SUM(D12:E12)") });
            worksheet.Worksheet.GetFirstChild<SheetData>()!.Append(new Row(new Cell { CellReference = "F13", CellFormula = new CellFormula("SUBTOTAL(9,C10:C12)") }) { RowIndex = 13 });
        });
        var result = _fixture.Merge(bytes, _twoAxis.Data());
        Assert.Equal("Row 0", _fixture.ReadCellText(result.Document, "A4")); Assert.Equal("Row 1", _fixture.ReadCellText(result.Document, "A5"));
        Assert.Equal("Row 0", _fixture.ReadCellText(result.Document, "A7")); Assert.Equal("Row 1", _fixture.ReadCellText(result.Document, "A8"));
        using var document = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        var sheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        var separator = Assert.Single(sheet.Descendants<Row>(), row => row.Height?.Value == 0);
        Assert.Equal(6U, separator.RowIndex!.Value); Assert.Null(separator.Hidden);
        var cells = sheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal("SUBTOTAL(9,B4:B5)", cells["B6"].CellFormula!.Text);
        Assert.Equal("SUM(C7:J7)", cells["K7"].CellFormula!.Text); Assert.Equal("SUM(C8:J8)", cells["K8"].CellFormula!.Text);
        Assert.Equal("SUBTOTAL(9,B4:B8)", cells["K9"].CellFormula!.Text);
        Assert.Equal("A1:K9", sheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
    }

    [Fact]
    public void Crop_Removes_control_and_shared_sentinel_content_and_preserves_nonanchor_styles()
    {
        const string sentinel = "DISPOSABLE_M2_SHARED_SENTINEL";
        var bytes = _twoAxis.Create((workbook, worksheet) =>
        {
            workbook.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable(new SharedStringItem(new Text(sentinel)));
            var data = worksheet.Worksheet!.GetFirstChild<SheetData>()!;
            data.Append(new Row(new Cell { CellReference = "G99", DataType = CellValues.SharedString, CellValue = new CellValue("0") }) { RowIndex = 99 });
            workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()),
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(new Border()), new CellStyleFormats(new CellFormat()), new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 14 }));
            worksheet.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "E7").StyleIndex = 1;
            data.Elements<Row>().Single(row => row.RowIndex!.Value == 11).Append(new Cell { CellReference = "F11", CellFormula = new CellFormula("SUM(D10:E10)") });
        });
        var result = _fixture.Merge(bytes, _twoAxis.Data());
        using var output = SpreadsheetDocument.Open(new MemoryStream(result.Document), false);
        Assert.Empty(new OpenXmlValidator().Validate(output)); Assert.Null(output.WorkbookPart!.SharedStringTablePart);
        var sheet = output.WorkbookPart.WorksheetParts.Single().Worksheet!;
        Assert.Equal("A1:K7", sheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        Assert.Equal("'Report'!$A$1:$K$7", output.WorkbookPart.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text);
        var cells = sheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.All(new[] { "D1", "E1", "F1", "H1", "I1", "J1" }, address => Assert.Equal(1U, cells[address].StyleIndex!.Value));
        Assert.Equal("SUM(C4:J5)", cells["K6"].CellFormula!.Text);
        using var zip = new ZipArchive(new MemoryStream(result.Document));
        foreach (var entry in zip.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal)))
        {
            using var reader = new StreamReader(entry.Open()); var xml = reader.ReadToEnd();
            Assert.DoesNotContain(sentinel, xml); Assert.DoesNotContain("$header(", xml); Assert.DoesNotContain("$h(", xml); Assert.DoesNotContain("$l(", xml);
        }
    }

    [Theory]
    [InlineData(0, 2, "C1", "D1")]
    [InlineData(3, 0, null, null)]
    public void Empty_Collapse_owned_merges_keep_descriptors_and_remove_orphan_headings(int criteria, int companies, string? alpha, string? beta)
    {
        var bytes = _twoAxis.Create(); var source = _twoAxis.Plan(bytes); var sheet = source.Sheets.Single();
        using var json = JsonDocument.Parse(_twoAxis.Data(criteria, 2, companies)); using var services = _fixture.Services();
        var context = new XlsxOperationContext(new MergeOptions()); var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        var geometry = services.GetRequiredService<XlsxGeometryPlanner>(); var data = planner.BindSheet(sheet, json.RootElement, context);
        var positions = geometry.ExpansionGeometry(sheet, services.GetRequiredService<XlsxTemplatePlanner>().ScheduleSheet(sheet, data, context), geometry.HorizontalSpans(sheet, data)); var output = geometry.EmitCells(sheet, positions);
        geometry.ProjectMerges(sheet, output); output = planner.BindInstances(sheet, output, data, context);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Assert.Empty(output.Settings.Descendants(ns + "mergeCell"));
        var values = output.Rows.SelectMany(row => row.Cells).Where(cell => cell.Value?.Value is not null).ToDictionary(cell => cell.Output.ToString(), cell => cell.Value!.Value);
        Assert.Equal("Row 0", values["A4"]); Assert.Equal("Row 1", values["A5"]);
        if (alpha is not null) { Assert.Equal("Alpha", values[alpha]); Assert.Equal("Beta", values[beta!]); }
        else { Assert.DoesNotContain("Alpha", values.Values); Assert.DoesNotContain("Beta", values.Values); }
    }

    [Theory]
    [InlineData("C7:E7", false, "owner")]
    [InlineData("D9:E10", false, "vertical owner")]
    [InlineData("D7:E7", true, "nonanchor")]
    [InlineData("A7:D7", false, "PrintArea")]
    public void Empty_Reject_unsafe_merge_ownership_and_nonanchor_data_loss(string range, bool content, string remedy)
    {
        var bytes = _twoAxis.Create((_, worksheet) =>
        {
            worksheet.Worksheet!.GetFirstChild<MergeCells>()!.Elements<MergeCell>().Single().Reference = range;
            if (content)
            {
                var cell = worksheet.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "E7");
                cell.DataType = CellValues.InlineString; cell.InlineString = new InlineString(new Text("must not lose"));
            }
            else if (range == "C7:E7" || range == "D9:E10")
                foreach (var cell in worksheet.Worksheet.Descendants<Cell>().Where(cell => range == "C7:E7"
                    ? cell.CellReference!.Value == "D7" : cell.CellReference!.Value is "D10" or "E10"))
                { cell.RemoveAllChildren(); cell.DataType = null; }
        });
        var sheet = _twoAxis.Plan(bytes).Sheets.Single();
        using var services = _fixture.Services(); using var json = JsonDocument.Parse(_twoAxis.Data());
        var context = new XlsxOperationContext(new MergeOptions()); var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
        var data = services.GetRequiredService<XlsxTemplatePlanner>().BindSheet(sheet, json.RootElement, context);
        var output = geometry.EmitCells(sheet, geometry.ExpansionGeometry(sheet, services.GetRequiredService<XlsxTemplatePlanner>().ScheduleSheet(sheet, data, context), geometry.HorizontalSpans(sheet, data)));
        Assert.Contains(remedy, Assert.Throws<TemplateException>(() => geometry.ProjectMerges(sheet, output)).Message);
    }

    [Fact]
    public void Cells_Emit_tuple_values_first_record_ordinary_fields_and_single_parent_headings()
    {
        var bytes = _twoAxis.Create((_, worksheet) => worksheet.Worksheet!.RemoveAllChildren<MergeCells>());
        var source = _twoAxis.Plan(bytes); var sheet = source.Sheets.Single();
        using var json = JsonDocument.Parse(_twoAxis.Data()); using var services = _fixture.Services();
        var context = new XlsxOperationContext(new MergeOptions());
        var planner = services.GetRequiredService<XlsxTemplatePlanner>(); var geometry = services.GetRequiredService<XlsxGeometryPlanner>();
        var data = planner.BindSheet(sheet, json.RootElement, context);
        var positions = geometry.ExpansionGeometry(sheet, services.GetRequiredService<XlsxTemplatePlanner>().ScheduleSheet(sheet, data, context), geometry.HorizontalSpans(sheet, data));
        var output = planner.BindInstances(sheet, geometry.EmitCells(sheet, positions), data, context);
        using var buffer = new MemoryStream(); buffer.Write(bytes); buffer.Position = 0;
        using (var document = SpreadsheetDocument.Open(buffer, true))
            services.GetRequiredService<XlsxDocumentRenderer>().Render(document, new([output], source.Source.Names), context);
        var result = buffer.ToArray();
        Assert.Equal("Alpha", _fixture.ReadCellText(result, "C1")); Assert.Equal("Beta", _fixture.ReadCellText(result, "G1"));
        Assert.Equal("Row 0", _fixture.ReadCellText(result, "A4")); Assert.Equal("Row 0", _fixture.ReadCellText(result, "G4"));
        using var package = SpreadsheetDocument.Open(new MemoryStream(result), false);
        var worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
        var cells = worksheetPart.Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal(new[] { "100", "101", "102", "103", "104", "105" }, new[] { "D4", "E4", "F4", "H4", "I4", "J4" }.Select(address => cells[address].CellValue!.Text));
        Assert.All(new[] { "D4", "E4", "F4", "H4", "I4", "J4" }, address => Assert.Equal(CellValues.Number, cells[address].DataType!.Value));
        Assert.Equal(18, worksheetPart.Worksheet.Descendants<Row>().Single(row => row.RowIndex!.Value == 5).Height!.Value);
        Assert.Equal(new uint[] { 3, 7 }, worksheetPart.Worksheet.Descendants<Column>().Select(column => column.Min!.Value));
        Assert.All(worksheetPart.Worksheet.Descendants<Column>(), column => { Assert.Equal(0, column.Width!.Value); Assert.Null(column.Hidden); });
    }
}
