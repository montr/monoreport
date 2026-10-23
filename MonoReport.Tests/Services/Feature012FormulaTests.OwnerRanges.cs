using DocumentFormat.OpenXml.Spreadsheet;
using MonoReport.Models;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012FormulaTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OwnerRanges_Body_range_partially_overlapping_parent_header_rejects_incomplete_projection(bool mergedHeader)
    {
        var source = _fixture.Create((workbook, part) =>
        {
            part.Worksheet = new Worksheet(new SheetData(
                new Row(_fixture.Text("A2", "$l(10).Items"), _fixture.Text("B2", "{Name}")) { RowIndex = 2 },
                new Row(_fixture.Text("A3", "$l(40).Items"), _fixture.Text("B3", "{Price}"),
                    GroupFormula("C3", "SUM(B2:B3)")) { RowIndex = 3 }));
            if (mergedHeader) part.Worksheet.Append(new MergeCells(new MergeCell { Reference = "B2:C2" }));
            workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text = "'Report'!$B$2:$C$3";
        });
        var original = source.ToArray();
        GroupRenderResult? result = null;

        // Exercise the real internal pipeline while public nested generation remains guarded.
        // Both B2 (parent header) and B3 (current body) are operands; dropping B2 is unsafe.
        var error = Assert.Throws<TemplateException>(() => result = RenderGroups(source,
            """{"Items":[{"$level":10,"Name":"Parent"},{"$level":40,"Price":11}]}"""));

        Assert.Null(result);
        Assert.Equal(original, source);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, error.Code);
        var issue = Assert.Single(error.Errors);
        Assert.Equal(DiagnosticCodes.XlsxUnsafeFormula, issue.Code);
        Assert.Contains("sheet 'Report'", issue.Location);
        Assert.EndsWith(", C3", issue.Location);
        Assert.Contains("Reference B2:B3", issue.Message);
        Assert.Contains("empty, discarded, disjoint or ambiguous image", issue.Message);
        Assert.Contains("Re-author it inside one compatible owner using a retained contiguous A1 range", issue.Message);
    }
}
