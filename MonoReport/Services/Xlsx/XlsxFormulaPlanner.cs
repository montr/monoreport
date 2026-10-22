using MonoReport.Models;
using MonoReport.Models.Xlsx;
using System.Collections.Immutable;

namespace MonoReport.Services.Xlsx;

internal sealed class XlsxFormulaPlanner(XlsxFormulaReferences references, XlsxGeometryPlanner geometry)
{
    private readonly XlsxFormulaReferences _references = references ?? throw new ArgumentNullException(nameof(references));
    private readonly XlsxGeometryPlanner _geometry = geometry ?? throw new ArgumentNullException(nameof(geometry));

    internal void ValidateStaticReferences(XlsxTemplatePlan template, XlsxOperationContext context)
    {
        foreach (var sheet in template.Sheets)
            foreach (var cell in sheet.Cells.Where(cell => cell.Formula is not null))
            {
                if (context.LimitReached) break;
                var address = cell.Source.Id.Address;
                // Only these cells survive independently of JSON counts; discarded repeated formulas are not evaluated.
                if (sheet.Regions.Any(region => region.FirstRow <= address.Row && address.Row <= region.LastRow) ||
                    sheet.Horizontal is { } horizontal && horizontal.FirstColumn <= address.Column && address.Column <= horizontal.LastColumn) continue;
                foreach (var reference in cell.Formula!.References.Reverse())
                {
                    if (context.LimitReached) break;
                    if (!sheet.Source.PrintArea.Contains(reference.First.Address) || !sheet.Source.PrintArea.Contains(reference.Last.Address))
                    {
                        context.Report(UnsafeImage(cell.Location, cell.Formula, reference));
                        break; // As in formula finalization, report one failing operand per source formula.
                    }
                }
            }
        context.ThrowIfErrors();
    }

    internal XlsxWorkbookPlan Finalize(XlsxTemplatePlan template, XlsxWorkbookPlan workbook, XlsxOperationContext context)
    {
        var parsed = template.Sheets.SelectMany(sheet => sheet.Cells).Where(cell => cell.Formula is not null).ToDictionary(cell => cell.Source.Id, cell => cell.Formula!);
        var sheets = ImmutableArray.CreateBuilder<XlsxSheetOutput>();
        foreach (var sheet in workbook.Sheets)
        {
            var groups = GroupBlocks(template.Sheets.Single(source => source.Source.Id == sheet.Geometry.Source.Id));
            var formulas = new Dictionary<CellInstanceId, string>();
            foreach (var cell in sheet.Rows.SelectMany(row => row.Cells).OrderBy(cell => cell.Source.Id.Address.Row)
                .ThenBy(cell => cell.Source.Id.Address.Column).ThenBy(cell => cell.Output.Row).ThenBy(cell => cell.Output.Column))
            {
                if (context.LimitReached) break;
                if (cell.Suppressed || !parsed.TryGetValue(cell.Source.Id, out var formula)) continue;
                try { formulas.Add(cell.Id, BindFormula(formula, cell, sheet.Geometry,
                    groups.Any(block => block.FirstRow <= cell.Source.Id.Address.Row && cell.Source.Id.Address.Row <= block.LastRow))); }
                catch (TemplateException error) { context.Report(error); }
            }
            sheets.Add(sheet with { Rows = sheet.Rows.Select(row => row with { Cells = row.Cells.Select(cell =>
                cell with { Formula = formulas.GetValueOrDefault(cell.Id) }).ToImmutableArray() }).ToImmutableArray() });
            if (context.LimitReached) break;
        }
        context.ThrowIfErrors(); return workbook with { Sheets = sheets.ToImmutable() };
    }

    private static List<XlsxVerticalBlock> GroupBlocks(XlsxSheetTemplate sheet)
    {
        var blocks = new List<XlsxVerticalBlock>();
        var pending = new Stack<XlsxVerticalNode>(sheet.Regions.Select(region => region.Root));
        while (pending.TryPop(out var node))
        {
            if (node.Header is { } header) blocks.Add(header);
            if (node.Footer is { } footer) blocks.Add(footer);
            foreach (var child in node.Children) pending.Push(child);
        }
        return blocks;
    }

    internal string BindFormula(XlsxParsedFormula formula, XlsxCellPlacement placement, XlsxSheetGeometry geometry,
        bool groupFormula = false)
    {
        var location = geometry.Source.Part + ", sheet '" + geometry.Source.Name + "', " + placement.Source.Id.Address;
        return _references.Rewrite(formula, reference =>
        {
            if (reference.Prefix.Length != 0) throw Unsafe(location, "Sheet-qualified operands require M3; use a local retained A1 operand in this milestone.");
            // Group range ownership is structural, including when dollar flags suppress authored copying.
            // Scalars and body/leaf formulas retain their exact-owner/absolute-anchor rules.
            var descendants = groupFormula && reference.IsRange && placement.Id.Vertical is not null;
            var vertical = !descendants && (reference.First.RowAbsolute || reference.Last.RowAbsolute) ? null : placement.Id.Vertical;
            var horizontal = reference.First.ColumnAbsolute || reference.Last.ColumnAbsolute ? null : placement.Id.Horizontal;
            var range = new XlsxRectangle(reference.First.Address, reference.Last.Address);
            var image = _geometry.ProjectReference(range, geometry, vertical, horizontal, descendants);
            if (geometry.Expanded && reference.IsRange &&
                (reference.First.RowAbsolute != reference.Last.RowAbsolute && placement.Id.Vertical is not null ||
                 reference.First.ColumnAbsolute != reference.Last.ColumnAbsolute && placement.Id.Horizontal is not null))
            {
                var scoped = _geometry.ProjectReference(range, geometry,
                    reference.First.RowAbsolute != reference.Last.RowAbsolute ? placement.Id.Vertical : vertical,
                    reference.First.ColumnAbsolute != reference.Last.ColumnAbsolute ? placement.Id.Horizontal : horizontal, descendants);
                if (scoped.Kind != image.Kind || scoped.Rectangle != image.Rectangle)
                    throw Unsafe(location, "Mixed endpoint copy flags select incompatible repeated owners. Re-author the A1 range with consistent flags or uniquely retained endpoints.");
            }
            if (!image.Complete || image.Kind is XlsxRangeImageKind.Empty or XlsxRangeImageKind.Disjoint or XlsxRangeImageKind.Ambiguous ||
                !reference.IsRange && image.Kind != XlsxRangeImageKind.Cell)
                throw UnsafeImage(location, formula, reference);
            var result = image.Rectangle!.Value;
            return (reference.First with { Address = result.First }).ToString() +
                (reference.IsRange ? ":" + (reference.Last with { Address = result.Last }) : string.Empty);
        });
    }

    private static TemplateException Unsafe(string location, string remedy) => new StructuralDiagnostic(DiagnosticCodes.XlsxUnsafeFormula,
        remedy, Location: location, Explanation: remedy).ToException();

    private static TemplateException UnsafeImage(string location, XlsxParsedFormula formula, XlsxFormulaReference reference) =>
        Unsafe(location, "Reference " + formula.Text.Substring(reference.Start, reference.Length) +
            " has an empty, discarded, disjoint or ambiguous image. Re-author it inside one compatible owner using a retained contiguous A1 range; scalar operands must identify one cell.");
}
