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
        RequireScalarFormulas(template, context);
        var sheets = template.Sheets.ToDictionary(sheet => sheet.Source.Id);
        foreach (var name in template.DefinedNames)
        {
            if (context.LimitReached) break;
            var reference = name.Formula.References[0];
            var target = sheets[reference.TargetSheetId!.Value];
            if (!target.Source.PrintArea.Contains(reference.First.Address) || !target.Source.PrintArea.Contains(reference.Last.Address))
                context.Report(UnsafeImage(name.Location, name.Formula, reference));
        }
        context.ThrowIfErrors();
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
                    if (reference.NameIndex is not null) continue; // Definitions have independent, fixed workbook semantics.
                    var target = sheets[reference.TargetSheetId ?? sheet.Source.Id];
                    if (!target.Source.PrintArea.Contains(reference.First.Address) || !target.Source.PrintArea.Contains(reference.Last.Address))
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
        RequireScalarFormulas(template, context);
        var parsed = template.Sheets.SelectMany(sheet => sheet.Cells).Where(cell => cell.Formula is not null).ToDictionary(cell => cell.Source.Id, cell => cell.Formula!);
        var geometries = workbook.Sheets.ToDictionary(sheet => sheet.Geometry.Source.Id, sheet => sheet.Geometry);
        var names = workbook.Names.Select(name => new System.Xml.Linq.XElement(name)).ToArray();
        foreach (var name in template.DefinedNames)
        {
            if (context.LimitReached) break;
            try
            {
                names[name.Index].Value = Relocate(name.Formula, geometries[name.Formula.References[0].TargetSheetId!.Value],
                    null, null, false, geometries, name.Location);
            }
            catch (TemplateException error) { context.Report(error); }
        }
        context.ThrowIfErrors();
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
                    groups.Any(block => block.FirstRow <= cell.Source.Id.Address.Row && cell.Source.Id.Address.Row <= block.LastRow), geometries)); }
                catch (TemplateException error) { context.Report(error); }
            }
            sheets.Add(sheet with { Rows = sheet.Rows.Select(row => row with { Cells = row.Cells.Select(cell =>
                cell with { Formula = formulas.GetValueOrDefault(cell.Id) }).ToImmutableArray() }).ToImmutableArray() });
            if (context.LimitReached) break;
        }
        context.ThrowIfErrors(); return workbook with { Sheets = sheets.ToImmutable(), Names = names.ToImmutableArray() };
    }

    private static void RequireScalarFormulas(XlsxTemplatePlan template, XlsxOperationContext context)
    {
        foreach (var cell in template.Sheets.SelectMany(sheet => sheet.Cells).Where(cell => cell.SharedFormulaIndex is not null))
        {
            if (context.LimitReached) break;
            context.Report(Unsafe(cell.Location, "Shared formula group " + cell.SharedFormulaIndex +
                " requires shared-formula relocation before scalar output. Re-save the group as ordinary per-cell scalar A1 formulas."));
        }
        context.ThrowIfErrors();
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
        bool groupFormula = false, IReadOnlyDictionary<uint, XlsxSheetGeometry>? geometries = null)
    {
        var location = geometry.Source.Part + ", sheet '" + geometry.Source.Name + "', " + placement.Source.Id.Address;
        return Relocate(formula, geometry, placement.Id.Vertical, placement.Id.Horizontal, groupFormula, geometries, location);
    }

    private string Relocate(XlsxParsedFormula formula, XlsxSheetGeometry geometry, VerticalInstancePath? verticalOwner,
        HorizontalInstancePath? horizontalOwner, bool groupFormula, IReadOnlyDictionary<uint, XlsxSheetGeometry>? geometries, string location)
    {
        return _references.Rewrite(formula, reference =>
        {
            if (reference.NameIndex is not null) return formula.Text.Substring(reference.Start, reference.Length);
            var targetId = reference.TargetSheetId ?? (reference.Prefix.Length == 0 ? geometry.Source.Id :
                throw Unsafe(location, "Reference " + formula.Text.Substring(reference.Start, reference.Length) +
                    " has no resolved worksheet identity. Use an existing same-workbook worksheet and a retained A1 operand."));
            var target = targetId == geometry.Source.Id ? geometry :
                geometries is not null && geometries.TryGetValue(targetId, out var resolved) ? resolved :
                throw Unsafe(location, "Reference " + formula.Text.Substring(reference.Start, reference.Length) +
                    " has no retained worksheet geometry. Use an existing same-workbook worksheet and a retained A1 operand.");
            // Instance paths are worksheet-local; equal region/record indices on another sheet are not shared owners.
            var ownerVertical = targetId == geometry.Source.Id ? verticalOwner : null;
            var ownerHorizontal = targetId == geometry.Source.Id ? horizontalOwner : null;
            // Group range ownership is structural, including when dollar flags suppress authored copying.
            // Scalars and body/leaf formulas retain their exact-owner/absolute-anchor rules.
            var descendants = groupFormula && reference.IsRange && ownerVertical is not null;
            var vertical = !descendants && (reference.First.RowAbsolute || reference.Last.RowAbsolute) ? null : ownerVertical;
            var horizontal = reference.First.ColumnAbsolute || reference.Last.ColumnAbsolute ? null : ownerHorizontal;
            // Lexical endpoints retain their flags; geometry always consumes increasing bounds.
            var first = reference.First.Address;
            var last = reference.Last.Address;
            var range = new XlsxRectangle(new(Math.Min(first.Row, last.Row), Math.Min(first.Column, last.Column)),
                new(Math.Max(first.Row, last.Row), Math.Max(first.Column, last.Column)));
            var image = _geometry.ProjectReference(range, target, vertical, horizontal, descendants);
            if (target.Expanded && reference.IsRange &&
                (reference.First.RowAbsolute != reference.Last.RowAbsolute && ownerVertical is not null ||
                 reference.First.ColumnAbsolute != reference.Last.ColumnAbsolute && ownerHorizontal is not null))
            {
                var scoped = _geometry.ProjectReference(range, target,
                    reference.First.RowAbsolute != reference.Last.RowAbsolute ? ownerVertical : vertical,
                    reference.First.ColumnAbsolute != reference.Last.ColumnAbsolute ? ownerHorizontal : horizontal, descendants);
                if (scoped.Kind != image.Kind || scoped.Rectangle != image.Rectangle)
                    throw Unsafe(location, "Mixed endpoint copy flags select incompatible repeated owners. Re-author the A1 range with consistent flags or uniquely retained endpoints.");
            }
            if (!image.Complete || image.Kind is XlsxRangeImageKind.Empty or XlsxRangeImageKind.Disjoint or XlsxRangeImageKind.Ambiguous ||
                !reference.IsRange && image.Kind != XlsxRangeImageKind.Cell)
                throw UnsafeImage(location, formula, reference);
            var result = image.Rectangle!.Value;
            var firstOutput = new XlsxAddress(first.Row <= last.Row ? result.First.Row : result.Last.Row,
                first.Column <= last.Column ? result.First.Column : result.Last.Column);
            var lastOutput = new XlsxAddress(first.Row <= last.Row ? result.Last.Row : result.First.Row,
                first.Column <= last.Column ? result.Last.Column : result.First.Column);
            return reference.Prefix + (reference.First with { Address = firstOutput }).ToString() +
                (reference.IsRange ? ":" + (reference.Last with { Address = lastOutput }) : string.Empty);
        });
    }

    private static TemplateException Unsafe(string location, string remedy) => new StructuralDiagnostic(DiagnosticCodes.XlsxUnsafeFormula,
        remedy, Location: location, Explanation: remedy).ToException();

    private static TemplateException UnsafeImage(string location, XlsxParsedFormula formula, XlsxFormulaReference reference) =>
        Unsafe(location, "Reference " + formula.Text.Substring(reference.Start, reference.Length) +
            " has an empty, discarded, disjoint or ambiguous image. Re-author it inside one compatible owner using a retained contiguous A1 range; scalar operands must identify one cell.");
}
