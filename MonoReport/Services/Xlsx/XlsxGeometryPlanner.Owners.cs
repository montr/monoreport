using System.Collections.Immutable;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxGeometryPlanner
{
    private static XlsxSheetOutput RetainMergeAnchor(XlsxSheetOutput output, XlsxCellSource anchor, XlsxRectangle rectangle,
        VerticalInstancePath? vertical, HorizontalInstancePath? horizontal)
    {
        var target = rectangle.First;
        var placement = new XlsxCellPlacement(new(anchor.Id, vertical, horizontal), target, anchor, null);
        // Invert the physical intervals, not populated placements: a blank source may have no authored c element.
        var rowSpan = output.Geometry.RowSpans.Single(span => span.OutputFirst <= target.Row && span.OutputLast >= target.Row);
        var columnSpan = output.Geometry.ColumnSpans.Single(span => span.OutputFirst <= target.Column && span.OutputLast >= target.Column);
        var physicalSource = new SourceCellId(anchor.Id.SheetId, new(rowSpan.First + target.Row - rowSpan.OutputFirst,
            columnSpan.First + target.Column - columnSpan.OutputFirst));
        var displaced = output.Geometry.DisplacedSources;
        if (physicalSource != anchor.Id)
            displaced = displaced.Add(new(physicalSource, rowSpan.Vertical, columnSpan.Horizontal, target));
        return output with
        {
            Geometry = output.Geometry with
            {
                ContentOwners = output.Geometry.ContentOwners.Add(new(anchor.Id, vertical, horizontal, target)),
                DisplacedSources = displaced
            },
            Rows = output.Rows.Select(row => row.OutputRow < rectangle.First.Row || row.OutputRow > rectangle.Last.Row ? row : row with
            {
                // The merge's source nonanchors have already been checked for independent meaningful content.
                // Clear only redundant copies of its anchor; preserve every cell's authored nonanchor styling.
                Cells = row.Cells.Where(cell => cell.Output != target).Select(cell =>
                    rectangle.Contains(cell.Output) && cell.Source.Id == anchor.Id
                        ? cell with { Value = new(null, null), Suppressed = true } : cell)
                    .Concat(row.OutputRow == target.Row ? new[] { placement } : [])
                    .OrderBy(cell => cell.Output.Column).ToImmutableArray()
            }).ToImmutableArray()
        };
    }

    // Decide content ownership before binding. Blank physical copies retain their authored styling.
    internal XlsxSheetOutput PlanHeaderOwners(XlsxSheetTemplate template, XlsxSheetOutput output)
    {
        var replacements = new Dictionary<CellInstanceId, XlsxCellPlacement>();
        var owners = output.Geometry.ContentOwners.ToBuilder();
        foreach (var header in template.Controls.Where(marker => marker.Kind == XlsxControlKind.Header))
            foreach (var source in template.Cells.Where(cell => cell.Source.Id.Address.Row == header.Cell.Id.Address.Row))
            {
                var instances = output.Rows.SelectMany(row => row.Cells).Where(cell => cell.Source.Id == source.Source.Id)
                    .GroupBy(cell => (cell.Id.Vertical, Horizontal: HeaderOwner(cell.Id.Horizontal, header.Level)));
                foreach (var group in instances)
                {
                    if (group.Key.Horizontal is null) continue; // Static columns outside the horizontal strip.
                    var retained = group.Where(cell => !cell.Suppressed).OrderBy(cell => cell.Output.Row).ThenBy(cell => cell.Output.Column).First();
                    foreach (var cell in group)
                        replacements.Add(cell.Id, cell == retained
                            ? cell with { Id = cell.Id with { Horizontal = group.Key.Horizontal } }
                            : cell with { Value = new(null, null), Suppressed = true });
                    for (var i = owners.Count - 1; i >= 0; i--)
                        if (owners[i].Source == source.Source.Id && owners[i].Vertical == group.Key.Vertical && owners[i].Horizontal == group.Key.Horizontal)
                            owners.RemoveAt(i);
                    owners.Add(new(source.Source.Id, group.Key.Vertical, group.Key.Horizontal, retained.Output));
                }
            }
        return output with
        {
            Geometry = output.Geometry with { ContentOwners = owners.ToImmutable() },
            Rows = output.Rows.Select(row => row with { Cells = row.Cells.Select(cell => replacements.GetValueOrDefault(cell.Id, cell)).ToImmutableArray() }).ToImmutableArray()
        };
    }

    private static HorizontalInstancePath? HeaderOwner(HorizontalInstancePath? path, int level)
    {
        while (path is not null && path.Level != level) path = path.Parent;
        return path;
    }

    // Reference images follow retained content owners, rather than suppressed physical header copies.
    internal XlsxRangeImage ProjectReference(XlsxRectangle range, XlsxSheetGeometry geometry,
        VerticalInstancePath? vertical, HorizontalInstancePath? horizontal, bool descendants = false)
    {
        // Retained content and the overwritten blank source are distinct identities, even when their physical
        // intervals end at the same output cell. Reject ranges too: merging their images would hide the alias.
        if (geometry.DisplacedSources.Any(source => range.Contains(source.Source.Address) && Compatible(source.Horizontal, horizontal) &&
            CompatibleVertical(source.Vertical, vertical, descendants)))
            return new(XlsxRangeImageKind.Ambiguous, null,
                geometry.Source.PrintArea.Contains(range.First) && geometry.Source.PrintArea.Contains(range.Last));
        var affected = geometry.ContentOwners.Where(owner => range.Contains(owner.Source.Address)).ToArray();
        if (affected.Length == 0) return geometry.Expanded || !geometry.RowSpans.IsEmpty
            ? ProjectInstances(range, geometry, vertical, horizontal, descendants) : ProjectRange(range, geometry);
        var remaining = new List<XlsxRectangle> { range };
        foreach (var address in affected.Select(owner => owner.Source.Address).Distinct())
        {
            var pieces = new List<XlsxRectangle>();
            foreach (var rectangle in remaining)
            {
                if (!rectangle.Contains(address)) { pieces.Add(rectangle); continue; }
                if (rectangle.First.Row < address.Row) pieces.Add(new(rectangle.First, new(address.Row - 1, rectangle.Last.Column)));
                if (rectangle.Last.Row > address.Row) pieces.Add(new(new(address.Row + 1, rectangle.First.Column), rectangle.Last));
                if (rectangle.First.Column < address.Column) pieces.Add(new(new(address.Row, rectangle.First.Column), new(address.Row, address.Column - 1)));
                if (rectangle.Last.Column > address.Column) pieces.Add(new(new(address.Row, address.Column + 1), new(address.Row, rectangle.Last.Column)));
            }
            remaining = pieces;
        }
        var images = new List<XlsxRectangle>();
        foreach (var piece in remaining)
        {
            var rows = AxisImage(piece.First.Row, piece.Last.Row, geometry.RowSpans.Where(span =>
                CompatibleVertical(span.Vertical, vertical, descendants)));
            var columns = AxisImage(piece.First.Column, piece.Last.Column, geometry.ColumnSpans.Where(span => Compatible(span.Horizontal, horizontal)));
            foreach (var row in rows)
                foreach (var column in columns) images.Add(new(new(row.First, column.First), new(row.Last, column.Last)));
        }
        foreach (var owner in affected.Where(owner => Compatible(owner.Horizontal, horizontal) &&
            CompatibleVertical(owner.Vertical, vertical, descendants)))
            images.Add(new(owner.Output, owner.Output));
        var complete = geometry.Source.PrintArea.Contains(range.First) && geometry.Source.PrintArea.Contains(range.Last);
        if (images.Count == 0) return new(XlsxRangeImageKind.Empty, null, complete);
        var first = new XlsxAddress(images.Min(image => image.First.Row), images.Min(image => image.First.Column));
        var last = new XlsxAddress(images.Max(image => image.Last.Row), images.Max(image => image.Last.Column));
        // Prove every row band has the same complete interval; never fill holes with a bounding rectangle.
        foreach (var row in images.SelectMany(image => new[] { image.First.Row, image.Last.Row + 1 }).Distinct().Order().SkipLast(1))
        {
            var end = first.Column - 1;
            foreach (var image in images.Where(image => image.First.Row <= row && image.Last.Row >= row).OrderBy(image => image.First.Column))
            {
                if (image.First.Column > end + 1) return new(XlsxRangeImageKind.Disjoint, null, complete);
                end = Math.Max(end, image.Last.Column);
            }
            if (end != last.Column) return new(XlsxRangeImageKind.Disjoint, null, complete);
        }
        return new(first == last ? XlsxRangeImageKind.Cell : XlsxRangeImageKind.Rectangle, new(first, last), complete);
    }
}
