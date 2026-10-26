using System.Collections.Immutable;
using System.Xml.Linq;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxGeometryPlanner
{
    internal XlsxRangeImage ProjectInstances(XlsxRectangle range, XlsxSheetGeometry geometry,
        VerticalInstancePath? vertical, HorizontalInstancePath? horizontal, bool descendants = false)
    {
        var area = geometry.Source.RetentionArea;
        var complete = area.Contains(range.First) && area.Contains(range.Last);
        var rows = AxisImage(range.First.Row, range.Last.Row, geometry.RowSpans.Where(span => CompatibleVertical(span.Vertical, vertical, descendants)), out var rowsComplete);
        var columns = AxisImage(range.First.Column, range.Last.Column, geometry.ColumnSpans.Where(span => Compatible(span.Horizontal, horizontal)), out var columnsComplete);
        complete &= rowsComplete && columnsComplete;
        if (rows.Count == 0 || columns.Count == 0) return new(XlsxRangeImageKind.Empty, null, complete);
        if (rows.Count != 1 || columns.Count != 1) return new(XlsxRangeImageKind.Disjoint, null, complete);
        var result = new XlsxRectangle(new(rows[0].First, columns[0].First), new(rows[0].Last, columns[0].Last));
        return new(result.First == result.Last ? XlsxRangeImageKind.Cell : XlsxRangeImageKind.Rectangle, result, complete);
    }

    private static bool CompatibleVertical(VerticalInstancePath? candidate, VerticalInstancePath? scope, bool descendants) =>
        scope is null || candidate is null || candidate.RegionId != scope.RegionId || candidate == scope ||
        descendants && candidate.LogicalRows.AsSpan().StartsWith(scope.LogicalRows.AsSpan());

    private static bool Compatible(HorizontalInstancePath? candidate, HorizontalInstancePath? scope)
    {
        for (var expected = scope; expected is not null; expected = expected.Parent)
            for (var actual = candidate; actual is not null; actual = actual.Parent)
                if ((expected.NodeId == actual.NodeId || expected.Level == actual.Level) && expected.ItemIndex != actual.ItemIndex) return false;
        return true;
    }

    private static List<(int First, int Last)> AxisImage(int first, int last, IEnumerable<XlsxAxisSpan> spans, out bool complete)
    {
        var clipped = spans.Where(span => span.First <= last && first <= span.Last)
            .Select(span => (First: Math.Max(first, span.First), Last: Math.Min(last, span.Last),
                Offset: span.OutputFirst - span.First)).ToArray();
        // Repetitions can cover the same source interval; output contiguity cannot prove source coverage.
        var next = first;
        complete = true;
        foreach (var span in clipped.OrderBy(span => span.First))
        {
            if (span.First > next) complete = false;
            next = Math.Max(next, span.Last + 1);
        }
        complete &= next > last;
        var projected = clipped.Select(span => (First: span.First + span.Offset, Last: span.Last + span.Offset))
            .OrderBy(span => span.First);
        var result = new List<(int First, int Last)>();
        foreach (var span in projected)
            if (result.Count > 0 && span.First <= result[^1].Last + 1) result[^1] = (result[^1].First, Math.Max(result[^1].Last, span.Last));
            else result.Add(span);
        return result;
    }

    internal XlsxSheetOutput ProjectMerges(XlsxSheetTemplate sheet, XlsxSheetOutput output)
    {
        var container = output.Settings.Element(S + "mergeCells");
        if (container is null) return output;
        var source = container.Elements().ToArray(); container.RemoveNodes();
        var nodes = new List<XlsxHorizontalNode>();
        void Index(XlsxHorizontalNode node) { nodes.Add(node); foreach (var child in node.Children) Index(child); }
        if (sheet.Horizontal is { } root) Index(root);
        var used = new List<XlsxRectangle>();
        foreach (var merge in source)
        {
            var range = ParseRange((string?)merge.Attribute("ref"), output.Geometry);
            if (range.Last.Row < sheet.Source.RetentionArea.First.Row || range.First.Row > sheet.Source.RetentionArea.Last.Row ||
                range.Last.Column < sheet.Source.RetentionArea.First.Column || range.First.Column > sheet.Source.RetentionArea.Last.Column) continue;
            if (!sheet.Source.RetentionArea.Contains(range.First) || !sheet.Source.RetentionArea.Contains(range.Last))
                throw Unsupported(sheet.Source, "Move the merge wholly inside or outside PrintArea; a partial merge cannot survive cropping.", range.ToString());
            if (sheet.Cells.Any(cell => range.Contains(cell.Source.Id.Address) && cell.Source.Id.Address != range.First &&
                (cell.Source.HasFormula || cell.Source.Text is { Length: > 0 } || cell.Source.Xml.Element(S + "v") is { Value.Length: > 0 })))
                throw Unsupported(sheet.Source, "Clear meaningful nonanchor merge content before generation; merging would lose data.", range.ToString());
            foreach (var node in nodes)
                if (node.FirstColumn <= range.Last.Column && range.First.Column <= node.LastColumn &&
                    !(range.First.Column <= node.FirstColumn && range.Last.Column >= node.LastColumn) &&
                    !(node.FirstColumn <= range.First.Column && node.LastColumn >= range.Last.Column))
                    throw Unsupported(sheet.Source, "Keep merges inside one owner or fully include its children; partial/cross-owner merges are unsafe.", range.ToString());
            var owner = nodes.Where(node => node.FirstColumn <= range.First.Column && node.LastColumn >= range.Last.Column)
                .OrderBy(node => node.LastColumn - node.FirstColumn).FirstOrDefault();
            var header = sheet.Controls.FirstOrDefault(marker => marker.Kind == XlsxControlKind.Header && marker.Cell.Id.Address.Row == range.First.Row);
            if (header is not null)
            {
                if (sheet.Controls.Any(marker => marker.Kind == XlsxControlKind.Header && marker.Cell.Id.Address.Row <= range.Last.Row &&
                    marker.Cell.Id.Address.Row >= range.First.Row && marker.Level != header.Level))
                    throw Unsupported(sheet.Source, "Keep heading merges within one header level; different header scopes cannot share an owner.", range.ToString());
                owner = nodes.SingleOrDefault(node => node.Level == header.Level && node.FirstColumn <= range.First.Column && node.LastColumn >= range.Last.Column);
            }
            if (owner is null && nodes.Any(node => node.FirstColumn <= range.Last.Column && range.First.Column <= node.LastColumn))
                throw Unsupported(sheet.Source, "Keep the merge inside one horizontal owner; a merge cannot mix static and repeated columns.", range.ToString());
            var region = sheet.Regions.FirstOrDefault(item => item.FirstRow <= range.First.Row && item.LastRow >= range.Last.Row);
            if (region is null && sheet.Regions.Any(item => item.FirstRow <= range.Last.Row && range.First.Row <= item.LastRow))
                throw Unsupported(sheet.Source, "Keep the merge inside one vertical owner; a merge cannot cross repeated/static rows.", range.ToString());
            if (region is not null)
            {
                var pending = new Stack<XlsxVerticalNode>(); pending.Push(region.Root);
                var contained = false;
                while (pending.TryPop(out var node))
                {
                    if (new[] { node.Header, node.Body, node.Footer }.Any(block => block is not null &&
                        block.FirstRow <= range.First.Row && range.Last.Row <= block.LastRow)) contained = true;
                    foreach (var child in node.Children) pending.Push(child);
                }
                if (!contained)
                    throw Unsupported(sheet.Source, "Keep the merge wholly inside one vertical source block; cross-block merges cannot share an owner safely.", range.ToString());
            }
            var horizontal = owner is null ? new HorizontalInstancePath?[] { null } : output.Geometry.ColumnSpans
                .Select(span => FindOwner(span.Horizontal, owner.Id)).Where(path => path is not null).Distinct().ToArray();
            var vertical = region is null ? new VerticalInstancePath?[] { null } : output.Geometry.RowSpans
                .Where(span => span.First <= range.Last.Row && range.First.Row <= span.Last)
                .Select(span => span.Vertical).Where(path => path?.RegionId == region.Id).Distinct().ToArray();
            foreach (var v in vertical)
                foreach (var h in horizontal)
                {
                    var image = ProjectInstances(range, output.Geometry, v, h);
                    if (image.Kind == XlsxRangeImageKind.Empty) continue;
                    if (image.Kind is not (XlsxRangeImageKind.Cell or XlsxRangeImageKind.Rectangle))
                        throw Unsupported(sheet.Source, "Re-author the merge inside one contiguous owner; its image is disjoint.", range.ToString());
                    var rectangle = image.Rectangle!.Value;
                    if (used.Any(prior => prior.First.Row <= rectangle.Last.Row && rectangle.First.Row <= prior.Last.Row &&
                        prior.First.Column <= rectangle.Last.Column && rectangle.First.Column <= prior.Last.Column))
                        throw Unsupported(sheet.Source, "Remove overlapping merge projections.", range.ToString());
                    used.Add(rectangle);
                    if (sheet.Cells.FirstOrDefault(cell => cell.Source.Id.Address == range.First) is { } anchor)
                        output = RetainMergeAnchor(output, anchor.Source, rectangle, v, h);
                    if (image.Kind != XlsxRangeImageKind.Cell) container.Add(new XElement(S + "mergeCell", new XAttribute("ref", rectangle)));
                }
        }
        if (!container.HasElements) container.Remove(); else container.SetAttributeValue("count", container.Elements().Count());
        return output;
    }

    private static HorizontalInstancePath? FindOwner(HorizontalInstancePath? path, int node)
    {
        while (path is not null && path.NodeId != node) path = path.Parent;
        return path;
    }
}
