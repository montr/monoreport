using System.Collections.Immutable;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxTemplatePlanner
{
    internal XlsxSheetTemplate Structure(XlsxSheetTemplate sheet)
    {
        var geometry = sheet.Controls.Where(marker => marker.Kind == XlsxControlKind.Geometry).ToArray();
        var headers = sheet.Controls.Where(marker => marker.Kind == XlsxControlKind.Header).ToArray();
        TemplateException Error(string remedy) => ControlError(sheet.Controls.FirstOrDefault()?.Location ?? sheet.Source.Part, remedy);
        if (!geometry.Select(marker => marker.Level).Distinct().Order().SequenceEqual(headers.Select(marker => marker.Level).Distinct().Order()))
            throw Error("Associate every top geometry level with a header collection and every header with geometry above PrintArea.");
        var nextId = 0;
        var roots = new List<XlsxHorizontalNode>();
        foreach (var level in geometry.GroupBy(marker => marker.Level).OrderByDescending(group => group.Key))
        {
            if (level.Select(marker => marker.Cell.Id.Address.Row).Distinct().Count() != 1)
                throw Error("Use one separate top control row per horizontal depth.");
            var intervals = new List<(int First, int Last)>();
            foreach (var column in level.Select(marker => marker.Cell.Id.Address.Column).Order())
                if (intervals.Count > 0 && intervals[^1].Last + 1 == column) intervals[^1] = (intervals[^1].First, column);
                else intervals.Add((column, column));
            var parents = new List<XlsxHorizontalNode>();
            foreach (var interval in intervals)
            {
                var children = roots.Where(node => node.FirstColumn >= interval.First && node.LastColumn <= interval.Last).ToImmutableArray();
                parents.Add(new(nextId++, level.Key, interval.First, interval.Last, children));
            }
            if (roots.Any(child => !parents.Any(parent => parent.Children.Contains(child))))
                throw Error("Keep each child horizontal block wholly inside one parent block; partial or orphan children cannot be expanded.");
            roots = parents;
        }
        if (roots.Count > 1) throw Error("Use at most one outer horizontal region per worksheet; join adjacent parent columns or move the other report to another sheet.");
        var topRows = geometry.GroupBy(marker => marker.Cell.Id.Address.Row);
        if (topRows.Any(group => group.Select(marker => marker.Level).Distinct().Count() != 1))
            throw Error("Use a different top control row for each horizontal depth.");
        var orderedLevels = geometry.GroupBy(marker => marker.Level).OrderBy(group => group.Key).Select(group => group.First().Cell.Id.Address.Row).ToArray();
        if (!orderedLevels.SequenceEqual(orderedLevels.Order())) throw Error("Place deeper horizontal geometry controls below their parents.");
        var headerEnd = -1;
        foreach (var group in headers.GroupBy(marker => marker.Level).OrderBy(group => group.Key))
        {
            var positions = group.Select(marker => marker.Cell.Id.Address.Row).Order().ToArray();
            if (positions[0] <= headerEnd || positions[^1] - positions[0] + 1 != positions.Length)
                throw Error("Keep equivalent header rows consecutive and place deeper header blocks below their parents.");
            headerEnd = positions[^1];
        }
        var allNodes = new List<XlsxHorizontalNode>();
        void Collect(XlsxHorizontalNode node) { allNodes.Add(node); foreach (var child in node.Children) Collect(child); }
        foreach (var root in roots) Collect(root);
        foreach (var header in headers)
            foreach (var cell in sheet.Cells.Where(cell => cell.Source.Id.Address.Row == header.Cell.Id.Address.Row))
                if (cell.Tokens.Any(token => token.Expression?.Path is { Origin: not TemplatePathOrigin.Root }) &&
                    !allNodes.Any(node => node.Level == header.Level && node.FirstColumn <= cell.Source.Id.Address.Column && node.LastColumn >= cell.Source.Id.Address.Column))
                    throw ControlError(cell.Location, "Place header fields inside their matching horizontal block or use explicit root fields outside it.");
        var regions = new List<XlsxVerticalRegion>();
        foreach (var marker in sheet.Controls.Where(marker => marker.Kind == XlsxControlKind.Vertical).OrderBy(marker => marker.Cell.Id.Address.Row))
        {
            var row = marker.Cell.Id.Address.Row;
            if (regions.Count > 0 && row == regions[^1].LastRow + 1)
            {
                var region = regions[^1];
                if (region.Level != marker.Level || region.Collection != marker.CollectionPath)
                    throw new StructuralDiagnostic(DiagnosticCodes.XlsxUnsupportedFeature, "Nested vertical groups require M3.",
                        Location: marker.Location, Explanation: "Nested vertical groups require M3. Use one level/collection per marked run in M2; separate independent runs with an unmarked control row.").ToException();
                regions[^1] = region with { LastRow = row };
            }
            else regions.Add(new(regions.Count, row, row, marker.Level, marker.CollectionPath!));
        }
        return sheet with { Horizontal = roots.SingleOrDefault(), Regions = regions.ToImmutableArray() };
    }
}
