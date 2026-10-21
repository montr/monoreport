using System.Collections.Immutable;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxTemplatePlanner
{
    private static ImmutableArray<XlsxVerticalRegion> VerticalRegions(XlsxSheetTemplate sheet)
    {
        var regions = ImmutableArray.CreateBuilder<XlsxVerticalRegion>();
        var run = new List<XlsxControlMarker>();
        void Finish()
        {
            if (run.Count == 0) return;
            var collection = run[0].CollectionPath!;
            var mismatch = run.FirstOrDefault(marker => marker.CollectionPath != collection);
            if (mismatch is not null)
                throw ControlError(mismatch.Location, "Use one case-sensitive root collection throughout a vertical region; separate independent regions with an unmarked control row.");
            regions.Add(new(regions.Count, run[0].Cell.Id.Address.Row, run[^1].Cell.Id.Address.Row,
                collection, VerticalTree(run)));
            run.Clear();
        }
        foreach (var marker in sheet.Controls.Where(marker => marker.Kind == XlsxControlKind.Vertical).OrderBy(marker => marker.Cell.Id.Address.Row))
        {
            if (run.Count > 0 && marker.Cell.Id.Address.Row != run[^1].Cell.Id.Address.Row + 1) Finish();
            run.Add(marker);
        }
        Finish();
        return regions.ToImmutable();
    }

    private static XlsxVerticalNode VerticalTree(List<XlsxControlMarker> run)
    {
        var blocks = new List<(XlsxControlMarker Marker, XlsxVerticalBlock Rows)>();
        foreach (var marker in run)
        {
            var row = marker.Cell.Id.Address.Row;
            if (blocks.Count > 0 && blocks[^1].Marker.Level == marker.Level)
                blocks[^1] = (blocks[^1].Marker, blocks[^1].Rows with { LastRow = row });
            else blocks.Add((marker, new(row, row)));
        }

        // Peel the outer numeric level from the ends of its source interval. A parent
        // can precede and/or follow its deeper body, never split that body. Iteration
        // avoids consuming the call stack for deeply nested authored levels.
        var nodes = new List<XlsxVerticalNode>();
        var first = 0;
        var last = blocks.Count - 1;
        foreach (var level in blocks.Select((block, index) => (block.Marker.Level, Index: index))
            .GroupBy(block => block.Level).OrderBy(group => group.Key))
        {
            var positions = level.Select(block => block.Index).ToArray();
            if (positions.Length > 2 || (positions[0] != first && positions[^1] != last) ||
                (positions.Length == 2 && (positions[0] != first || positions[1] != last)))
            {
                var offending = blocks[positions.Length > 1 ? positions[1] : positions[0]].Marker;
                throw ControlError(offending.Location, "The vertical header/body/footer interleaving is ambiguous. Separate independent regions with an unmarked control row.");
            }
            if (first == last)
            {
                nodes.Add(new(level.Key, null, blocks[first].Rows, null, []));
                break;
            }
            XlsxVerticalBlock? header = null;
            XlsxVerticalBlock? footer = null;
            if (positions[0] == first) header = blocks[first++].Rows;
            if (positions[^1] == last) footer = blocks[last--].Rows;
            nodes.Add(new(level.Key, header, null, footer, []));
        }
        var root = nodes[^1];
        for (var index = nodes.Count - 2; index >= 0; index--)
            root = nodes[index] with { Children = [root] };
        return root;
    }
}
