using System.Collections.Immutable;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxGeometryPlanner
{
    internal XlsxSheetGeometry ExpansionGeometry(XlsxSheetTemplate sheet, XlsxLogicalSheetSchedule schedule,
        ImmutableArray<XlsxAxisSpan> columns)
    {
        if (schedule.SheetId != sheet.Source.Id || !sheet.Regions.Select(region => region.Id).SequenceEqual(schedule.Regions.Select(region => region.RegionId)))
            throw new InvalidOperationException("The vertical schedule must match the sheet and its source regions in order.");
        var rows = ImmutableArray.CreateBuilder<XlsxAxisSpan>();
        var source = sheet.Source.PrintArea.First.Row;
        var at = 1;
        long height = sheet.Source.PrintArea.Height;
        for (var index = 0; index < sheet.Regions.Length; index++)
        {
            var region = sheet.Regions[index];
            height = checked(height - (region.LastRow - region.FirstRow + 1));
            foreach (var block in schedule.Regions[index].Blocks)
                height = checked(height + (long)block.Source.LastRow - block.Source.FirstRow + 1);
        }
        var width = columns.IsEmpty ? 1 : columns[^1].OutputLast;
        Dimensions(Math.Max(1, height), width, 1, 1);
        void Emit(int first, int last, VerticalInstancePath? path)
        {
            if (last < first) return;
            rows.Add(new(first, last, at, path)); at = checked(at + last - first + 1);
        }
        for (var index = 0; index < sheet.Regions.Length; index++)
        {
            var region = sheet.Regions[index];
            Emit(source, region.FirstRow - 1, null);
            foreach (var block in schedule.Regions[index].Blocks)
                Emit(block.Source.FirstRow, block.Source.LastRow, block.Owner);
            source = region.LastRow + 1;
        }
        Emit(source, sheet.Source.PrintArea.Last.Row, null);
        return new(sheet.Source, new(new(1, 1), new(Math.Max(1, at - 1), width)),
            1 - sheet.Source.PrintArea.First.Row, 1 - sheet.Source.PrintArea.First.Column)
        { Expanded = true, RowSpans = rows.ToImmutable(), ColumnSpans = columns };
    }

    internal ImmutableArray<XlsxAxisSpan> HorizontalSpans(XlsxSheetTemplate sheet, XlsxBoundSheet data)
    {
        var spans = ImmutableArray.CreateBuilder<XlsxAxisSpan>();
        var at = 1;
        void Emit(int first, int last, HorizontalInstancePath? path)
        {
            if (last < first) return;
            var width = checked(last - first + 1);
            Dimensions(1, (long)at - 1 + width, 1, 1);
            spans.Add(new(first, last, at, Horizontal: path)); at = checked(at + width);
        }
        int Width(XlsxHorizontalNode node)
        {
            var own = node.LastColumn - node.FirstColumn + 1 - node.Children.Sum(child => child.LastColumn - child.FirstColumn + 1);
            var count = data.Headers.Single(header => header.Level == node.Level).Items.Length;
            if (count == 0) return 0;
            var perItem = checked((long)own + node.Children.Sum(child => (long)Width(child)));
            return Dimensions(1, perItem, 1, count).Columns;
        }
        void Node(XlsxHorizontalNode node, HorizontalInstancePath? parent)
        {
            var count = data.Headers.Single(header => header.Level == node.Level).Items.Length;
            for (var index = 0; index < count; index++)
            {
                var path = new HorizontalInstancePath(node.Id, index, parent) { Level = node.Level };
                var source = node.FirstColumn;
                foreach (var child in node.Children)
                {
                    Emit(source, child.FirstColumn - 1, path); Node(child, path); source = child.LastColumn + 1;
                }
                Emit(source, node.LastColumn, path);
            }
        }
        var area = sheet.Source.PrintArea;
        if (sheet.Horizontal is { } root)
        {
            var width = checked((long)area.Width - (root.LastColumn - root.FirstColumn + 1) + Width(root));
            Dimensions(1, width, 1, 1);
            Emit(area.First.Column, root.FirstColumn - 1, null);
            Node(root, null);
            Emit(root.LastColumn + 1, area.Last.Column, null);
        }
        else Emit(area.First.Column, area.Last.Column, null);
        return spans.ToImmutable();
    }
}
