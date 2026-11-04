using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;

namespace MonoReport.Xlsx.Planning;

public sealed partial class XlsxTemplatePlanner
{
    internal XlsxLogicalSheetSchedule ScheduleSheet(XlsxSheetTemplate sheet, XlsxBoundSheet data,
        XlsxOperationContext context)
    {
        var regions = ImmutableArray.CreateBuilder<XlsxVerticalRegionSchedule>();
        foreach (var region in sheet.Regions)
        {
            var levels = new Dictionary<int, XlsxVerticalNode>();
            var source = new Stack<XlsxVerticalNode>();
            source.Push(region.Root);
            while (source.TryPop(out var node))
            {
                levels.Add(node.Level, node);
                foreach (var child in node.Children) source.Push(child);
            }

            var blocks = ImmutableArray.CreateBuilder<XlsxLogicalBlockInstance>();
            var rows = data.Rows[region.Id];
            var byIndex = rows.ToDictionary(row => row.FirstIndex);
            var pending = new Stack<(XlsxLogicalRow Row, VerticalInstancePath Owner, bool Exit)>();
            foreach (var root in rows.Where(row => row.Parent is null).Reverse())
                pending.Push((root, new VerticalInstancePath(region.Id, [root.FirstIndex]), false));
            while (pending.TryPop(out var visit))
            {
                levels.TryGetValue(visit.Row.Level, out var node);
                if (visit.Exit)
                {
                    if (node?.Footer is { } footer)
                        blocks.Add(new XlsxLogicalBlockInstance(footer, XlsxLogicalBlockRole.Footer, visit.Owner));
                    continue;
                }

                if (node?.Header is { } header)
                    blocks.Add(new XlsxLogicalBlockInstance(header, XlsxLogicalBlockRole.Header, visit.Owner));
                if (node?.Body is { } body)
                    blocks.Add(new XlsxLogicalBlockInstance(body, XlsxLogicalBlockRole.Body, visit.Owner));
                pending.Push((visit.Row, visit.Owner, true));
                foreach (var child in visit.Row.ChildFirstIndices.Reverse())
                    pending.Push((byIndex[child],
                        new VerticalInstancePath(region.Id, visit.Owner.LogicalRows.Add(child)), false));
            }

            regions.Add(new XlsxVerticalRegionSchedule(region.Id, blocks.ToImmutable()));
        }

        return new XlsxLogicalSheetSchedule(sheet.Source.Id, regions.ToImmutable());
    }
}