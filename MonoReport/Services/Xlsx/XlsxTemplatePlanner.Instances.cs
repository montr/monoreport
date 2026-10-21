using System.Collections.Immutable;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxTemplatePlanner
{
    internal XlsxSheetOutput BindInstances(XlsxSheetTemplate template, XlsxSheetOutput output, XlsxBoundSheet data, XlsxOperationContext context)
    {
        var plans = template.Cells.ToDictionary(cell => cell.Source.Id);
        var headers = template.Controls.Where(marker => marker.Kind == XlsxControlKind.Header).ToDictionary(marker => marker.Cell.Id.Address.Row);
        var headerDepths = data.Headers.Select((header, index) => (header.Level, index)).ToDictionary(item => item.Level, item => item.index);
        var logicalRows = data.Rows.ToDictionary(region => region.Key, region => region.Value.ToDictionary(row => row.FirstIndex));
        var nodes = new Dictionary<int, XlsxHorizontalNode>();
        void Index(XlsxHorizontalNode node) { nodes.Add(node.Id, node); foreach (var child in node.Children) Index(child); }
        if (template.Horizontal is { } root) Index(root);
        XlsxCellPlacement BindCell(XlsxCellPlacement placement)
        {
            if (context.LimitReached || placement.Suppressed) return placement;
            var cell = plans[placement.Source.Id];
            var scope = new DataScope(data.Root, [data.Root], ["$"]);
            var path = new List<HorizontalInstancePath>();
            for (var item = placement.Id.Horizontal; item is not null; item = item.Parent) path.Add(item);
            path.Reverse();
            if (headers.TryGetValue(cell.Source.Id.Address.Row, out var header))
            {
                var depth = headerDepths[header.Level];
                if (path.Count <= depth)
                {
                    if (cell.Tokens.Any(token => token.Expression?.Path is { Origin: not TemplatePathOrigin.Root }))
                        throw ControlError(cell.Location, "Place header fields inside their matching horizontal block or use explicit root fields outside it.");
                }
                else
                {
                    scope = HeaderScope(data.Root, data.Headers, path.Take(depth + 1).Select(item => item.ItemIndex).ToArray());
                }
            }
            else if (placement.Id.Vertical is { } vertical)
            {
                var region = template.Regions.Single(item => item.Id == vertical.RegionId);
                var row = logicalRows[region.Id][vertical.LogicalRows[0]];
                int? tuple = null;
                if (path.Count == data.Headers.Length && path.Count > 0 && nodes[path[^1].NodeId].Children.IsEmpty)
                {
                    var index = 0;
                    foreach (var headerCollection in data.Headers)
                    {
                        var instance = path.Single(item => nodes[item.NodeId].Level == headerCollection.Level);
                        index = checked(index * headerCollection.Items.Length + instance.ItemIndex);
                    }
                    tuple = index;
                }
                scope = RowScope(data.Root, row, region.Collection, tuple);
            }
            return cell.Tokens.Length == 0 ? placement : placement with { Value = PrepareValue(cell, scope, context) };
        }
        var prepared = new Dictionary<CellInstanceId, XlsxCellPlacement>();
        foreach (var cell in output.Rows.SelectMany(row => row.Cells).OrderBy(cell => cell.Source.Id.Address.Row)
            .ThenBy(cell => cell.Source.Id.Address.Column).ThenBy(cell => cell.Output.Row).ThenBy(cell => cell.Output.Column))
        {
            if (context.LimitReached) break;
            prepared.Add(cell.Id, BindCell(cell));
        }
        context.ThrowIfErrors();
        var result = output with { Rows = output.Rows.Select(row => row with { Cells = row.Cells.Select(cell => prepared[cell.Id]).ToImmutableArray() }).ToImmutableArray() };
        context.ThrowIfErrors(); return result;
    }
}
