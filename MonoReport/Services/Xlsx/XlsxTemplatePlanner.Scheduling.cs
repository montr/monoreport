using System.Collections.Immutable;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxTemplatePlanner
{
    internal XlsxLogicalSheetSchedule ScheduleSheet(XlsxSheetTemplate sheet, XlsxBoundSheet data, XlsxOperationContext context)
    {
        var regions = ImmutableArray.CreateBuilder<XlsxVerticalRegionSchedule>();
        foreach (var region in sheet.Regions)
        {
            // Public preflight rejects nested reports; traversal belongs to M3.4.
            if (!region.Root.Children.IsEmpty)
                throw new NotSupportedException("Nested block scheduling requires M3.4 traversal.");
            var blocks = ImmutableArray.CreateBuilder<XlsxLogicalBlockInstance>();
            foreach (var row in data.Rows[region.Id])
            {
                if (row.Level != region.Level)
                {
                    context.Warnings.Add(new("No template block matches this data level; its own rows were omitted.", "$level", "$level",
                        sheet.Controls.First(marker => marker.Kind == XlsxControlKind.Vertical && marker.Cell.Id.Address.Row == region.FirstRow).Location, sheet.Source.Part,
                        region.Collection + "[" + row.FirstIndex + "]") { Code = DiagnosticCodes.XlsxMissingLevel });
                    continue;
                }
                blocks.Add(new(region.Root.Body!, XlsxLogicalBlockRole.Body, new(region.Id, [row.FirstIndex])));
            }
            regions.Add(new(region.Id, blocks.ToImmutable()));
        }
        return new(sheet.Source.Id, regions.ToImmutable());
    }
}
