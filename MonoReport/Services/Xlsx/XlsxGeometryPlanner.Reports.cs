using System.Collections.Immutable;
using System.Xml.Linq;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxGeometryPlanner
{
    internal XlsxWorkbookPlan? Preflight(XlsxTemplatePlan template, XlsxOperationContext context, bool forAnalysis = false)
    {
        if (template.Sheets.All(sheet => sheet.Controls.IsEmpty))
        {
            return Plan(template, ImmutableDictionary<SourceCellId, XlsxPreparedValue>.Empty, context);
        }
        foreach (var sheet in template.Sheets)
        {
            if (context.LimitReached) break;
            try
            {
                var nested = sheet.Regions.FirstOrDefault(region => !region.Root.Children.IsEmpty);
                // Analysis validates the parsed template without scheduling JSON-specific output.
                // Public generation remains gated independently of analyzer acceptance.
                if (nested is not null && !forAnalysis)
                    throw new StructuralDiagnostic(DiagnosticCodes.XlsxUnsupportedFeature,
                        "Nested emission and placement are pending in M3.",
                        Location: sheet.Controls.First(marker => marker.Kind == XlsxControlKind.Vertical && marker.Cell.Id.Address.Row == nested.FirstRow).Location,
                        Explanation: "Use one level per marked run until nested emission and placement are implemented in M3; separate independent runs with an unmarked control row.").ToException();
                ValidateFeatures(template.Source, sheet.Source);
                var area = sheet.Source.PrintArea;
                var identity = new XlsxSheetGeometry(sheet.Source, new(new(1, 1), new(area.Height, area.Width)), 1 - area.First.Row, 1 - area.First.Column)
                {
                    RowSpans = [new(area.First.Row, area.Last.Row, 1)], ColumnSpans = [new(area.First.Column, area.Last.Column, 1)]
                };
                var output = new XlsxSheetOutput(identity, new XElement(sheet.Source.Xml.Root!), []);
                if (sheet.Controls.IsEmpty) { CropColumns(output.Settings, identity); CropMerges(output.Settings, identity); }
                else ProjectMerges(sheet, output);
                CropViews(output.Settings, identity);
                CropBreaks(output.Settings, identity, true); CropBreaks(output.Settings, identity, false);
            }
            catch (MonoReport.Models.TemplateException error) { context.Report(error); }
        }
        context.ThrowIfErrors();
        return null;
    }

    internal XlsxWorkbookPlan PlanReports(XlsxTemplatePlan template, ImmutableArray<XlsxBoundSheet> data,
        ImmutableArray<XlsxLogicalSheetSchedule> schedules, XlsxOperationContext context)
    {
        var names = template.Source.Names.Select(name => new XElement(name)).ToArray();
        var sheets = ImmutableArray.CreateBuilder<XlsxSheetOutput>();
        for (var index = 0; index < template.Sheets.Length; index++)
        {
            var sheet = template.Sheets[index];
            var geometry = ExpansionGeometry(sheet, schedules[index], HorizontalSpans(sheet, data[index]));
            var output = EmitCells(sheet, geometry);
            output = PlanHeaderOwners(sheet, ProjectMerges(sheet, output));
            // Expanded view/break reconstruction is later fidelity work; never reuse offset-only transforms for it.
            if (sheet.Controls.Length > 0 && output.Settings.Descendants().Any(element => element.Name == S + "pane" || element.Name == S + "rowBreaks" || element.Name == S + "colBreaks"))
                throw Unsupported(sheet.Source, "Remove panes/manual page breaks from repeated reports until their unique-boundary transformation is implemented in M4.");
            CropViews(output.Settings, geometry);
            if (sheet.Controls.IsEmpty) { CropBreaks(output.Settings, geometry, true); CropBreaks(output.Settings, geometry, false); }
            names[sheet.Source.PrintAreaNameIndex].Value = SheetPrefix(sheet.Source.Name) + Absolute(geometry.OutputArea);
            sheets.Add(output);
        }
        context.OutputCells = sheets.Sum(sheet => sheet.Rows.Sum(row => (long)row.Cells.Length));
        return new(sheets.ToImmutable(), names.ToImmutableArray());
    }
}
