using System.Collections.Immutable;
using System.Text.Json;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using DocumentFormat.OpenXml.Packaging;

namespace MonoReport.Services.Xlsx;

internal sealed class XlsxTemplateHandler(XlsxPackageReader reader, XlsxTemplatePlanner planner,
    XlsxGeometryPlanner geometry, XlsxFormulaPlanner formulas, XlsxDocumentRenderer renderer)
{
    private readonly XlsxPackageReader _reader = reader ?? throw new ArgumentNullException(nameof(reader));
    private readonly XlsxTemplatePlanner _planner = planner ?? throw new ArgumentNullException(nameof(planner));
    private readonly XlsxGeometryPlanner _geometry = geometry ?? throw new ArgumentNullException(nameof(geometry));
    private readonly XlsxFormulaPlanner _formulas = formulas ?? throw new ArgumentNullException(nameof(formulas));
    private readonly XlsxDocumentRenderer _renderer = renderer ?? throw new ArgumentNullException(nameof(renderer));

    internal TemplateAnalysisResult Analyze(Stream package)
    {
        try
        {
            using var document = SpreadsheetDocument.Open(package, false);
            var context = new XlsxOperationContext(new MergeOptions()) { SourceBytes = package.Length };
            var plan = _planner.Plan(_reader.Read(document, context), context);
            var staticGeometry = _geometry.Preflight(plan, context);
            _formulas.ValidateStaticReferences(plan, context);
            if (staticGeometry is not null) _formulas.Finalize(plan, staticGeometry, context);
            return new([], requirements: plan.Requirements, occurrences: plan.Occurrences);
        }
        catch (TemplateException error)
        {
            return new(error.Errors.Select(item => new TemplateValidationIssue(item.Code, item.Message,
                item.Field, item.Part, item.Location)));
        }
        catch (Exception error) when (error is OpenXmlPackageException or System.Xml.XmlException or InvalidDataException or KeyNotFoundException or ArgumentException or FileFormatException)
        {
            return new([new(DiagnosticCodes.XlsxInvalidPackage, "The XLSX package is inconsistent. Repair or re-save the workbook with valid parts and relationships.")]);
        }
    }

    internal MergeResult Merge(MemoryStream package, JsonDocument data, MergeOptions options)
    {
        try
        {
            var context = new XlsxOperationContext(options) { SourceBytes = package.Length };
            using (var document = SpreadsheetDocument.Open(package, true))
            {
                var source = _planner.Plan(_reader.Read(document, context), context);
                // Complete data-independent preflight before binding or mutating any worksheet.
                var staticGeometry = _geometry.Preflight(source, context);
                _formulas.ValidateStaticReferences(source, context);
                XlsxWorkbookPlan plan;
                if (staticGeometry is not null)
                {
                    plan = _planner.AttachValues(staticGeometry, _planner.Bind(source, data.RootElement, context));
                }
                else
                {
                    var bound = source.Sheets.Select(sheet => _planner.BindSheet(sheet, data.RootElement, context)).ToImmutableArray();
                    var schedules = source.Sheets.Select((sheet, index) => _planner.ScheduleSheet(sheet, bound[index], context)).ToImmutableArray();
                    plan = _geometry.PlanReports(source, bound, schedules, context);
                    plan = plan with { Sheets = plan.Sheets.Select((sheet, index) => _planner.BindInstances(source.Sheets[index], sheet, bound[index], context)).ToImmutableArray() };
                }
                plan = _formulas.Finalize(source, plan, context);
                _renderer.Render(document, plan, context);
                var sourceOrder = source.Sheets.SelectMany((sheet, index) => sheet.Cells.Select(cell => (cell.Location, Sheet: index, Address: cell.Source.Id.Address))
                    .Concat(sheet.Controls.Select(marker => (marker.Location, Sheet: index, Address: marker.Cell.Id.Address))))
                    .ToDictionary(item => item.Location, item => (item.Sheet, item.Address.Row, item.Address.Column), StringComparer.Ordinal);
                var warnings = context.Warnings.OrderBy(warning => sourceOrder.GetValueOrDefault(warning.Location ?? string.Empty,
                    (int.MaxValue, int.MaxValue, int.MaxValue))).ToArray();
                context.Warnings.Clear(); context.Warnings.AddRange(warnings);
            }
            return new(package.ToArray(), context.Warnings.AsReadOnly()) { Format = Models.DocumentFormat.Xlsx };
        }
        catch (TemplateException) { throw; }
        catch (Exception error) when (error is OpenXmlPackageException or System.Xml.XmlException or InvalidDataException or KeyNotFoundException or ArgumentException or FileFormatException)
        {
            throw new TemplateException("The XLSX package is inconsistent.", code: DiagnosticCodes.XlsxInvalidPackage);
        }
    }
}
