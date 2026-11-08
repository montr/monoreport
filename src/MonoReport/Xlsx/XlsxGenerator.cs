using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Templates;
using MonoReport.Xlsx.Planning;

namespace MonoReport.Xlsx;

public sealed class XlsxGenerator(
    XlsxPackageReader reader,
    XlsxTemplatePlanner planner,
    XlsxGeometryPlanner geometry,
    XlsxFormulaPlanner formulas,
    XlsxDocumentRenderer renderer)
{
    private readonly XlsxFormulaPlanner _formulas = formulas ?? throw new ArgumentNullException(nameof(formulas));
    private readonly XlsxGeometryPlanner _geometry = geometry ?? throw new ArgumentNullException(nameof(geometry));
    private readonly XlsxTemplatePlanner _planner = planner ?? throw new ArgumentNullException(nameof(planner));
    private readonly XlsxPackageReader _reader = reader ?? throw new ArgumentNullException(nameof(reader));
    private readonly XlsxDocumentRenderer _renderer = renderer ?? throw new ArgumentNullException(nameof(renderer));

    internal GenerationResult Generate(MemoryStream package, JsonDocument data, GenerationOptions options)
    {
        try
        {
            var context = new XlsxOperationContext(options) {SourceBytes = package.Length};
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
                    var bound = source.Sheets.Select(sheet => _planner.BindSheet(sheet, data.RootElement, context))
                        .ToImmutableArray();
                    var schedules = source.Sheets
                        .Select((sheet, index) => _planner.ScheduleSheet(sheet, bound[index], context))
                        .ToImmutableArray();
                    plan = _geometry.PlanReports(source, bound, schedules, context);
                    plan = plan with
                    {
                        Sheets = plan.Sheets.Select((sheet, index) =>
                                _planner.BindInstances(source.Sheets[index], sheet, bound[index], context))
                            .ToImmutableArray()
                    };
                }

                plan = _formulas.Finalize(source, plan, context);
                _renderer.Render(document, plan, context);
                var sourceOrder = source.Sheets.SelectMany((sheet, index) => sheet.Cells
                        .Select(cell => (cell.Location, Sheet: index, cell.Source.Id.Address))
                        .Concat(sheet.Controls.Select(marker =>
                            (marker.Location, Sheet: index, marker.Cell.Id.Address))))
                    .ToDictionary(item => item.Location, item => (item.Sheet, item.Address.Row, item.Address.Column),
                        StringComparer.Ordinal);
                var warnings = context.Warnings.OrderBy(warning => sourceOrder.GetValueOrDefault(
                    warning.Location ?? string.Empty,
                    (int.MaxValue, int.MaxValue, int.MaxValue))).ToArray();
                context.Warnings.Clear();
                context.Warnings.AddRange(warnings);
            }

            return new GenerationResult(package.ToArray(), context.Warnings.AsReadOnly())
                {Format = DocumentFormat.Xlsx};
        }
        catch (TemplateException)
        {
            throw;
        }
        catch (Exception error) when (error is OpenXmlPackageException or XmlException or InvalidDataException
                                          or KeyNotFoundException or ArgumentException or FileFormatException)
        {
            throw new TemplateException("The XLSX package is inconsistent.", code: DiagnosticCodes.XlsxInvalidPackage);
        }
    }
}