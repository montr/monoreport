using System.Text.Json;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using DocumentFormat.OpenXml.Packaging;

namespace MonoReport.Services.Xlsx;

internal sealed class XlsxTemplateHandler(XlsxPackageReader reader, XlsxTemplatePlanner planner,
    XlsxGeometryPlanner geometry, XlsxDocumentRenderer renderer)
{
    private readonly XlsxPackageReader _reader = reader ?? throw new ArgumentNullException(nameof(reader));
    private readonly XlsxTemplatePlanner _planner = planner ?? throw new ArgumentNullException(nameof(planner));
    private readonly XlsxGeometryPlanner _geometry = geometry ?? throw new ArgumentNullException(nameof(geometry));
    private readonly XlsxDocumentRenderer _renderer = renderer ?? throw new ArgumentNullException(nameof(renderer));

    internal TemplateAnalysisResult Analyze(Stream package)
    {
        try
        {
            using var document = SpreadsheetDocument.Open(package, false);
            var context = new XlsxOperationContext(new MergeOptions()) { SourceBytes = package.Length };
            var plan = _planner.Plan(_reader.Read(document, context), context);
            _geometry.Plan(plan, System.Collections.Immutable.ImmutableDictionary<SourceCellId, XlsxPreparedValue>.Empty, context);
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
                var geometry = _geometry.Plan(source, System.Collections.Immutable.ImmutableDictionary<SourceCellId, XlsxPreparedValue>.Empty, context);
                var values = _planner.Bind(source, data.RootElement, context);
                var plan = _planner.AttachValues(geometry, values);
                _renderer.Render(document, plan, context);
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
