using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Templates;
using MonoReport.Xlsx.Planning;
using System.IO;
using System;
using System.Linq;
using System.Collections.Generic;

namespace MonoReport.Xlsx;

public sealed class XlsxTemplateAnalyzer(
    XlsxPackageReader reader,
    XlsxTemplatePlanner planner,
    XlsxGeometryPlanner geometry,
    XlsxFormulaPlanner formulas)
{
    private readonly XlsxFormulaPlanner _formulas = formulas ?? throw new ArgumentNullException(nameof(formulas));
    private readonly XlsxGeometryPlanner _geometry = geometry ?? throw new ArgumentNullException(nameof(geometry));
    private readonly XlsxTemplatePlanner _planner = planner ?? throw new ArgumentNullException(nameof(planner));
    private readonly XlsxPackageReader _reader = reader ?? throw new ArgumentNullException(nameof(reader));

    internal TemplateAnalysisResult Analyze(Stream package)
    {
        try
        {
            using var document = SpreadsheetDocument.Open(package, false);
            var context = new XlsxOperationContext(new GenerationOptions()) {SourceBytes = package.Length};
            var plan = _planner.Plan(_reader.Read(document, context), context);
            var staticGeometry = _geometry.Preflight(plan, context, true);
            _formulas.ValidateStaticReferences(plan, context);
            if (staticGeometry is not null) _formulas.Finalize(plan, staticGeometry, context);
            return new TemplateAnalysisResult([], requirements: plan.Requirements, occurrences: plan.Occurrences);
        }
        catch (TemplateException error)
        {
            return new TemplateAnalysisResult(error.Errors.Select(item => new TemplateValidationIssue(item.Code,
                item.Message, item.Field, item.Part, item.Location)));
        }
        catch (Exception error) when (error is OpenXmlPackageException or XmlException or InvalidDataException
                                          or KeyNotFoundException or ArgumentException or FileFormatException)
        {
            return new TemplateAnalysisResult([
                new TemplateValidationIssue(DiagnosticCodes.XlsxInvalidPackage,
                    "The XLSX package is inconsistent. Repair or re-save the workbook with valid parts and relationships.")
            ]);
        }
    }
}
