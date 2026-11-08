using System.Collections.Generic;
using Microsoft.AspNetCore.Http;
using MonoReport;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Templates;

namespace MonoReport.Api.Services;

public sealed class ApiProblemFactory
{
    public IResult Create(HttpContext context, int status, string code, string? field = null,
        IReadOnlyList<string>? errors = null)
    {
        var extensions = new Dictionary<string, object?>
        {
            ["code"] = code, ["traceId"] = context.TraceIdentifier, ["field"] = field
        };
        if (errors is not null) extensions["errors"] = errors;
        return Results.Problem(
            statusCode: status, type: $"urn:monoreport:error:{code}", title: status switch
            {
                400 => "Invalid request", 404 => "Formatter not found", 413 => "Request too large", 415 => "Unsupported media type",
                422 when code == ApiDiagnosticCodes.FormatterInvalidValue => "Formatter execution failed",
                422 => "Document generation failed", _ => "Internal server error"
            }, detail: SafeMessage(code), extensions: extensions);
    }

    public IResult Create(HttpContext context, TemplateException error)
    {
        var sampleRequest = string.Equals(context.Request.Path.Value, ApiRoutePaths.DocumentSample,
            System.StringComparison.Ordinal);
        var extensions = new Dictionary<string, object?>
        {
            ["code"] = error.Code,
            ["traceId"] = context.TraceIdentifier,
            ["field"] = error.Field,
            ["errors"] = error.Errors,
            ["limitReached"] = error.LimitReached
        };
        return Results.Problem(statusCode: 422, type: $"urn:monoreport:error:{error.Code}",
            title: sampleRequest ? "Sample generation failed" : "Document generation failed",
            detail: error.Code is DiagnosticCodes.SampleElementLimit or DiagnosticCodes.SampleIncompatibleDemand
                ? error.Message
                : SafeMessage(error.Code), extensions: extensions);
    }

    public string SafeMessage(string code)
    {
        return code switch
        {
            ApiDiagnosticCodes.RequestMissingInput => "A required input is missing.",
            ApiDiagnosticCodes.RequestEmptyInput => "A required input is empty.",
            ApiDiagnosticCodes.RequestDuplicateInput => "An input was supplied more than once.",
            ApiDiagnosticCodes.RequestInvalidInput => "An input has the wrong representation.",
            ApiDiagnosticCodes.RequestUnexpectedInput => "Validation accepts only the template file.",
            ApiDiagnosticCodes.RequestInvalidMultipart => "The multipart request is malformed.",
            ApiDiagnosticCodes.RequestTooLarge => "The whole request exceeds the configured size limit.",
            ApiDiagnosticCodes.RequestUnsupportedMediaType => "The request content type is not supported.",
            ApiDiagnosticCodes.DataInvalidJson => "The data field must contain one valid JSON value.",
            ApiDiagnosticCodes.CultureInvalid => "The culture name is invalid.",
            ApiDiagnosticCodes.SampleCountInvalid =>
                SampleJsonRequestFields.ItemsPerCollection + " must be a positive decimal integer.",
            ApiDiagnosticCodes.SampleValuesInvalid => SampleJsonRequestFields.Values + " must be " +
                SampleJsonRequestFields.Filled + " or " + SampleJsonRequestFields.Empty + ".",
            ApiDiagnosticCodes.FormatterNotFound => "The requested formatter is not registered.",
            ApiDiagnosticCodes.FormatterInvalidValue => "The value cannot be formatted as requested.",
            DiagnosticCodes.TemplateInvalidPackage =>
                "The content is not a readable supported Word or ordinary XLSX package. Re-save the document; for Excel, use an ordinary .xlsx workbook.",
            DiagnosticCodes.TemplateInvalidXml => "A document part contains invalid XML.",
            DiagnosticCodes.TemplateUnsupportedStructure => "The DOCX structure is unsupported.",
            DiagnosticCodes.TemplateInvalidExpression => "A template expression has invalid syntax.",
            DiagnosticCodes.TemplateUnmatchedBrace => "Template braces are not balanced or escaped correctly.",
            DiagnosticCodes.TemplateInvalidContext => "A template path is invalid in its current loop context.",
            DiagnosticCodes.TemplateInvalidStructure => "A template block has an unsupported or ambiguous structure.",
            DiagnosticCodes.TemplateRequirementConflict =>
                "The same dependency path is used with incompatible requirement types.",
            DiagnosticCodes.TemplateTrackedRevision =>
                "Tracked revisions must be accepted or rejected before processing.",
            DiagnosticCodes.GenerationInvalidData => "The data is incompatible with the template.",
            DiagnosticCodes.GenerationInvalidFormat => "A value cannot be formatted as requested.",
            DiagnosticCodes.XlsxInvalidPrintArea => "Set one finite rectangular PrintArea on every worksheet.",
            DiagnosticCodes.XlsxUnsupportedFeature =>
                "The XLSX feature cannot be preserved safely; remove or re-author the indicated structure.",
            DiagnosticCodes.XlsxDimensionLimit =>
                "The generated report exceeds Excel worksheet dimensions; reduce the report size.",
            DiagnosticCodes.XlsxUnsafeNumber =>
                "The value cannot be stored faithfully as an Excel number. Supply it as a JSON string or render it with an explicit text formatter.",
            DiagnosticCodes.XlsxInvalidPackage =>
                "The XLSX package is inconsistent. Repair or re-save the workbook with valid parts and relationships.",
            DiagnosticCodes.XlsxUnsafeFormula =>
                "The XLSX formula reference cannot be transformed safely; re-author it using explicit retained A1 operands.",
            DiagnosticCodes.SampleElementLimit => $"The sample exceeds the {SampleJsonLimits.MaximumArrayElements:N0} element limit; reduce itemsPerCollection.",
            DiagnosticCodes.SampleIncompatibleDemand => "Shared template collections have incompatible sample requirements.",
            _ => "The server could not complete the operation."
        };
    }
}
