namespace MonoReport.Api.Services;

using MonoReport.Models;

public sealed class ApiProblemFactory
{
    public IResult Create(HttpContext context, int status, string code, string? field = null) => Results.Problem(
        statusCode: status, type: $"urn:monoreport:error:{code}", title: status switch
        {
            400 => "Invalid request", 413 => "Request too large", 415 => "Unsupported media type",
            422 => "Document generation failed", _ => "Internal server error"
        }, detail: SafeMessage(code), extensions: new Dictionary<string, object?>
        {
            ["code"] = code, ["traceId"] = context.TraceIdentifier, ["field"] = field
        });

    public IResult Create(HttpContext context, TemplateException error)
    {
        var extensions = new Dictionary<string, object?>
        {
            ["code"] = error.Code,
            ["traceId"] = context.TraceIdentifier,
            ["field"] = error.Field,
            ["errors"] = error.Errors,
            ["limitReached"] = error.LimitReached
        };
        return Results.Problem(statusCode: 422, type: $"urn:monoreport:error:{error.Code}",
            title: "Document generation failed", detail: SafeMessage(error.Code), extensions: extensions);
    }

    public string SafeMessage(string code) => code switch
    {
        ApiDiagnosticCodes.RequestMissingInput => "A required input is missing.",
        ApiDiagnosticCodes.RequestEmptyInput => "A required input is empty.",
        ApiDiagnosticCodes.RequestDuplicateInput => "An input was supplied more than once.",
        ApiDiagnosticCodes.RequestInvalidInput => "An input has the wrong representation.",
        ApiDiagnosticCodes.RequestUnexpectedInput => "Validation accepts only the template file.",
        ApiDiagnosticCodes.RequestInvalidMultipart => "The multipart request is malformed.",
        ApiDiagnosticCodes.RequestTooLarge => "The whole request exceeds the configured size limit.",
        ApiDiagnosticCodes.RequestUnsupportedMediaType => "Use multipart/form-data.",
        ApiDiagnosticCodes.DataInvalidJson => "The data field must contain one valid JSON value.",
        ApiDiagnosticCodes.CultureInvalid => "The culture name is invalid.",
        DiagnosticCodes.TemplateInvalidPackage => "The content is not a readable DOCX package.",
        DiagnosticCodes.TemplateInvalidXml => "A document part contains invalid XML.",
        DiagnosticCodes.TemplateUnsupportedStructure => "The DOCX structure is unsupported.",
        DiagnosticCodes.TemplateInvalidExpression => "A template expression has invalid syntax.",
        DiagnosticCodes.TemplateUnmatchedBrace => "Template braces are not balanced or escaped correctly.",
        DiagnosticCodes.TemplateInvalidContext => "A template path is invalid in its current loop context.",
        DiagnosticCodes.TemplateInvalidStructure => "A template block has an unsupported or ambiguous structure.",
        DiagnosticCodes.TemplateRequirementConflict => "The same dependency path is used with incompatible requirement types.",
        DiagnosticCodes.TemplateTrackedRevision => "Tracked revisions must be accepted or rejected before processing.",
        DiagnosticCodes.GenerationInvalidData => "The data is incompatible with the template.",
        DiagnosticCodes.GenerationInvalidFormat => "A value cannot be formatted as requested.",
        _ => "The server could not complete the operation."
    };
}
