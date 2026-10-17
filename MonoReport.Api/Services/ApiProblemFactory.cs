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
        "request.missing_input" => "A required input is missing.",
        "request.empty_input" => "A required input is empty.",
        "request.duplicate_input" => "An input was supplied more than once.",
        "request.invalid_input" => "An input has the wrong representation.",
        "request.unexpected_input" => "Validation accepts only the template file.",
        "request.invalid_multipart" => "The multipart request is malformed.",
        "request.too_large" => "The whole request exceeds the configured size limit.",
        "request.unsupported_media_type" => "Use multipart/form-data.",
        "data.invalid_json" => "The data field must contain one valid JSON value.",
        "culture.invalid" => "The culture name is invalid.",
        "template.invalid_package" => "The content is not a readable DOCX package.",
        "template.invalid_xml" => "A document part contains invalid XML.",
        "template.unsupported_structure" => "The DOCX structure is unsupported.",
        "template.invalid_expression" => "A template expression has invalid syntax.",
        "template.unmatched_brace" => "Template braces are not balanced or escaped correctly.",
        "template.invalid_context" => "A template path is invalid in its current loop context.",
        "template.invalid_structure" => "A template block has an unsupported or ambiguous structure.",
        "template.requirement_conflict" => "The same dependency path is used with incompatible requirement types.",
        "template.tracked_revision" => "Tracked revisions must be accepted or rejected before processing.",
        "generation.invalid_data" => "The data is incompatible with the template.",
        "generation.invalid_format" => "A value cannot be formatted as requested.",
        _ => "The server could not complete the operation."
    };
}
