namespace MonoReport.Api.Services;

/// <summary>Canonical values of HTTP-layer diagnostic-code identifiers published by the API.</summary>
public static class ApiDiagnosticCodes
{
    public const string RequestMissingInput = "request.missing_input";
    public const string RequestEmptyInput = "request.empty_input";
    public const string RequestDuplicateInput = "request.duplicate_input";
    public const string RequestInvalidInput = "request.invalid_input";
    public const string RequestUnexpectedInput = "request.unexpected_input";
    public const string RequestInvalidMultipart = "request.invalid_multipart";
    public const string RequestTooLarge = "request.too_large";
    public const string RequestUnsupportedMediaType = "request.unsupported_media_type";
    public const string DataInvalidJson = "data.invalid_json";
    public const string CultureInvalid = "culture.invalid";
    public const string ServerInternalError = "server.internal_error";
}
