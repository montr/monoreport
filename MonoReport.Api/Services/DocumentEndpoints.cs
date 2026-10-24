using System.Globalization;
using System.Text.Json;
using Microsoft.Extensions.Options;
using Microsoft.Net.Http.Headers;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Api.Services;

public sealed class DocumentEndpoints(MultipartRequestReader reader, MergeEngine engine, TemplateAnalyzer analyzer, GenerationOptionsFactory options,
    IOptions<Microsoft.AspNetCore.Http.Json.JsonOptions> jsonOptions)
{
    private readonly MultipartRequestReader _reader = reader ?? throw new ArgumentNullException(nameof(reader));
    private readonly MergeEngine _engine = engine ?? throw new ArgumentNullException(nameof(engine));
    private readonly TemplateAnalyzer _analyzer = analyzer ?? throw new ArgumentNullException(nameof(analyzer));
    private readonly GenerationOptionsFactory _options = options ?? throw new ArgumentNullException(nameof(options));
    private readonly JsonSerializerOptions _jsonOptions = (jsonOptions ?? throw new ArgumentNullException(nameof(jsonOptions))).Value.SerializerOptions;

    public async Task<IResult> Generate(HttpContext context)
    {
        context.Response.Headers.Append(HeaderNames.Vary, HeaderNames.Accept);
        var form = await _reader.Read(context.Request, false, context.RequestAborted);
        JsonDocument data;
        try { data = JsonDocument.Parse(form["data"].ToString()); }
        catch (JsonException) { throw new RequestInputException(400, ApiDiagnosticCodes.DataInvalidJson, "data"); }
        using (data)
        {
            var culture = CultureInfo.InvariantCulture;
            var name = form["culture"].ToString();
            if (!string.IsNullOrWhiteSpace(name))
            {
                try { culture = CultureInfo.GetCultureInfo(name); }
                catch (CultureNotFoundException) { throw new RequestInputException(400, ApiDiagnosticCodes.CultureInvalid, "culture"); }
            }
            using var template = form.Files.GetFile("template")!.OpenReadStream();
            var result = _engine.Merge(template, data, _options.Create(culture));
            var format = result.Format == Models.DocumentFormat.Xlsx ? "xlsx" : "docx";
            var fileName = "document." + format;
            var contentType = result.Format == Models.DocumentFormat.Xlsx
                ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            IResult response;
            if (WantsMultipart(context.Request))
            {
                var diagnostics = JsonSerializer.SerializeToUtf8Bytes(new { format, fileName, warnings = result.Warnings }, _jsonOptions);
                response = new GenerationMultipartResult(result.Document, contentType, fileName, diagnostics);
            }
            else response = Results.File(result.Document, contentType, fileName);
            if (result.Warnings.Count != 0) context.Response.Headers["X-MonoReport-Warning-Count"] = result.Warnings.Count.ToString(CultureInfo.InvariantCulture);
            return response;
        }
    }

    private static bool WantsMultipart(HttpRequest request)
    {
        if (!MediaTypeHeaderValue.TryParseList(request.Headers.Accept, out var accepted)) return false;
        var multipart = 0d;
        var binary = 0d;
        foreach (var item in accepted)
        {
            var hasQuality = false;
            foreach (var parameter in item.Parameters)
            {
                if (!parameter.Name.Equals("q", StringComparison.OrdinalIgnoreCase)) continue;
                hasQuality = true;
                // Validate the whole token: the framework Quality getter can accept a numeric prefix.
                var token = parameter.Value.AsSpan();
                if (token.IsEmpty || token[0] is not ('0' or '1')
                    || (token.Length > 1 && (token[1] != '.'
                        || token[2..].ContainsAnyExceptInRange('0', '9')
                        || (token[0] == '1' && token[2..].ContainsAnyExcept('0')))))
                    return false;
            }
            var parsedQuality = item.Quality;
            if (hasQuality && (parsedQuality is null || !double.IsFinite(parsedQuality.Value) || parsedQuality.Value is < 0 or > 1))
                return false;
            var quality = parsedQuality ?? 1d;
            if (item.MediaType.Equals("multipart/mixed", StringComparison.OrdinalIgnoreCase)) multipart = Math.Max(multipart, quality);
            else if (item.MediaType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", StringComparison.OrdinalIgnoreCase)
                || item.MediaType.Equals("application/vnd.openxmlformats-officedocument.wordprocessingml.document", StringComparison.OrdinalIgnoreCase))
                binary = Math.Max(binary, quality);
        }
        return multipart > binary;
    }

    public async Task<IResult> Analyze(HttpContext context)
    {
        var form = await _reader.Read(context.Request, true, context.RequestAborted);
        using var template = form.Files.GetFile("template")!.OpenReadStream();
        return Results.Json(_analyzer.Analyze(template));
    }
}
