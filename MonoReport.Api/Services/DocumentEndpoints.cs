using System.Globalization;
using System.Text.Json;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Api.Services;

public sealed class DocumentEndpoints(MultipartRequestReader reader, MergeEngine engine, TemplateAnalyzer analyzer, GenerationOptionsFactory options)
{
    private readonly MultipartRequestReader _reader = reader ?? throw new ArgumentNullException(nameof(reader));
    private readonly MergeEngine _engine = engine ?? throw new ArgumentNullException(nameof(engine));
    private readonly TemplateAnalyzer _analyzer = analyzer ?? throw new ArgumentNullException(nameof(analyzer));
    private readonly GenerationOptionsFactory _options = options ?? throw new ArgumentNullException(nameof(options));

    public async Task<IResult> Generate(HttpContext context)
    {
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
            if (result.Warnings.Count != 0) context.Response.Headers["X-MonoReport-Warning-Count"] = result.Warnings.Count.ToString(CultureInfo.InvariantCulture);
            return result.Format == Models.DocumentFormat.Xlsx
                ? Results.File(result.Document, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "document.xlsx")
                : Results.File(result.Document, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "document.docx");
        }
    }

    public async Task<IResult> Analyze(HttpContext context)
    {
        var form = await _reader.Read(context.Request, true, context.RequestAborted);
        using var template = form.Files.GetFile("template")!.OpenReadStream();
        return Results.Json(_analyzer.Analyze(template));
    }
}
