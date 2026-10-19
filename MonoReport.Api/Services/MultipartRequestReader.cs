using Microsoft.Net.Http.Headers;

namespace MonoReport.Api.Services;

public sealed class MultipartRequestReader
{
    public async Task<IFormCollection> Read(HttpRequest request, bool analysis, CancellationToken cancellation)
    {
        if (!MediaTypeHeaderValue.TryParse(request.ContentType, out var media) ||
            !string.Equals(media.MediaType.Value, "multipart/form-data", StringComparison.OrdinalIgnoreCase))
            throw new RequestInputException(415, ApiDiagnosticCodes.RequestUnsupportedMediaType);
        if (!media.Boundary.HasValue)
            throw new RequestInputException(400, ApiDiagnosticCodes.RequestInvalidMultipart);
        IFormCollection form;
        try { form = await request.ReadFormAsync(cancellation); }
        catch (IOException error) when (ContainsStatus(error, StatusCodes.Status413PayloadTooLarge))
        {
            throw new RequestInputException(413, ApiDiagnosticCodes.RequestTooLarge);
        }
        catch (Exception error) when (error is InvalidDataException or IOException)
        {
            throw new RequestInputException(400, ApiDiagnosticCodes.RequestInvalidMultipart);
        }
        RequireFile(form, "template");
        if (analysis && (form.Keys.Any(key => key != "template") || form.Files.Any(f => f.Name != "template")))
            throw new RequestInputException(400, ApiDiagnosticCodes.RequestUnexpectedInput);
        if (!analysis)
        {
            RequireText(form, "data", true);
            RequireText(form, "culture", false);
        }
        return form;
    }

    private static bool ContainsStatus(Exception error, int status)
    {
        for (Exception? current = error; current is not null; current = current.InnerException)
            if (current is BadHttpRequestException requestError && requestError.StatusCode == status)
                return true;
        return false;
    }

    private static void RequireFile(IFormCollection form, string name)
    {
        var files = form.Files.Where(f => f.Name == name).ToArray();
        var count = files.Length + (form.TryGetValue(name, out var values) ? values.Count : 0);
        if (count == 0) throw new RequestInputException(400, ApiDiagnosticCodes.RequestMissingInput, name);
        if (count > 1) throw new RequestInputException(400, ApiDiagnosticCodes.RequestDuplicateInput, name);
        if (files.Length != 1) throw new RequestInputException(400, ApiDiagnosticCodes.RequestInvalidInput, name);
        if (files[0].Length == 0) throw new RequestInputException(400, ApiDiagnosticCodes.RequestEmptyInput, name);
    }

    private static void RequireText(IFormCollection form, string name, bool required)
    {
        var files = form.Files.Count(f => f.Name == name);
        var values = form[name];
        var count = files + values.Count;
        if (required && count == 0) throw new RequestInputException(400, ApiDiagnosticCodes.RequestMissingInput, name);
        if (count > 1) throw new RequestInputException(400, ApiDiagnosticCodes.RequestDuplicateInput, name);
        if (files != 0) throw new RequestInputException(400, ApiDiagnosticCodes.RequestInvalidInput, name);
        if (required && string.IsNullOrWhiteSpace(values.ToString())) throw new RequestInputException(400, ApiDiagnosticCodes.RequestEmptyInput, name);
    }
}
