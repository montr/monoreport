using System.Text;

namespace MonoReport.Api.Services;

// All fallible generation and JSON preparation precedes construction. Only transport writes remain.
internal sealed class GenerationMultipartResult : IResult
{
    private readonly byte[] _document;
    private readonly byte[] _diagnostics;
    private readonly string _boundary = "monoreport-" + Guid.NewGuid().ToString("N");
    private readonly byte[] _documentHeader;
    private readonly byte[] _diagnosticsHeader;
    private readonly byte[] _closing;

    internal GenerationMultipartResult(byte[] document, string contentType, string fileName, byte[] diagnostics)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        _documentHeader = Encoding.ASCII.GetBytes($"--{_boundary}\r\nContent-Type: {contentType}\r\nContent-Disposition: attachment; filename=\"{fileName}\"\r\n\r\n");
        _diagnosticsHeader = Encoding.ASCII.GetBytes($"\r\n--{_boundary}\r\nContent-Type: application/json\r\nContent-Disposition: attachment; filename=\"diagnostics.json\"\r\n\r\n");
        _closing = Encoding.ASCII.GetBytes($"\r\n--{_boundary}--\r\n");
    }

    public async Task ExecuteAsync(HttpContext context)
    {
        context.Response.StatusCode = StatusCodes.Status200OK;
        context.Response.ContentType = $"multipart/mixed; boundary=\"{_boundary}\"";
        context.Response.ContentLength = (long)_documentHeader.Length + _document.Length + _diagnosticsHeader.Length + _diagnostics.Length + _closing.Length;
        var body = context.Response.Body;
        var cancellation = context.RequestAborted;
        await body.WriteAsync(_documentHeader, cancellation);
        await body.WriteAsync(_document, cancellation);
        await body.WriteAsync(_diagnosticsHeader, cancellation);
        await body.WriteAsync(_diagnostics, cancellation);
        await body.WriteAsync(_closing, cancellation);
    }
}
