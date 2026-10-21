namespace MonoReport.Models.Xlsx;

internal sealed class XlsxOperationContext(MergeOptions options)
{
    internal MergeOptions Options { get; } = options ?? throw new ArgumentNullException(nameof(options));
    internal List<MergeDiagnostic> Warnings { get; } = [];
    internal List<GenerationDiagnostic> Errors { get; } = [];
    internal long SourceCells { get; set; }
    internal long SourceBytes { get; set; }
    internal long PartBytesRead { get; set; }
    internal long SourceFormulas { get; set; }
    internal long ResolvedFields { get; set; }
    internal long OutputCells { get; set; }
    internal bool LimitReached => Errors.Count == GenerationDiagnosticPolicy.ErrorLimit;
    private TemplateException? _first;

    internal void Report(TemplateException error, string? dataPath = null)
    {
        if (LimitReached) return;
        _first ??= error;
        Errors.Add(error.ToDiagnostic(dataPath));
    }

    internal void ThrowIfErrors()
    {
        if (_first is not null) throw new TemplateException(_first, Errors, LimitReached);
    }
}
