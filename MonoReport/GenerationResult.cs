using MonoReport.Core.Diagnostics;
using System.Collections.Generic;

namespace MonoReport;

public sealed record GenerationResult(byte[] Document, IReadOnlyList<GenerationWarning> Warnings)
{
    public DocumentFormat Format { get; init; } = DocumentFormat.Docx;
}