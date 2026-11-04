using System.Collections.Generic;
using MonoReport.Core.Diagnostics;

namespace MonoReport;

public sealed record GenerationResult(byte[] Document, IReadOnlyList<GenerationWarning> Warnings)
{
    public DocumentFormat Format { get; init; } = DocumentFormat.Docx;
}