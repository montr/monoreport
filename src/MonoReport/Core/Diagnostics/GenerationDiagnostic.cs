namespace MonoReport.Core.Diagnostics;

public sealed record GenerationDiagnostic(
    string Code,
    string Message,
    string? Field = null,
    string? Part = null,
    string? Location = null,
    string? Formatter = null,
    string? DataPath = null);