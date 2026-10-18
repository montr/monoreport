namespace MonoReport.Models;

internal sealed record StructuralDiagnostic(
    string Code,
    string Detail,
    string? Field = null,
    string? FieldCode = null,
    string? Location = null,
    string? Formatter = null,
    long SourceOrder = long.MaxValue,
    string? SourceContext = null,
    string? Explanation = null)
{
    internal TemplateException ToException() => new(this);
}
