namespace MonoReport.Models
{
    public sealed record MergeDiagnostic(
        string Message,
        string Field,
        string? FieldCode,
        string? Location,
        string? Part = null,
        string? DataPath = null);
}
