namespace MonoReport.Models
{
    public sealed record MergeDiagnostic(
        string Message,
        string Field,
        string? FieldCode,
        string? Location,
        string? Part = null,
        string? DataPath = null)
    {
        [System.Text.Json.Serialization.JsonIgnore(Condition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull)]
        public string? Code { get; init; }
    }
}
