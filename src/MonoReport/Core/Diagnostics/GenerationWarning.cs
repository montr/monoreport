using System.Text.Json.Serialization;

namespace MonoReport.Core.Diagnostics;

public sealed record GenerationWarning(
    string Message,
    string Field,
    string? FieldCode,
    string? Location,
    string? Part = null,
    string? DataPath = null)
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Code { get; init; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? RemovedCharacterCount { get; init; }
}
