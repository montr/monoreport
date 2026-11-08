using System.Text.Json.Serialization;

namespace MonoReport.Core.Templates;

public sealed record TemplateOccurrence(
    string Kind,
    string Expression,
    string Path,
    string Part,
    string Location)
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SemanticType { get; init; }
}
