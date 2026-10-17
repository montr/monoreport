using System.Text.Json.Serialization;

namespace MonoReport.Models;

public sealed record TemplateRequirement
{
    public TemplateRequirement(string name, string type, IEnumerable<TemplateRequirement>? fields = null)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Type = type ?? throw new ArgumentNullException(nameof(type));
        Fields = fields is null ? null : Array.AsReadOnly(fields.ToArray());
    }

    public string Name { get; }
    public string Type { get; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<TemplateRequirement>? Fields { get; }
}
