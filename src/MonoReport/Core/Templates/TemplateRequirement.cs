using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace MonoReport.Core.Templates;

public sealed record TemplateRequirement
{
    public const string ObjectType = "object";
    public const string ArrayType = "array";
    public const string ValueType = "value";
    public const string BooleanType = "boolean";

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
