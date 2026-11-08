using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json.Nodes;
using MonoReport.Core.Formatting;

namespace MonoReport;

internal static class SampleJsonScalarValue
{
    internal static SampleSemanticResolution Resolve(IEnumerable<SemanticType?> types, bool structuralBoolean = false)
    {
        var known = types.Where(type => type is not null).ToArray();
        var selected = known.FirstOrDefault();
        var conflict = known.Distinct().Skip(1).Any() ||
                       (structuralBoolean && known.Any(type => type != SemanticType.Boolean));
        return new SampleSemanticResolution(structuralBoolean ? SemanticType.Boolean : selected, conflict);
    }

    internal static JsonNode? CreateValue(SampleJsonOptions options, SemanticType? semanticType, string name,
        int ordinal)
    {
        if (options.Values == SampleJsonValues.Empty)
            return semanticType == SemanticType.Boolean ? JsonValue.Create(false) : JsonValue.Create(string.Empty);

        return semanticType switch
        {
            SemanticType.Number => JsonValue.Create(42.5m + Math.Max(0, ordinal - 1)),
            SemanticType.Date => JsonValue.Create(new DateOnly(2024, 1, 15)
                .AddDays(Math.Max(0, ordinal - 1)).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)),
            SemanticType.Boolean => JsonValue.Create(ordinal % 2 == 1),
            _ => JsonValue.Create("Sample " + (string.IsNullOrEmpty(name) ? "value" : name) + " " + ordinal)
        };
    }
}

internal sealed record SampleSemanticResolution(SemanticType? SemanticType, bool HasConflict);
