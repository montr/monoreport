using System.Text.Json;
using MonoReport.Models;

namespace MonoReport.Services;

internal enum ResolutionState { Missing, Null, Present }

internal sealed record DataResolution(ResolutionState State, JsonElement Value);

public sealed class DataResolver
{
    internal string DataPath(TemplatePath path, MergeContext context)
    {
        string basis;
        if (path.Origin == TemplatePathOrigin.Root) basis = "$";
        else
        {
            var paths = context.DataPaths.ToArray();
            basis = path.ParentDepth < paths.Length ? paths[path.ParentDepth] : "$";
        }
        return path.CurrentElement || path.Segments.Count == 0
            ? basis
            : basis + "." + string.Join('.', path.Segments);
    }

    internal DataResolution Resolve(TemplatePath path, MergeContext context)
    {
        JsonElement value;
        if (path.Origin == TemplatePathOrigin.Root) value = context.Root;
        else
        {
            var values = context.Values.ToArray();
            value = path.ParentDepth < values.Length ? values[path.ParentDepth] : context.Root;
        }
        if (path.CurrentElement) return State(value);
        foreach (var segment in path.Segments)
        {
            if (value.ValueKind != JsonValueKind.Object || !value.TryGetProperty(segment, out value))
                return new(ResolutionState.Missing, default);
        }
        return State(value);
    }

    private static DataResolution State(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Undefined => new(ResolutionState.Missing, value),
        JsonValueKind.Null => new(ResolutionState.Null, value),
        _ => new(ResolutionState.Present, value)
    };
}
