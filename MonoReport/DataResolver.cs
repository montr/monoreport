using System.Collections.Generic;
using System.Text.Json;
using MonoReport.Core.Templates;

namespace MonoReport;

internal enum ResolutionState
{
    Missing,
    Null,
    Present
}

internal sealed record DataResolution(ResolutionState State, JsonElement Value);

// Current value/path first, followed by logical ancestors; neither document format owns this input.
internal sealed record DataScope(JsonElement Root, IReadOnlyList<JsonElement> Values, IReadOnlyList<string> Paths);

public sealed class DataResolver
{
    internal string DataPath(TemplatePath path, DataScope scope)
    {
        string basis;
        if (path.Origin == TemplatePathOrigin.Root) basis = "$";
        else
            basis = path.ParentDepth < scope.Paths.Count ? scope.Paths[path.ParentDepth] : "$";
        return path.CurrentElement || path.Segments.Count == 0
            ? basis
            : basis + "." + string.Join('.', path.Segments);
    }

    internal DataResolution Resolve(TemplatePath path, DataScope scope)
    {
        JsonElement value;
        if (path.Origin == TemplatePathOrigin.Root) value = scope.Root;
        else
            value = path.ParentDepth < scope.Values.Count ? scope.Values[path.ParentDepth] : scope.Root;
        if (path.CurrentElement) return State(value);
        foreach (var segment in path.Segments)
            if (value.ValueKind != JsonValueKind.Object || !value.TryGetProperty(segment, out value))
                return new DataResolution(ResolutionState.Missing, default);
        return State(value);
    }

    private static DataResolution State(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.Undefined => new DataResolution(ResolutionState.Missing, value),
            JsonValueKind.Null => new DataResolution(ResolutionState.Null, value),
            _ => new DataResolution(ResolutionState.Present, value)
        };
    }
}