using System.Text.Json;
using MonoReport.Models;

namespace MonoReport.Services;

internal enum ResolutionState { Missing, Null, Present }

internal sealed record DataResolution(ResolutionState State, JsonElement Value);

public sealed class DataResolver
{
    internal string DataPath(TemplatePath path, MergeContext context)
        => DataPath(path, new DataScope(context.Root, context.Values.ToArray(), context.DataPaths.ToArray()));

    internal string DataPath(TemplatePath path, DataScope scope)
    {
        string basis;
        if (path.Origin == TemplatePathOrigin.Root) basis = "$";
        else
        {
            basis = path.ParentDepth < scope.Paths.Count ? scope.Paths[path.ParentDepth] : "$";
        }
        return path.CurrentElement || path.Segments.Count == 0
            ? basis
            : basis + "." + string.Join('.', path.Segments);
    }

    internal DataResolution Resolve(TemplatePath path, MergeContext context)
        => Resolve(path, new DataScope(context.Root, context.Values.ToArray(), context.DataPaths.ToArray()));

    internal DataResolution Resolve(TemplatePath path, DataScope scope)
    {
        JsonElement value;
        if (path.Origin == TemplatePathOrigin.Root) value = scope.Root;
        else
        {
            value = path.ParentDepth < scope.Values.Count ? scope.Values[path.ParentDepth] : scope.Root;
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
