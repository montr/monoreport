using System;
using System.Collections.Generic;
using System.Linq;
using MonoReport.Core.Diagnostics;

namespace MonoReport.Core.Templates;

internal sealed class TemplateRequirementBuilder
{
    private readonly List<TemplateOccurrence> _occurrences = [];
    private readonly HashSet<string> _booleanPaths = new(StringComparer.Ordinal);
    private readonly Scope _root = new();

    internal IReadOnlyList<TemplateRequirement> Requirements => _root.Items.Select(ToPublic).ToArray();
    internal IReadOnlyList<TemplateOccurrence> Occurrences => _occurrences.Select(item =>
        item.SemanticType is null && item.Kind == PlannedDependency.FieldKind && _booleanPaths.Contains(item.Path)
            ? item with {SemanticType = "boolean"}
            : item).ToArray();

    internal IReadOnlyList<StructuralDiagnostic> Diagnostics => Conflicted(_root)
        .SelectMany(node => node.Uses.Select(use => new StructuralDiagnostic(
            DiagnosticCodes.TemplateRequirementConflict,
            $"Dependency '{node.Path}' has incompatible requirement types.", use.Expression.Source,
            use.Expression.Source, use.Location, SourceOrder: use.SourceOrder))).ToArray();

    internal void Add(PlannedDependency dependency)
    {
        if (dependency.Kind == PlannedDependency.IfKind) _booleanPaths.Add(dependency.CanonicalPath);
        var semanticType = dependency.Expression.Formatter?.SemanticType?.ToString().ToLowerInvariant();
        _occurrences.Add(new TemplateOccurrence(dependency.Kind, dependency.Expression.Source, dependency.CanonicalPath,
            dependency.Part, dependency.Location) {SemanticType = semanticType});
        if (dependency.Expression.Path?.CurrentElement == true) return;
        var components = Components(dependency.CanonicalPath);
        if (components.Count == 0) return;
        var scope = _root;
        var path = "$";
        for (var index = 0; index < components.Count; index++)
        {
            var component = components[index];
            path += "." + component.Name + (component.Array ? "[]" : string.Empty);
            var last = index == components.Count - 1;
            var needed = component.Array ? TemplateRequirement.ArrayType :
                last ? dependency.RequirementType : TemplateRequirement.ObjectType;
            if (!scope.ByName.TryGetValue(component.Name, out var node))
            {
                node = new Node(component.Name, needed, path);
                scope.Add(node);
            }

            node.Uses.Add(dependency);
            if (!Compatible(node.Type, needed))
            {
                node.Conflicted = true;
                return;
            }

            if (node.Type == TemplateRequirement.ValueType && needed == TemplateRequirement.BooleanType)
                node.Type = TemplateRequirement.BooleanType;
            if (!last) scope = node.Children;
        }
    }

    private static IEnumerable<Node> Conflicted(Scope scope)
    {
        foreach (var node in scope.Items)
        {
            if (node.Conflicted) yield return node;
            foreach (var child in Conflicted(node.Children)) yield return child;
        }
    }

    private static bool Compatible(string current, string needed)
    {
        return current == needed ||
               (current is TemplateRequirement.ValueType or TemplateRequirement.BooleanType &&
                needed is TemplateRequirement.ValueType or TemplateRequirement.BooleanType);
    }

    private static IReadOnlyList<(string Name, bool Array)> Components(string path)
    {
        if (path == "$" || !path.StartsWith("$.", StringComparison.Ordinal)) return [];
        return path[2..].Split('.').Select(segment => segment.EndsWith("[]", StringComparison.Ordinal)
            ? (segment[..^2], true)
            : (segment, false)).ToArray();
    }

    private static TemplateRequirement ToPublic(Node node)
    {
        return new TemplateRequirement(node.Name, node.Type,
            node.Children.Items.Count == 0 ? null : node.Children.Items.Select(ToPublic));
    }

    private sealed class Scope
    {
        internal List<Node> Items { get; } = [];
        internal Dictionary<string, Node> ByName { get; } = new(StringComparer.Ordinal);

        internal void Add(Node node)
        {
            Items.Add(node);
            ByName.Add(node.Name, node);
        }
    }

    private sealed class Node(string name, string type, string path)
    {
        internal string Name { get; } = name;
        internal string Type { get; set; } = type;
        internal string Path { get; } = path;
        internal Scope Children { get; } = new();
        internal List<PlannedDependency> Uses { get; } = [];
        internal bool Conflicted { get; set; }
    }
}
