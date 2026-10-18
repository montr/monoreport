using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;

namespace MonoReport.Services;

public sealed class TemplateAnalyzer(DocxStoryReader storyReader, DocxTemplatePlanner planner)
{
    private readonly DocxStoryReader _storyReader = storyReader ?? throw new ArgumentNullException(nameof(storyReader));
    private readonly DocxTemplatePlanner _planner = planner ?? throw new ArgumentNullException(nameof(planner));

    public TemplateAnalysisResult Analyze(Stream template)
    {
        ArgumentNullException.ThrowIfNull(template);
        if (!template.CanRead) throw new ArgumentException("Template stream must be readable.", nameof(template));
        using var buffer = new MemoryStream();
        template.CopyTo(buffer);
        buffer.Position = 0;
        var document = _storyReader.OpenForValidation(buffer, false, out var packageDiagnostic);
        if (packageDiagnostic is not null) return new([Issue(packageDiagnostic)]);
        using (document)
        {
            var inspection = Inspect(document!);
            return new(inspection.Diagnostics.Select(Issue), inspection.Warnings,
                inspection.Requirements, inspection.Occurrences);
        }
    }

    internal TemplateInspection Inspect(WordprocessingDocument document)
    {
        var storyInspection = _storyReader.Inspect(document);
        var stories = storyInspection.Stories.Select(story => _planner.Plan(story.Part, story.Xml)).ToArray();
        var identities = _planner.InspectIdentities(storyInspection, stories);
        var builder = new RequirementBuilder();
        foreach (var dependency in stories.SelectMany(story => story.Dependencies)) builder.Add(dependency);
        var diagnostics = storyInspection.Diagnostics.Concat(stories.SelectMany(story => story.Diagnostics))
            .Concat(builder.Diagnostics).Concat(identities.Diagnostics)
            .OrderBy(item => DiagnosticPartOrder(item, storyInspection.PartOrder))
            .ThenBy(item => item.SourceOrder)
            .ToArray();
        return new(stories, diagnostics, [], builder.Requirements, builder.Occurrences, storyInspection, identities);
    }

    private static int DiagnosticPartOrder(StructuralDiagnostic diagnostic,
        IReadOnlyDictionary<string, int> partOrder)
    {
        var part = diagnostic.Location?.Split(',', 2)[0];
        return part is not null && partOrder.TryGetValue(part, out var order) ? order : int.MaxValue;
    }

    private static TemplateValidationIssue Issue(StructuralDiagnostic diagnostic)
    {
        var part = diagnostic.Location?.Split(',', 2)[0];
        return new(diagnostic.Code, diagnostic.ToException().ToDiagnostic(null).Message, diagnostic.Field, part, diagnostic.Location);
    }

    private sealed class RequirementBuilder
    {
        private readonly Scope _root = new();
        private readonly List<TemplateOccurrence> _occurrences = [];

        internal IReadOnlyList<TemplateRequirement> Requirements => _root.Items.Select(ToPublic).ToArray();
        internal IReadOnlyList<TemplateOccurrence> Occurrences => _occurrences;
        internal IReadOnlyList<StructuralDiagnostic> Diagnostics => Conflicted(_root)
            .SelectMany(node => node.Uses.Select(use => new StructuralDiagnostic(
                "template.requirement_conflict",
                $"Dependency '{node.Path}' has incompatible requirement types.", use.Expression.Source,
                use.Expression.Source, use.Location, SourceOrder: use.SourceOrder))).ToArray();

        internal void Add(PlannedDependency dependency)
        {
            _occurrences.Add(new(dependency.Kind, dependency.Expression.Source, dependency.CanonicalPath,
                dependency.Part, dependency.Location));
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
                var needed = component.Array ? "array" : last ? dependency.RequirementType : "object";
                if (!scope.ByName.TryGetValue(component.Name, out var node))
                {
                    node = new(component.Name, needed, path);
                    scope.Add(node);
                }
                node.Uses.Add(dependency);
                if (!Compatible(node.Type, needed))
                {
                    node.Conflicted = true;
                    return;
                }
                else if (node.Type == "value" && needed == "boolean") node.Type = "boolean";
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

        private static bool Compatible(string current, string needed) => current == needed ||
            current is "value" or "boolean" && needed is "value" or "boolean";

        private static IReadOnlyList<(string Name, bool Array)> Components(string path)
        {
            if (path == "$" || !path.StartsWith("$.", StringComparison.Ordinal)) return [];
            return path[2..].Split('.').Select(segment => segment.EndsWith("[]", StringComparison.Ordinal)
                ? (segment[..^2], true) : (segment, false)).ToArray();
        }

        private static TemplateRequirement ToPublic(Node node) =>
            new(node.Name, node.Type, node.Children.Items.Count == 0 ? null : node.Children.Items.Select(ToPublic));

        private sealed class Scope
        {
            internal List<Node> Items { get; } = [];
            internal Dictionary<string, Node> ByName { get; } = new(StringComparer.Ordinal);
            internal void Add(Node node) { Items.Add(node); ByName.Add(node.Name, node); }
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
}
