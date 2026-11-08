using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Formatting;
using MonoReport.Core.Templates;
using MonoReport.Docx.Planning;

namespace MonoReport.Docx;

public sealed class DocxSampleJsonGenerator(DocxStoryReader storyReader, DocxTemplateAnalyzer analyzer)
{
    private const int ElementLimit = SampleJsonLimits.MaximumArrayElements;
    private const string UnresolvedShape = DiagnosticCodes.SampleUnresolvedShape;
    private const string SemanticConflict = DiagnosticCodes.SampleSemanticConflict;
    private const string ElementLimitCode = DiagnosticCodes.SampleElementLimit;
    private readonly DocxStoryReader _storyReader = storyReader ?? throw new ArgumentNullException(nameof(storyReader));
    private readonly DocxTemplateAnalyzer _analyzer = analyzer ?? throw new ArgumentNullException(nameof(analyzer));

    internal SampleJsonResult Generate(System.IO.Stream package, SampleJsonOptions options)
    {
        var document = _storyReader.OpenForValidation(package, false, out var packageDiagnostic);
        if (packageDiagnostic is not null) throw packageDiagnostic.ToException();
        using (document)
        {
            var inspection = _analyzer.Inspect(document!);
            if (inspection.Diagnostics.Count != 0)
            {
                var errors = inspection.Diagnostics.Select(item => item.ToException().ToDiagnostic(null)).ToArray();
                throw new TemplateException(inspection.Diagnostics[0].ToException(), errors, false);
            }

            var model = SampleModel.Create(inspection.Stories.SelectMany(story => story.Dependencies));
            var warnings = new List<GenerationDiagnostic>();
            var count = model.CountElements(options.ItemsPerCollection);
            if (count > ElementLimit)
                throw Failure(ElementLimitCode,
                    $"The sample would contain more than {ElementLimit:N0} array elements. Reduce ItemsPerCollection.");
            var data = model.Materialize(options, warnings, "$") as JsonObject ?? new JsonObject();
            using var json = JsonDocument.Parse(data.ToJsonString());
            return new SampleJsonResult(json.RootElement, warnings);
        }
    }

    private static TemplateException Failure(string code, string message, string? location = null)
    {
        var error = new TemplateException(message, location: location, code: code);
        return new TemplateException(error,
            [new GenerationDiagnostic(code, message, Part: location?.Split(',', 2)[0], Location: location)], false);
    }

    private sealed class SampleModel
    {
        private readonly Node _root = new("$", TemplateRequirement.ObjectType, "$");

        internal static SampleModel Create(IEnumerable<PlannedDependency> dependencies)
        {
            var model = new SampleModel();
            foreach (var dependency in dependencies) model.Add(dependency);
            return model;
        }

        internal BigInteger CountElements(BigInteger requested)
        {
            return _root.CountElements(requested, 1, SampleJsonLimits.PreflightSentinel);
        }

        internal JsonNode? Materialize(SampleJsonOptions options, List<GenerationDiagnostic> warnings, string path)
        {
            return _root.Materialize(options, warnings, new HashSet<WarningKey>(), path);
        }

        private void Add(PlannedDependency dependency)
        {
            var components = Components(dependency.CanonicalPath);
            if (components.Count == 0) return;
            var parent = _root;
            var canonical = "$";
            foreach (var (name, array) in components)
            {
                canonical += "." + name + (array ? "[]" : string.Empty);
                var node = parent.GetOrAdd(name,
                    array ? TemplateRequirement.ArrayType : TemplateRequirement.ObjectType, canonical);
                parent = array ? node.ArrayItem ?? node : node;
            }

            var target = parent;
            if (dependency.Kind == PlannedDependency.ForKind)
            {
                target.Type = TemplateRequirement.ArrayType;
                target.ArrayItem ??= new Node(string.Empty, TemplateRequirement.ObjectType,
                    dependency.CanonicalPath + "[]");
            }

            if (dependency.RequirementType == TemplateRequirement.BooleanType)
                parent.StructuralBoolean = true;
            if (dependency.Kind == PlannedDependency.FieldKind)
            {
                var semanticType = dependency.Expression.Formatter?.SemanticType;
                parent.Values.Add(new ValueUse(semanticType, dependency.Location,
                    dependency.Expression.Source));
                if (dependency.Expression.Path?.CurrentElement == true) parent.CurrentElement = true;
            }
            if (dependency.Kind == PlannedDependency.ForKind) target.ArrayLocation ??= dependency.Location;
        }

        private static IReadOnlyList<(string Name, bool Array)> Components(string path)
        {
            if (path == "$" || !path.StartsWith("$.", StringComparison.Ordinal)) return [];
            return path[2..].Split('.').Select(segment => segment.EndsWith("[]", StringComparison.Ordinal)
                ? (segment[..^2], true)
                : (segment, false)).ToArray();
        }
    }

    private sealed class Node(string name, string type, string path)
    {
        private readonly Dictionary<string, Node> _childrenByName = new(StringComparer.Ordinal);
        private readonly List<Node> _children = [];
        internal string Name { get; } = name;
        internal string Type { get; set; } = type;
        internal bool StructuralBoolean { get; set; }
        internal bool CurrentElement { get; set; }
        internal string? ArrayLocation { get; set; }
        internal List<ValueUse> Values { get; } = [];
        internal Node? ArrayItem { get; set; }

        internal Node GetOrAdd(string childName, string childType, string childPath)
        {
            if (_childrenByName.TryGetValue(childName, out var current)) return current;
            current = new Node(childName, childType, childPath);
            _childrenByName.Add(childName, current);
            _children.Add(current);
            if (childType == TemplateRequirement.ArrayType)
                current.ArrayItem = new Node(string.Empty, TemplateRequirement.ObjectType, childPath + "[]");
            return current;
        }

        internal BigInteger CountElements(BigInteger requested, BigInteger instances, int limit)
        {
            if (Type == TemplateRequirement.ArrayType)
            {
                if (ArrayItem is null || (!ArrayItem.CurrentElement && ArrayItem._children.Count == 0))
                    return BigInteger.Zero;
                var thisCount = requested * instances;
                if (thisCount >= limit) return limit;
                var nested = ArrayItem?.CountElements(requested, requested * instances, limit - (int)thisCount) ?? BigInteger.Zero;
                return BigInteger.Min(limit, thisCount + nested);
            }

            var total = BigInteger.Zero;
            foreach (var child in _children)
            {
                total += child.CountElements(requested, instances, limit - (int)BigInteger.Min(total, limit));
                if (total >= limit) return limit;
            }
            return total;
        }

        internal JsonNode? Materialize(SampleJsonOptions options, List<GenerationDiagnostic> warnings,
            HashSet<WarningKey> warned, string jsonPath)
        {
            if (Type == TemplateRequirement.ArrayType)
            {
                var array = new JsonArray();
                if (ArrayItem is null || (!ArrayItem.CurrentElement && ArrayItem._children.Count == 0))
                {
                    var location = ArrayLocation;
                    AddWarning(warnings, warned, new GenerationDiagnostic(UnresolvedShape,
                        $"The item shape for collection '{path}' cannot be inferred, so the sample collection is empty.",
                        Part: Part(location), Location: location));
                    return array;
                }

                var item = ArrayItem!;
                var primitive = item.CurrentElement && item._children.Count == 0;
                if (item.CurrentElement && item._children.Count != 0)
                    AddWarning(warnings, warned, new GenerationDiagnostic(UnresolvedShape,
                        $"Collection '{path}' uses both whole-item and property expressions; the sample uses the property shape.",
                        Part: Part(item.Values.FirstOrDefault()?.Location), Location: item.Values.FirstOrDefault()?.Location));
                var count = (int) options.ItemsPerCollection;
                for (var index = 0; index < count; index++)
                {
                    if (primitive)
                    {
                        array.Add(item.Scalar(item.Values, options, warnings, warned,
                            jsonPath + "[" + index + "]", Name));
                    }
                    else
                    {
                        array.Add(item.Materialize(options, warnings, warned, jsonPath + "[" + index + "]"));
                    }
                }
                return array;
            }

            if ((Values.Count != 0 || StructuralBoolean) && _children.Count == 0)
                return Scalar(Values, options, warnings, warned, jsonPath, Name);
            var result = new JsonObject();
            foreach (var child in _children)
                result[child.Name] = child.Materialize(options, warnings, warned, jsonPath + "." + child.Name);
            return result;
        }

        private JsonNode? Scalar(List<ValueUse> uses, SampleJsonOptions options,
            List<GenerationDiagnostic> warnings, HashSet<WarningKey> warned, string jsonPath, string fieldName)
        {
            var ordered = uses.ToArray();
            var resolution = SampleJsonScalarValue.Resolve(ordered.Select(use => use.SemanticType), StructuralBoolean);
            if (resolution.HasConflict)
            {
                var firstKnown = ordered.FirstOrDefault(use => use.SemanticType is not null);
                var location = firstKnown?.Location;
                AddWarning(warnings, warned, new GenerationDiagnostic(SemanticConflict,
                    $"Conflicting semantic hints for '{path}' use the first source hint, with structural Boolean requirements taking precedence.",
                    firstKnown?.Expression, Part(location), location, DataPath: jsonPath), path);
            }

            var itemIndex = jsonPath.LastIndexOf('[');
            var itemEnd = jsonPath.IndexOf(']', itemIndex + 1);
            var ordinal = itemIndex < 0 || itemEnd < 0 ? 1 :
                int.Parse(jsonPath[(itemIndex + 1)..itemEnd], System.Globalization.CultureInfo.InvariantCulture) + 1;
            return SampleJsonScalarValue.CreateValue(options, resolution.SemanticType, fieldName,
                ordinal);
        }

        private static string? Part(string? location) => location?.Split(',', 2)[0];

        private static void AddWarning(List<GenerationDiagnostic> warnings,
            HashSet<WarningKey> warned,
            GenerationDiagnostic warning,
            string? logicalPath = null)
        {
            var key = new WarningKey(warning.Code, warning.Location, logicalPath ?? warning.DataPath,
                warning.Field, logicalPath is null ? warning.Message : string.Empty);
            if (warned.Add(key)) warnings.Add(warning);
        }
    }

    private readonly record struct WarningKey(string Code, string? Location, string? DataPath, string? Field,
        string Message);

    private sealed record ValueUse(SemanticType? SemanticType, string Location,
        string Expression);
}
