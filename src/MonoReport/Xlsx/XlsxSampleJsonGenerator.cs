using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Xml;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Formatting;
using MonoReport.Core.Templates;
using MonoReport.Xlsx.Planning;

namespace MonoReport.Xlsx;

public sealed class XlsxSampleJsonGenerator(XlsxTemplateAnalyzer analyzer, XlsxTemplatePlanner planner)
{
    private const int ElementLimit = SampleJsonLimits.MaximumArrayElements;
    private const string LevelProperty = "$level";
    private const string PathProperty = "$path";
    private readonly XlsxTemplateAnalyzer _analyzer = analyzer ?? throw new ArgumentNullException(nameof(analyzer));
    private readonly XlsxTemplatePlanner _planner = planner ?? throw new ArgumentNullException(nameof(planner));

    internal SampleJsonResult Generate(Stream package, SampleJsonOptions options)
    {
        try
        {
            using var document = SpreadsheetDocument.Open(package, false);
            var plan = _analyzer.Inspect(document, package.Length);
            return new Builder(plan, _planner, options).Build();
        }
        catch (TemplateException)
        {
            throw;
        }
        catch (Exception error) when (error is OpenXmlPackageException or XmlException or InvalidDataException
                                          or KeyNotFoundException or ArgumentException or FileFormatException)
        {
            throw Failure(DiagnosticCodes.XlsxInvalidPackage,
                "The XLSX package is inconsistent. Repair or re-save the workbook with valid parts and relationships.");
        }
    }

    private static TemplateException Failure(string code, string message, string? location = null)
    {
        var error = new TemplateException(message, location: location, code: code);
        return new TemplateException(error,
            [new GenerationDiagnostic(code, message, Part: location?.Split(',', 2)[0], Location: location)], false);
    }

    private sealed class Builder(XlsxTemplatePlan plan, XlsxTemplatePlanner planner, SampleJsonOptions options)
    {
        private static readonly BigInteger PreflightLimit = SampleJsonLimits.PreflightSentinel;
        private readonly JsonObject _root = new();
        private readonly Dictionary<string, HeaderDemand> _headers = new(StringComparer.Ordinal);
        private readonly Dictionary<string, HierarchyDemand> _hierarchies = new(StringComparer.Ordinal);
        private readonly List<GenerationDiagnostic> _warnings = [];
        private readonly HashSet<(string Code, string? Location, string? DataPath, string? Field)> _warningKeys = [];
        private readonly Dictionary<string, PropertySlot> _rootSlots = new(StringComparer.Ordinal);

        internal SampleJsonResult Build()
        {
            CollectDemands();
            Preflight();
            MaterializeCollections();
            CollectFieldValues();
            WriteValues();
            using var json = JsonDocument.Parse(_root.ToJsonString());
            return new SampleJsonResult(json.RootElement, _warnings);
        }

        private void CollectDemands()
        {
            foreach (var sheet in plan.Sheets)
            foreach (var group in sheet.Controls.Where(marker => marker.Kind == XlsxControlKind.Header)
                         .GroupBy(marker => marker.Level).OrderBy(group => group.Key))
            {
                var marker = group.First();
                GetHeader(marker.CollectionPath!, marker.Location);
            }

            foreach (var header in _headers.Values)
                header.KnownShape = plan.Dependencies.Any(dependency =>
                    dependency.Kind == PlannedDependency.FieldKind &&
                    dependency.CanonicalPath.StartsWith(header.Path + "[]", StringComparison.Ordinal));

            foreach (var sheet in plan.Sheets)
            {
                foreach (var region in sheet.Regions)
                {
                    var intersects = sheet.Horizontal is { } horizontal && sheet.Cells.Any(cell =>
                        cell.Source.Id.Address.Row >= region.FirstRow && cell.Source.Id.Address.Row <= region.LastRow &&
                        cell.Source.Id.Address.Column >= horizontal.FirstColumn &&
                        cell.Source.Id.Address.Column <= horizontal.LastColumn);
                    var headerPaths = sheet.Controls.Where(marker => marker.Kind == XlsxControlKind.Header)
                        .GroupBy(marker => marker.Level).OrderBy(group => group.Key)
                        .Select(group => group.First().CollectionPath!).ToArray();
                    var headerCounts = headerPaths.Select(path => _headers[path].KnownShape
                        ? options.ItemsPerCollection
                        : BigInteger.Zero).ToArray();
                    var rawArity = intersects && !headerCounts.Contains(BigInteger.Zero)
                        ? headerCounts.Aggregate(BigInteger.One, SaturatingMultiply)
                        : BigInteger.One;
                    if (intersects && options.ItemsPerCollection <= ElementLimit && rawArity <= ElementLimit)
                        rawArity = planner.RecordArity(headerCounts.Select(count => (int) count));
                    var location = sheet.Controls.First(marker => marker.Kind == XlsxControlKind.Vertical &&
                        marker.Cell.Id.Address.Row == region.FirstRow).Location;
                    var demand = GetHierarchy(region.Collection, location, Levels(region.Root), rawArity, intersects);
                    demand.Regions.Add(new RegionUse(sheet, region));
                }
            }
        }

        private HeaderDemand GetHeader(string path, string location)
        {
            if (_hierarchies.ContainsKey(path))
                throw Failure(DiagnosticCodes.SampleIncompatibleDemand,
                    $"Collection '{path}' is used as both an XLSX header array and a vertical hierarchy.", location);
            if (_headers.TryGetValue(path, out var existing)) return existing;
            var demand = new HeaderDemand(path, location);
            _headers.Add(path, demand);
            return demand;
        }

        private HierarchyDemand GetHierarchy(string path, string location, IReadOnlyList<int> levels,
            BigInteger arity, bool requiresPath)
        {
            if (_headers.ContainsKey(path))
                throw Failure(DiagnosticCodes.SampleIncompatibleDemand,
                    $"Collection '{path}' is used as both an XLSX header array and a vertical hierarchy.", location);
            if (!_hierarchies.TryGetValue(path, out var demand))
            {
                demand = new HierarchyDemand(path, levels, arity, requiresPath, location);
                _hierarchies.Add(path, demand);
                return demand;
            }

            if ((!IsLevelPrefix(demand.Levels, levels) && !IsLevelPrefix(levels, demand.Levels)) || demand.Arity != arity)
                throw Failure(DiagnosticCodes.SampleIncompatibleDemand,
                    $"Shared collection '{path}' has incompatible hierarchy levels or horizontal record multiplicity.",
                    location);
            if (levels.Count > demand.Levels.Count) demand.Levels = levels.ToArray();
            demand.RequiresPath |= requiresPath;
            return demand;
        }

        private static bool IsLevelPrefix(IReadOnlyList<int> prefix, IReadOnlyList<int> values)
        {
            return prefix.Count <= values.Count && prefix.SequenceEqual(values.Take(prefix.Count));
        }

        private void Preflight()
        {
            var total = BigInteger.Zero;
            foreach (var header in _headers.Values.Where(header => header.KnownShape))
                total = SaturatingAdd(total, options.ItemsPerCollection);
            foreach (var demand in _hierarchies.Values)
            {
                var records = demand.LogicalRecordCount(options.ItemsPerCollection, PreflightLimit);
                total = SaturatingAdd(total, SaturatingMultiply(records, demand.Arity));
            }
            if (total > ElementLimit)
                throw Failure(DiagnosticCodes.SampleElementLimit,
                    $"The sample would contain more than {ElementLimit:N0} array elements. Reduce ItemsPerCollection.");
        }

        private static BigInteger SaturatingAdd(BigInteger left, BigInteger right)
        {
            if (left >= PreflightLimit || right >= PreflightLimit || left + right >= PreflightLimit)
                return PreflightLimit;
            return left + right;
        }

        private static BigInteger SaturatingMultiply(BigInteger left, BigInteger right)
        {
            if (left.IsZero || right.IsZero) return BigInteger.Zero;
            if (left >= PreflightLimit || right >= PreflightLimit || right > (PreflightLimit - 1) / left)
                return PreflightLimit;
            return left * right;
        }

        private void MaterializeCollections()
        {
            foreach (var header in _headers.Values)
            {
                var items = header.Materialize(header.KnownShape ? (int) options.ItemsPerCollection : 0);
                SetRootArray(header.Path, items);
                if (!header.KnownShape)
                    AddWarning(DiagnosticCodes.SampleUnresolvedShape,
                        $"The item shape for header collection '{header.Path}' cannot be inferred, so the sample collection is empty.",
                        header.Location, header.Path);
            }

            foreach (var hierarchy in _hierarchies.Values)
            {
                hierarchy.Materialize((int) options.ItemsPerCollection, (int) hierarchy.Arity);
                SetRootArray(hierarchy.Path, hierarchy.Data);
            }
        }

        private void CollectFieldValues()
        {
            foreach (var sheet in plan.Sheets)
            {
                var headerLevels = sheet.Controls.Where(marker => marker.Kind == XlsxControlKind.Header)
                    .GroupBy(marker => marker.Level).OrderBy(group => group.Key)
                    .Select(group => (group.Key, group.First().CollectionPath!)).ToArray();
                foreach (var cell in sheet.Cells.OrderBy(item => item.Source.Id.Address.Row)
                             .ThenBy(item => item.Source.Id.Address.Column))
                foreach (var token in cell.Tokens.Where(item => item.Expression?.Kind == TemplateExpressionKind.Field))
                {
                    var expression = token.Expression!;
                    var dependency = plan.Dependencies.FirstOrDefault(item =>
                        item.Kind == PlannedDependency.FieldKind &&
                        item.Location == cell.Location && item.Expression.Source == expression.Source &&
                        item.SourceOrder == SourceOrder(cell, token.Start));
                    if (dependency is null) continue;
                    var use = new SemanticUse(expression.Formatter?.SemanticType,
                        dependency.Location, expression.Source, dependency.CanonicalPath);
                    if (expression.Path!.Segments.Count == 0)
                    {
                        AddWarning(DiagnosticCodes.SampleUnresolvedShape,
                            "The scalar cell refers to a whole XLSX record; the sample leaves that value unresolved.",
                            dependency.Location, dependency.CanonicalPath);
                        continue;
                    }

                    if (expression.Path.Origin == TemplatePathOrigin.Root)
                    {
                        AddValue(_rootSlots, expression.Path.Segments, use);
                        continue;
                    }

                    var headerMarker = sheet.Controls.FirstOrDefault(marker => marker.Kind == XlsxControlKind.Header &&
                        marker.Cell.Id.Address.Row == cell.Source.Id.Address.Row);
                    if (headerMarker is not null)
                    {
                        AddHeaderValue(sheet, headerMarker, headerLevels, expression, use);
                        continue;
                    }

                    var regionUse = sheet.Regions.FirstOrDefault(region =>
                        cell.Source.Id.Address.Row >= region.FirstRow && cell.Source.Id.Address.Row <= region.LastRow);
                    if (regionUse is null)
                    {
                        AddValue(_rootSlots, expression.Path.Segments, use);
                        continue;
                    }

                    var hierarchy = _hierarchies[regionUse.Collection];
                    var depth = ScopeDepth(sheet, regionUse, hierarchy, cell.Source.Id.Address.Row);
                    var targetDepth = depth - expression.Path.ParentDepth;
                    if (targetDepth < 0)
                    {
                        AddWarning(DiagnosticCodes.SampleUnavailableAncestor,
                            $"No represented XLSX ancestor exists for '{dependency.CanonicalPath}'. The sample leaves the field unresolved.",
                            dependency.Location, dependency.CanonicalPath);
                        continue;
                    }

                    IEnumerable<int> tupleIndices = expression.Path.ParentDepth == 0 &&
                        IsLeafCell(sheet, cell, headerLevels.Length)
                        ? Enumerable.Range(0, (int) hierarchy.Arity)
                        : new[] {0};
                    foreach (var current in hierarchy.AtDepth[depth])
                    {
                        var target = current;
                        for (var parent = 0; parent < expression.Path.ParentDepth; parent++) target = target.Parent!;
                        foreach (var tuple in tupleIndices)
                            AddValue(target.Slots[tuple], expression.Path.Segments, use);
                    }
                }
            }
        }

        private void AddHeaderValue(XlsxSheetTemplate sheet, XlsxControlMarker marker,
            (int Level, string Path)[] levels, TemplateExpression expression, SemanticUse use)
        {
            var levelIndex = Array.FindIndex(levels, item => item.Level == marker.Level);
            var targetIndex = levelIndex - expression.Path!.ParentDepth;
            if (targetIndex == -1)
            {
                AddValue(_rootSlots, expression.Path.Segments, use);
                return;
            }
            if (targetIndex < -1)
            {
                AddWarning(DiagnosticCodes.SampleUnavailableAncestor,
                    $"No represented horizontal header ancestor exists for '{use.CanonicalPath}'. The sample leaves the field unresolved.",
                    use.Location, use.CanonicalPath);
                return;
            }

            var target = _headers[levels[targetIndex].Path];
            foreach (var item in target.Items) AddValue(item.Slots, expression.Path.Segments, use);
        }

        private void WriteValues()
        {
            WriteSlots(_root, _rootSlots, 1);
            foreach (var header in _headers.Values)
                for (var index = 0; index < header.Items.Count; index++)
                    WriteSlots(header.Items[index].Data, header.Items[index].Slots, index + 1);
            foreach (var hierarchy in _hierarchies.Values)
                foreach (var record in hierarchy.Records)
                for (var tuple = 0; tuple < record.Data.Count; tuple++)
                    WriteSlots(record.Data[tuple], record.Slots[tuple], record.Ordinal * record.Data.Count + tuple + 1);
        }

        private void WriteSlots(JsonObject target, Dictionary<string, PropertySlot> slots, int ordinal)
        {
            foreach (var (key, slot) in slots)
            {
                var values = slot.Uses.ToArray();
                var resolution = SampleJsonScalarValue.Resolve(values.Select(item => item.SemanticType));
                if (resolution.HasConflict)
                {
                    var firstKnown = values.First(item => item.SemanticType is not null);
                    AddWarning(DiagnosticCodes.SampleSemanticConflict,
                        $"Conflicting semantic hints for '{key}' use the first source hint.",
                        firstKnown.Location, key, firstKnown.Expression);
                }
                JsonNode? value = SampleJsonScalarValue.CreateValue(options, resolution.SemanticType,
                    key.Replace('\u001f', '.'), ordinal);
                SetPath(target, slot.Segments, value);
            }
        }

        private void AddWarning(string code, string message, string? location, string? dataPath,
            string? field = null)
        {
            var key = (code, location, dataPath, field);
            if (!_warningKeys.Add(key)) return;
            _warnings.Add(new GenerationDiagnostic(code, message, field, location?.Split(',', 2)[0], location,
                DataPath: dataPath));
        }

        private static void AddValue(Dictionary<string, PropertySlot> slots, IReadOnlyList<string> path,
            SemanticUse use)
        {
            var key = string.Join('\u001f', path);
            if (!slots.TryGetValue(key, out var slot))
            {
                slot = new PropertySlot(path.ToArray());
                slots.Add(key, slot);
            }
            slot.Uses.Add(use);
        }

        private void SetRootArray(string path, JsonArray value)
        {
            var segments = PathSegments(path);
            if (segments.Length == 0) return;
            var target = _root;
            for (var index = 0; index < segments.Length - 1; index++)
            {
                if (target[segments[index]] is not JsonObject child)
                {
                    child = new JsonObject();
                    target[segments[index]] = child;
                }
                target = child;
            }
            target[segments[^1]] = value;
        }

        private static void SetPath(JsonObject root, IReadOnlyList<string> segments, JsonNode? value)
        {
            var target = root;
            for (var index = 0; index < segments.Count - 1; index++)
            {
                if (target[segments[index]] is not JsonObject child)
                {
                    child = new JsonObject();
                    target[segments[index]] = child;
                }
                target = child;
            }
            target[segments[^1]] = value;
        }

        private static string[] PathSegments(string path) => path.StartsWith("$.", StringComparison.Ordinal)
            ? path[2..].Split('.')
            : [];

        private static int ScopeDepth(XlsxSheetTemplate sheet, XlsxVerticalRegion region,
            HierarchyDemand hierarchy, int row)
        {
            var marker = sheet.Controls.FirstOrDefault(item => item.Kind == XlsxControlKind.Vertical &&
                item.Cell.Id.Address.Row == row && item.CollectionPath == region.Collection);
            if (marker is not null)
            {
                for (var index = 0; index < hierarchy.Levels.Count; index++)
                    if (hierarchy.Levels[index] == marker.Level) return index;
            }
            return ScopeDepth(region.Root, row, 0);
        }

        private static int ScopeDepth(XlsxVerticalNode node, int row, int depth)
        {
            foreach (var child in node.Children)
            {
                var first = int.MaxValue;
                var last = int.MinValue;
                foreach (var block in new[] {child.Header, child.Body, child.Footer})
                    if (block is not null)
                    {
                        first = Math.Min(first, block.FirstRow);
                        last = Math.Max(last, block.LastRow);
                    }
                if (row >= first && row <= last) return ScopeDepth(child, row, depth + 1);
            }
            return depth;
        }

        private static IReadOnlyList<int> Levels(XlsxVerticalNode root)
        {
            var levels = new List<int>();
            for (var current = root; ; current = current.Children[0])
            {
                levels.Add(current.Level);
                if (current.Children.IsEmpty) return levels;
            }
        }

        private static long SourceOrder(XlsxCellPlan cell, int offset)
        {
            var address = cell.Source.Id.Address;
            return ((long) address.Row << 40) + ((long) address.Column << 24) + (uint) offset;
        }

        private static bool IsLeafCell(XlsxSheetTemplate sheet, XlsxCellPlan cell, int headerDepth)
        {
            if (sheet.Horizontal is null || headerDepth == 0) return false;
            var path = HorizontalPath(sheet.Horizontal, cell.Source.Id.Address.Column);
            return path.Count == headerDepth && path[^1].Children.IsEmpty;
        }

        private static IReadOnlyList<XlsxHorizontalNode> HorizontalPath(XlsxHorizontalNode root, int column)
        {
            if (column < root.FirstColumn || column > root.LastColumn) return [];
            foreach (var child in root.Children)
            {
                var path = HorizontalPath(child, column);
                if (path.Count != 0) return new[] {root}.Concat(path).ToArray();
            }
            return [root];
        }

    }

    private sealed class HeaderDemand(string path, string location)
    {
        internal string Path { get; } = path;
        internal string Location { get; } = location;
        internal bool KnownShape { get; set; }
        internal List<HeaderItem> Items { get; } = [];

        internal JsonArray Materialize(int count)
        {
            var result = new JsonArray();
            if (!KnownShape) return result;
            for (var index = 0; index < count; index++)
            {
                var item = new HeaderItem();
                Items.Add(item);
                result.Add(item.Data);
            }
            return result;
        }
    }

    private sealed class HeaderItem
    {
        internal JsonObject Data { get; } = new();
        internal Dictionary<string, PropertySlot> Slots { get; } = new(StringComparer.Ordinal);
    }

    private sealed class HierarchyDemand(string path, IReadOnlyList<int> levels, BigInteger arity, bool requiresPath,
        string location)
    {
        internal string Path { get; } = path;
        internal IReadOnlyList<int> Levels { get; set; } = levels;
        internal BigInteger Arity { get; } = arity;
        internal bool RequiresPath { get; set; } = requiresPath;
        internal string Location { get; } = location;
        internal List<RegionUse> Regions { get; } = [];
        internal List<RecordNode> Records { get; } = [];
        internal List<RecordNode>[] AtDepth { get; private set; } = [];
        internal JsonArray Data { get; private set; } = [];

        internal BigInteger LogicalRecordCount(BigInteger perCollection, BigInteger limit)
        {
            var total = BigInteger.Zero;
            var count = BigInteger.Min(perCollection, limit);
            foreach (var _ in Levels)
            {
                total = BigInteger.Min(limit, total + count);
                if (total == limit) return total;
                count = BigInteger.Min(limit, count * perCollection);
            }
            return total;
        }

        internal void Materialize(int count, int arity)
        {
            AtDepth = Levels.Select(_ => new List<RecordNode>()).ToArray();
            Data = [];
            var ordinal = 0;
            AddDepth(0, null, count, arity, RequiresPath, ref ordinal);
            foreach (var record in Records)
                foreach (var item in record.Data) Data.Add(item);
        }

        private void AddDepth(int depth, RecordNode? parent, int count, int arity, bool requiresPath,
            ref int ordinal)
        {
            for (var index = 0; index < count; index++)
            {
                var record = new RecordNode(Levels[depth], ordinal++, parent, arity, requiresPath);
                Records.Add(record);
                AtDepth[depth].Add(record);
                if (depth + 1 < Levels.Count)
                    AddDepth(depth + 1, record, count, arity, requiresPath, ref ordinal);
            }
        }
    }

    private sealed class RecordNode(int level, int ordinal, RecordNode? parent, int arity, bool requiresPath)
    {
        internal int Ordinal { get; } = ordinal;
        internal RecordNode? Parent { get; } = parent;
        internal List<JsonObject> Data { get; } = CreateData(level, ordinal, arity, requiresPath);
        internal Dictionary<string, PropertySlot>[] Slots { get; } =
            Enumerable.Range(0, arity).Select(_ => new Dictionary<string, PropertySlot>(StringComparer.Ordinal)).ToArray();

        private static List<JsonObject> CreateData(int level, int ordinal, int arity, bool requiresPath)
        {
            var path = "sample-" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return Enumerable.Range(0, arity).Select(_ =>
            {
                var item = new JsonObject {{LevelProperty, level}};
                if (requiresPath) item[PathProperty] = path;
                return item;
            }).ToList();
        }
    }

    private sealed class PropertySlot(string[] segments)
    {
        internal string[] Segments { get; } = segments;
        internal List<SemanticUse> Uses { get; } = [];
    }

    private sealed record SemanticUse(SemanticType? SemanticType, string Location,
        string Expression, string CanonicalPath);
    private sealed record RegionUse(XlsxSheetTemplate Sheet, XlsxVerticalRegion Region);
}
