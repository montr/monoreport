using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace MonoReport.Models
{
    public sealed class MergeContext(JsonElement root, MergeOptions options)
    {
        internal const int ErrorLimit = 100;

        public Stack<JsonElement> Values { get; } = new([root]);

        internal JsonElement Root { get; } = root;

        internal Stack<string> DataPaths { get; } = new(["$"]);

        public MergeOptions Options { get; } = options;

        public List<MergeDiagnostic> Warnings { get; } = new();

        internal List<GenerationDiagnostic> Errors { get; } = new();

        internal TemplateException? FirstError { get; private set; }

        internal bool LimitReached { get; private set; }

        internal HashSet<string> WarningKeys { get; } = new(StringComparer.Ordinal);

        private static readonly XName DrawingPropertiesName =
            XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing") + "docPr";

        private readonly HashSet<uint> _drawingIds = [];
        private readonly HashSet<int> _permissionIds = [];
        private WordprocessingCommentsPart? _commentsPart;
        private XDocument? _commentsXml;

        private uint _drawingIdCandidate = 1;

        internal DocxIdentityPlan Identities { get; private set; } = new(
            new Dictionary<XElement, SourceRange>(), []);

        internal void InitializePackageIds(WordprocessingDocument document, TemplateInspection inspection)
        {
            Identities = inspection.Identities;
            var elements = inspection.Package.PackageXml.SelectMany(item => item.Xml.Descendants()).ToArray();
            foreach (var element in elements.Where(element => element.Name == DrawingPropertiesName))
                if (uint.TryParse((string?)element.Attribute("id"), out var id) && id != 0) _drawingIds.Add(id);
            foreach (var element in elements.Where(element => element.Name == WordprocessingNamespace + "permStart"))
                if (int.TryParse((string?)element.Attribute(WordprocessingNamespace + "id"), out var id)) _permissionIds.Add(id);
            _commentsPart = document.MainDocumentPart?.WordprocessingCommentsPart;
            _commentsXml = inspection.Package.PackageXml.FirstOrDefault(item => item.Part == _commentsPart).Xml;
        }

        // One physical-copy path serves every story; drawing and permission reservations are operation-wide.
        internal IReadOnlyList<XElement> Clone(IReadOnlyList<XElement> roots, Dictionary<XElement, XElement> map)
        {
            var clones = roots.Select(root => new XElement(root)).ToArray();
            foreach (var pair in roots.SelectMany(root => root.DescendantsAndSelf())
                         .Zip(clones.SelectMany(root => root.DescendantsAndSelf()))) map.Add(pair.First, pair.Second);
            var permissionIds = new Dictionary<SourceRange, int>();
            foreach (var source in map.Keys)
            {
                if (!Identities.Ranges.TryGetValue(source, out var range)) continue;
                if (range.Kind == "permission" && !permissionIds.ContainsKey(range))
                    permissionIds.Add(range, NextPermissionId());
            }
            foreach (var (source, clone) in map)
            {
                if (clone.Name == DrawingPropertiesName) clone.SetAttributeValue("id", NextDrawingId());
                if (Identities.Ranges.TryGetValue(source, out var range) &&
                    permissionIds.TryGetValue(range, out var permissionId))
                    clone.SetAttributeValue(WordprocessingNamespace + "id", permissionId);
            }
            return clones;
        }

        internal uint NextDrawingId()
        {
            while (_drawingIds.Contains(_drawingIdCandidate)) _drawingIdCandidate++;
            var result = _drawingIdCandidate++;
            _drawingIds.Add(result);
            return result;
        }

        internal int NextPermissionId() => NextId(_permissionIds);

        internal void SaveDeletedCommentDefinitions()
        {
            if (_commentsPart is null || _commentsXml is null) return;
            var deleted = Identities.DeletedCommentDefinitions ?? [];
            if (deleted.Count == 0) return;
            foreach (var definition in deleted) definition.Remove();
            using var stream = _commentsPart.GetStream(FileMode.Create, FileAccess.Write);
            _commentsXml.Save(stream, SaveOptions.DisableFormatting);
        }

        private static int NextId(HashSet<int> ids)
        {
            var candidate = 0;
            while (ids.Contains(candidate)) candidate++;
            ids.Add(candidate);
            return candidate;
        }

        private static readonly XNamespace WordprocessingNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        internal bool Record(TemplateException error, string dataPath, string? location = null)
        {
            if (LimitReached) return false;
            FirstError ??= error;
            Errors.Add(error.ToDiagnostic(dataPath, location));
            if (Errors.Count == ErrorLimit) LimitReached = true;
            return !LimitReached;
        }
    }
}
