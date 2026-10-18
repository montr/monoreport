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
        private readonly HashSet<int> _bookmarkIds = [];
        private readonly HashSet<string> _bookmarkNames = new(StringComparer.Ordinal);
        private readonly HashSet<int> _permissionIds = [];
        private readonly HashSet<int> _commentIds = [];
        private WordprocessingCommentsPart? _commentsPart;
        private XDocument? _commentsXml;

        private uint _drawingIdCandidate = 1;

        internal DocxIdentityPlan Identities { get; private set; } = new(
            new Dictionary<XElement, SourceRange>(), new Dictionary<XElement, SourceRange>(), []);

        internal IdentityFrame? ActiveIdentities { get; set; }

        internal void InitializePackageIds(WordprocessingDocument document, TemplateInspection inspection)
        {
            Identities = inspection.Identities;
            var elements = inspection.Package.PackageXml.SelectMany(item => item.Xml.Descendants()).ToArray();
            foreach (var element in elements.Where(element => element.Name == DrawingPropertiesName))
                if (uint.TryParse((string?)element.Attribute("id"), out var id) && id != 0) _drawingIds.Add(id);
            foreach (var element in elements.Where(element => element.Name == WordprocessingNamespace + "bookmarkStart"))
            {
                if (int.TryParse((string?)element.Attribute(WordprocessingNamespace + "id"), out var id)) _bookmarkIds.Add(id);
                var name = (string?)element.Attribute(WordprocessingNamespace + "name");
                if (name is not null) _bookmarkNames.Add(name);
            }
            foreach (var element in elements.Where(element => element.Name == WordprocessingNamespace + "permStart"))
                if (int.TryParse((string?)element.Attribute(WordprocessingNamespace + "id"), out var id)) _permissionIds.Add(id);
            _commentsPart = document.MainDocumentPart?.WordprocessingCommentsPart;
            _commentsXml = inspection.Package.PackageXml.FirstOrDefault(item => item.Part == _commentsPart).Xml;
            if (_commentsXml is not null)
                foreach (var comment in _commentsXml.Root!.Elements(WordprocessingNamespace + "comment"))
                    if (int.TryParse((string?)comment.Attribute(WordprocessingNamespace + "id"), out var id)) _commentIds.Add(id);
        }

        internal IdentityFrame CreateIdentityFrame(BlockPlan block, IReadOnlyList<XElement> selected)
        {
            var frame = new IdentityFrame(ActiveIdentities);
            foreach (var range in Identities.Ranges.Values.Distinct().Where(range => range.Kind == "bookmark" &&
                         selected.Any(root => range.First.Element.AncestorsAndSelf().Contains(root))))
            {
                var owner = range.First.Loops.LastOrDefault();
                if ((block.Kind == TemplateExpressionKind.For ? owner != block : owner is not null) ||
                    frame.Find(range) is not null) continue;
                frame.Bookmarks.Add(range, new(NextBookmarkId(), owner is null ? range.Name! : NextBookmarkName(range.Name!)));
            }
            return frame;
        }

        // One physical-copy path serves stories and comment definitions; all reservations are operation-wide.
        internal IReadOnlyList<XElement> Clone(IReadOnlyList<XElement> roots, Dictionary<XElement, XElement> map,
            bool definition = false)
        {
            var clones = roots.Select(root => new XElement(root)).ToArray();
            foreach (var pair in roots.SelectMany(root => root.DescendantsAndSelf())
                         .Zip(clones.SelectMany(root => root.DescendantsAndSelf()))) map.Add(pair.First, pair.Second);
            var bookmarkBindings = new Dictionary<SourceRange, BookmarkIdentity>();
            var permissionIds = new Dictionary<SourceRange, int>();
            var commentIds = new Dictionary<SourceRange, int>();
            foreach (var source in map.Keys)
            {
                if (!Identities.Ranges.TryGetValue(source, out var range)) continue;
                if (range.Kind == "bookmark" && !bookmarkBindings.ContainsKey(range))
                    bookmarkBindings.Add(range, definition ? new(NextBookmarkId(), NextBookmarkName(range.Name!)) :
                        ActiveIdentities?.Find(range) ?? new(NextBookmarkId(), range.First.Loops.Count == 0 ? range.Name! : NextBookmarkName(range.Name!)));
                else if (range.Kind == "permission" && !permissionIds.ContainsKey(range)) permissionIds.Add(range, NextPermissionId());
                else if (range.Kind == "comment" && !definition && !commentIds.ContainsKey(range))
                    commentIds.Add(range, CloneComment(range));
            }
            foreach (var (source, clone) in map)
            {
                if (clone.Name == DrawingPropertiesName) clone.SetAttributeValue("id", NextDrawingId());
                if (Identities.Ranges.TryGetValue(source, out var range))
                {
                    if (bookmarkBindings.TryGetValue(range, out var bookmark))
                    {
                        clone.SetAttributeValue(WordprocessingNamespace + "id", bookmark.Id);
                        if (source.Name == WordprocessingNamespace + "bookmarkStart")
                            clone.SetAttributeValue(WordprocessingNamespace + "name", bookmark.Name);
                    }
                    else if (permissionIds.TryGetValue(range, out var permissionId)) clone.SetAttributeValue(WordprocessingNamespace + "id", permissionId);
                    else if (commentIds.TryGetValue(range, out var commentId)) clone.SetAttributeValue(WordprocessingNamespace + "id", commentId);
                }
                if (source.Name != WordprocessingNamespace + "hyperlink" ||
                    source.Attribute(XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships") + "id") is not null) continue;
                SourceRange? target = null;
                if (definition)
                {
                    var anchor = (string?)source.Attribute(WordprocessingNamespace + "anchor");
                    target = bookmarkBindings.Keys.SingleOrDefault(range => range.Name == anchor);
                }
                if (target is null) Identities.Anchors.TryGetValue(source, out target);
                if (target is null) continue;
                var binding = bookmarkBindings.GetValueOrDefault(target) ?? ActiveIdentities?.Find(target);
                clone.SetAttributeValue(WordprocessingNamespace + "anchor", binding?.Name ?? target.Name);
            }
            return clones;
        }

        internal sealed record BookmarkIdentity(int Id, string Name);

        internal sealed class IdentityFrame(IdentityFrame? parent)
        {
            internal Dictionary<SourceRange, BookmarkIdentity> Bookmarks { get; } = [];
            internal BookmarkIdentity? Find(SourceRange source) => Bookmarks.GetValueOrDefault(source) ?? parent?.Find(source);
        }

        internal uint NextDrawingId()
        {
            while (_drawingIds.Contains(_drawingIdCandidate)) _drawingIdCandidate++;
            var result = _drawingIdCandidate++;
            _drawingIds.Add(result);
            return result;
        }

        internal int NextBookmarkId() => NextId(_bookmarkIds);

        internal int NextPermissionId() => NextId(_permissionIds);

        private int CloneComment(SourceRange range)
        {
            var source = range.CommentDefinition ??
                throw new InvalidOperationException("Comment cloning requires a preflight-bound source definition.");
            var id = NextId(_commentIds);
            var clone = Clone([source], new Dictionary<XElement, XElement>(), definition: true)[0];
            clone.SetAttributeValue(WordprocessingNamespace + "id", id);
            _commentsXml!.Root!.Add(clone);
            return id;
        }

        internal void SavePackageIdentities()
        {
            if (_commentsPart is null || _commentsXml is null) return;
            using var stream = _commentsPart.GetStream(FileMode.Create, FileAccess.Write);
            _commentsXml.Save(stream, SaveOptions.DisableFormatting);
        }

        internal string NextBookmarkName(string basis)
        {
            for (var suffix = 1;; suffix++)
            {
                var ending = "_" + suffix;
                var prefix = basis[..Math.Min(basis.Length, 40 - ending.Length)];
                var candidate = prefix + ending;
                if (_bookmarkNames.Add(candidate)) return candidate;
            }
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
