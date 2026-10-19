using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;

namespace MonoReport.Services;

public sealed partial class DocxTemplatePlanner
{
    private static DocxIdentityPlan ValidateIdentities(StoryInspection package, IReadOnlyList<StoryPlan> stories,
        IReadOnlyDictionary<XElement, IdentityLocation> locations)
    {
        var w = WordXml.WordprocessingNamespace;
        var ranges = new Dictionary<XElement, SourceRange>();
        var diagnostics = new List<StructuralDiagnostic>();
        var commentXml = package.PackageXml.FirstOrDefault(item => item.Part is WordprocessingCommentsPart).Xml;
        var definitions = (commentXml?.Root?.Elements(w + "comment") ?? []).ToLookup(CommentIdentity);
        var shapes = new Dictionary<SourceRange, bool>();
        var blocks = stories.SelectMany(story => AllBlocks(story.Blocks)).ToArray();
        foreach (var family in new[]
        {
            (Kind: "bookmark", Names: new[] { "bookmarkStart", "bookmarkEnd" }),
            (Kind: "permission", Names: new[] { "permStart", "permEnd" }),
            (Kind: "comment", Names: new[] { "commentRangeStart", "commentRangeEnd", "commentReference" })
        })
        {
            var names = family.Names.Select(name => w + name).ToHashSet();
            foreach (var group in locations.Values.Where(item => names.Contains(item.Element.Name))
                         .GroupBy(item => (item.Root, Id: family.Kind == "comment" ? CommentIdentity(item.Element) :
                             (string?)item.Element.Attribute(w + "id") ?? "")))
            {
                var members = group.OrderBy(item => item.SourceOrder).ToArray();
                var matched = family.Kind == "comment" ? definitions[group.Key.Id].ToArray() : [];
                var range = new SourceRange(family.Kind, group.Key.Id, members, matched.Length == 1 ? matched[0] : null);
                foreach (var member in members) ranges.Add(member.Element, range);
                var point = family.Kind == "comment" && members.Length == 1 && members[0].Element.Name == w + "commentReference";
                var shape = point || (members.Length == family.Names.Length && family.Names.All(name =>
                    members.Count(member => member.Element.Name == w + name) == 1) &&
                    Array.FindIndex(members, member => member.Element.Name == w + family.Names[0]) <
                    Array.FindIndex(members, member => member.Element.Name == w + family.Names[1]));
                shapes.Add(range, shape);
                var affected = members.Any(member => member.Owners.Count != 0) || blocks.Any(block =>
                    IndependentRoot(block.StartContainer) == group.Key.Root &&
                    members[0].SourceOrder < locations[block.StartContainer].SourceOrder &&
                    members[^1].SourceOrder > locations[block.EndContainer].SourceOrder);
                if (!affected) continue;
                var local = members.Any(member => member.Owners.Any(owner => owner.Block.Kind != TemplateExpressionKind.Raw));
                var removed = members.Select(RawDelimiterOwner).ToArray();
                var splitDeletion = removed.Any(item => item is not null) &&
                    (removed.Any(item => item is null) || removed.Any(item => item != removed[0]));
                var samePermissionScope = family.Kind != "permission" || members.All(member =>
                    member.Owners.Where(owner => owner.Block.Kind != TemplateExpressionKind.Raw).SequenceEqual(
                        members[0].Owners.Where(owner => owner.Block.Kind != TemplateExpressionKind.Raw)));
                if (shape && !splitDeletion && samePermissionScope && !(local && family.Kind != "permission") &&
                    (family.Kind != "comment" || matched.Length == 1)) continue;
                var coordinates = string.Join("; ", members.Take(4).Select(member =>
                    member.Element.Name.LocalName + " at " + member.Location));
                var remedy = local && family.Kind != "permission"
                    ? "Move the complete annotation outside the For/If operation, or remove it in Word."
                    : "Keep complete, ordered endpoints and attachments together; move them off deleted delimiters in Word.";
                var detail = $"Unsupported {family.Kind} '{Bound(range.Id)}' '{Bound(range.Name)}': invalid range/attachment or structural ownership. {coordinates}. {remedy}";
                if (family.Kind == "comment" && matched.Length != 1)
                    detail += " Each affected comment must resolve to exactly one source definition.";
                diagnostics.Add(new(DiagnosticCodes.TemplateInvalidStructure, detail, Location: members[0].Location,
                    SourceOrder: members[0].SourceOrder, Explanation: detail));
            }
        }

        var uniqueRanges = ranges.Values.Distinct().ToArray();
        var deletedDefinitions = uniqueRanges.Where(range => range.Kind == "comment" && range.CommentDefinition is not null)
            .GroupBy(range => range.CommentDefinition!).Where(group => group.All(range =>
                shapes[range] && range.Members.All(member => RawDelimiterOwner(member) is not null) &&
                range.Members.All(member => RawDelimiterOwner(member) == RawDelimiterOwner(range.First))))
            .Select(group => group.Key).ToHashSet();
        var bookmarks = uniqueRanges.Where(range => range.Kind == "bookmark" && range.Name is not null)
            .GroupBy(range => range.Name!, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        var links = locations.Values.Where(item => item.Element.Name == WordXml.HyperlinkElementName &&
            item.Element.Attribute(WordXml.RelationshipsNamespace + "id") is null &&
            item.Element.Attribute(WordXml.AnchorAttributeName) is not null).ToArray();
        // A surviving explicit reference can require a candidate definition to remain. No identities are changed.
        bool retained;
        do
        {
            retained = false;
            foreach (var link in links.Where(link => !Removed(link)))
            {
                if (!bookmarks.TryGetValue((string)link.Element.Attribute(WordXml.AnchorAttributeName)!, out var targets)) continue;
                foreach (var target in targets)
                    foreach (var definition in target.First.Element.Ancestors().Where(deletedDefinitions.Contains).ToArray())
                        retained |= deletedDefinitions.Remove(definition);
            }
        } while (retained);

        foreach (var use in links)
        {
            if (Removed(use)) continue;
            var anchor = (string)use.Element.Attribute(WordXml.AnchorAttributeName)!;
            var local = use.Owners.Any(owner => owner.Block.Kind != TemplateExpressionKind.Raw);
            bookmarks.TryGetValue(anchor, out var targets);
            var affected = targets?.Any(target => target.Members.Any(member => member.Owners.Count != 0 || Removed(member))) == true;
            if (!local && !affected) continue; // Preserve unrelated pre-existing unresolved anchors.
            if (!local && targets is { Length: 1 } && shapes[targets[0]] &&
                targets[0].Members.All(member => !Removed(member) && member.Owners.All(owner => owner.Block.Kind == TemplateExpressionKind.Raw))) continue;
            var detail = $"Internal anchor '{Bound(anchor)}' is inside For/If or has a removed, repeated, ambiguous or invalid target" +
                (targets is { Length: > 0 } ? $" at {targets[0].First.Location}. " : ". ") +
                "Move the internal link and its unchanged target outside For/If and off Raw delimiters, or remove the link in Word.";
            diagnostics.Add(new(DiagnosticCodes.TemplateInvalidStructure, detail, Location: use.Location,
                SourceOrder: use.SourceOrder, Explanation: detail));
        }
        return new(ranges, diagnostics, deletedDefinitions.ToArray());

        bool Removed(IdentityLocation item) => RawDelimiterOwner(item) is not null ||
            item.Element.AncestorsAndSelf().Any(deletedDefinitions.Contains);
    }

    private static XElement? RawDelimiterOwner(IdentityLocation location)
    {
        var owner = location.Owners.FirstOrDefault(owner => owner.Block.Kind == TemplateExpressionKind.Raw &&
            owner.Branch.StartsWith("marker", StringComparison.Ordinal));
        return owner is null ? null : owner.Branch == "marker-start" ? owner.Block.StartContainer : owner.Block.EndContainer;
    }

    private static IEnumerable<BlockPlan> AllBlocks(IReadOnlyList<BlockPlan> blocks)
    {
        foreach (var block in blocks)
        {
            yield return block;
            foreach (var child in AllBlocks(block.Children)) yield return child;
        }
    }
}
