using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace MonoReport.Models;

internal sealed record TextNodeSpan(XElement Node, int Start, int Length, int SourceOrder);

internal sealed record ParagraphPlan(
    XElement Paragraph,
    string Text,
    IReadOnlyList<TextNodeSpan> TextNodes,
    IReadOnlyList<TemplateToken> Tokens,
    string Location,
    long SourceOrder)
{
    internal long PositionAt(int offset)
    {
        var span = TextNodes.FirstOrDefault(item =>
            offset >= item.Start && offset < item.Start + item.Length) ?? TextNodes[^1];
        return ((long)span.SourceOrder << 32) + Math.Max(0, offset - span.Start);
    }
}

internal sealed class BlockPlan(
    TemplateExpressionKind kind,
    TemplateExpression startExpression,
    XElement startContainer,
    XElement endContainer,
    XElement? elseContainer,
    IReadOnlyList<XElement> trueBody,
    IReadOnlyList<XElement> falseBody,
    string location,
    IReadOnlyList<BlockPlan> children)
{
    internal TemplateExpressionKind Kind { get; } = kind;
    internal TemplateExpression StartExpression { get; } = startExpression;
    internal XElement StartContainer { get; } = startContainer;
    internal XElement EndContainer { get; } = endContainer;
    internal XElement? ElseContainer { get; } = elseContainer;
    internal IReadOnlyList<XElement> TrueBody { get; } = trueBody;
    internal IReadOnlyList<XElement> FalseBody { get; } = falseBody;
    internal string Location { get; } = location;
    internal IReadOnlyList<BlockPlan> Children { get; } = children;
}

internal sealed record PlannedDependency(
    string Kind,
    TemplateExpression Expression,
    string CanonicalPath,
    string RequirementType,
    string Part,
    string Location,
    long SourceOrder);

internal sealed record StoryPlan(
    OpenXmlPart Part,
    XDocument Xml,
    IReadOnlyList<ParagraphPlan> Paragraphs,
    IReadOnlyList<BlockPlan> Blocks,
    IReadOnlyList<PlannedDependency> Dependencies,
    IReadOnlyList<StructuralDiagnostic> Diagnostics);

// Source XML identities remain stable when intermediate structural copies are discarded.
internal sealed record StructuralOwner(BlockPlan Block, string Branch);

internal sealed record IdentityLocation(XElement Element, string Part, string Location, long SourceOrder,
    XElement Root, IReadOnlyList<StructuralOwner> Owners);

internal sealed record SourceRange(string Kind, string Id, IReadOnlyList<IdentityLocation> Members,
    XElement? CommentDefinition = null)
{
    internal IdentityLocation First => Members[0];
    internal string? Name => (string?)First.Element.Attribute(
        XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main") + "name");
}

internal sealed record DocxIdentityPlan(
    IReadOnlyDictionary<XElement, SourceRange> Ranges,
    IReadOnlyList<StructuralDiagnostic> Diagnostics,
    IReadOnlyList<XElement>? DeletedCommentDefinitions = null);
