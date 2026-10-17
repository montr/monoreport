using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace MonoReport.Models;

internal sealed record TextNodeSpan(XElement Node, int Start, int Length);

internal sealed record ParagraphPlan(
    XElement Paragraph,
    string Text,
    IReadOnlyList<TextNodeSpan> TextNodes,
    IReadOnlyList<TemplateToken> Tokens,
    string Location);

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
    string Location);

internal sealed record StoryPlan(
    OpenXmlPart Part,
    XDocument Xml,
    IReadOnlyList<ParagraphPlan> Paragraphs,
    IReadOnlyList<BlockPlan> Blocks,
    IReadOnlyList<PlannedDependency> Dependencies,
    IReadOnlyList<StructuralDiagnostic> Diagnostics);
