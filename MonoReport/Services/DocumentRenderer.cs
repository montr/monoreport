using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using MonoReport.Models;

namespace MonoReport.Services;

public sealed class DocumentRenderer(DataResolver resolver, ValueFormatter formatter)
{
    private readonly DataResolver _resolver = resolver ?? throw new ArgumentNullException(nameof(resolver));
    private readonly ValueFormatter _formatter = formatter ?? throw new ArgumentNullException(nameof(formatter));

    internal bool Render(StoryPlan plan, MergeContext context)
    {
        var map = plan.Xml.Root!.DescendantsAndSelf().ToDictionary(element => element, element => element);
        return RenderScope(plan, plan.Paragraphs, plan.Blocks, map, context);
    }

    private bool RenderScope(StoryPlan plan, IReadOnlyList<ParagraphPlan> paragraphs,
        IReadOnlyList<BlockPlan> blocks, Dictionary<XElement, XElement> map, MergeContext context)
    {
        foreach (var block in blocks)
            if (!RenderBlock(plan, block, map, context)) return false;

        foreach (var paragraph in paragraphs)
        {
            if (blocks.Any(block => Contains(block, paragraph.Paragraph))) continue;
            if (!map.ContainsKey(paragraph.Paragraph)) continue;
            if (!RenderParagraph(paragraph, map, context)) return false;
        }
        return true;
    }

    private bool RenderBlock(StoryPlan plan, BlockPlan block, Dictionary<XElement, XElement> parentMap,
        MergeContext context)
    {
        if (!parentMap.TryGetValue(block.StartContainer, out var targetStart) || targetStart.Parent is null) return true;
        if (block.Kind == TemplateExpressionKind.Raw)
        {
            RemoveAndEnsureCellStructure([targetStart, parentMap[block.EndContainer]]);
            return true;
        }
        var resolution = _resolver.Resolve(block.StartExpression.Path!, context);
        if (block.Kind == TemplateExpressionKind.For)
        {
            if (resolution.State == ResolutionState.Missing)
            {
                WarnMissing(block, context);
            }
            else if (resolution.State == ResolutionState.Present && resolution.Value.ValueKind != JsonValueKind.Array)
            {
                if (!RecordTypeError(block, context, "For requires a JSON array.")) return false;
            }
            else if (resolution.State == ResolutionState.Present)
            {
                var itemIndex = 0;
                foreach (var item in resolution.Value.EnumerateArray())
                {
                    if (context.LimitReached) return false;
                    context.Values.Push(item);
                    context.DataPaths.Push(_resolver.DataPath(block.StartExpression.Path!, context) + $"[{itemIndex}]");
                    try
                    {
                        if (!RenderSelected(plan, block, block.TrueBody, targetStart, parentMap, context)) return false;
                    }
                    finally
                    {
                        context.DataPaths.Pop();
                        context.Values.Pop();
                    }
                    itemIndex++;
                }
            }
        }
        else
        {
            var selected = block.FalseBody;
            if (resolution.State == ResolutionState.Missing) WarnMissing(block, context);
            else if (resolution.State == ResolutionState.Present)
            {
                if (resolution.Value.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
                {
                    if (!RecordTypeError(block, context, "If requires a JSON Boolean.")) return false;
                }
                else if (resolution.Value.ValueKind == JsonValueKind.True) selected = block.TrueBody;
            }
            if (!RenderSelected(plan, block, selected, targetStart, parentMap, context)) return false;
        }

        RemoveFootprint(block, parentMap);
        return true;
    }

    private bool RenderSelected(StoryPlan plan, BlockPlan block, IReadOnlyList<XElement> selected,
        XElement targetStart, Dictionary<XElement, XElement> parentMap, MergeContext context)
    {
        if (selected.Count == 0) return true;
        var cloneMap = new Dictionary<XElement, XElement>();
        var clones = context.Clone(selected, cloneMap);
        var holder = new XElement("holder", clones);
        var selectedParagraphs = plan.Paragraphs.Where(paragraph => selected.Any(root =>
            paragraph.Paragraph == root || paragraph.Paragraph.Ancestors().Contains(root))).ToArray();
        var selectedBlocks = block.Children.Where(child => selected.Any(root =>
            child.StartContainer == root || child.StartContainer.Ancestors().Contains(root))).ToArray();
        if (!RenderScope(plan, selectedParagraphs, selectedBlocks, cloneMap, context)) return false;
        targetStart.AddBeforeSelf(holder.Elements().ToArray());
        return true;
    }

    private bool RenderParagraph(ParagraphPlan paragraph, Dictionary<XElement, XElement> map, MergeContext context)
    {
        var tokens = paragraph.Tokens.Where(token => token.Expression?.Kind is TemplateExpressionKind.Field ||
            token.Kind is TemplateTokenKind.LiteralOpenBrace or TemplateTokenKind.LiteralCloseBrace).ToArray();
        if (tokens.Length == 0) return true;
        var builders = paragraph.TextNodes.Select(_ => new StringBuilder()).ToArray();
        var tokenIndex = 0;
        for (var offset = 0; offset < paragraph.Text.Length;)
        {
            if (tokenIndex < tokens.Length && tokens[tokenIndex].Start == offset)
            {
                var token = tokens[tokenIndex++];
                var nodeIndex = NodeAt(paragraph.TextNodes, offset);
                var replacement = token.Kind switch
                {
                    TemplateTokenKind.LiteralOpenBrace => "{",
                    TemplateTokenKind.LiteralCloseBrace => "}",
                    _ => FieldText(token.Expression!, paragraph.Location, context)
                };
                builders[nodeIndex].Append(replacement);
                offset += token.Length;
                continue;
            }
            builders[NodeAt(paragraph.TextNodes, offset)].Append(paragraph.Text[offset]);
            offset++;
        }
        foreach (var (span, index) in paragraph.TextNodes.Select((span, index) => (span, index)))
            if (map.TryGetValue(span.Node, out var target))
            {
                target.Value = builders[index].ToString();
                if (target.Value.Length != 0 && (char.IsWhiteSpace(target.Value[0]) || char.IsWhiteSpace(target.Value[^1])))
                    target.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            }
        return !context.LimitReached;
    }

    private string FieldText(TemplateExpression expression, string location, MergeContext context)
    {
        var resolution = _resolver.Resolve(expression.Path!, context);
        if (resolution.State == ResolutionState.Missing)
        {
            var name = expression.Source;
            var message = $"Unknown template value: {name}";
            if (context.Options.MissingValues == MissingValueBehavior.Warning)
                context.Warnings.Add(new(message, name, expression.Source, location, Part(location),
                    _resolver.DataPath(expression.Path!, context)));
            else if (context.Options.MissingValues == MissingValueBehavior.Error)
                context.Record(new TemplateException(message, name, expression.Source, location),
                    _resolver.DataPath(expression.Path!, context), location);
            return string.Empty;
        }
        try
        {
            return _formatter.Format(resolution.Value, expression, context.Options, location);
        }
        catch (TemplateException error)
        {
            context.Record(error, _resolver.DataPath(expression.Path!, context), location);
            return string.Empty;
        }
    }

    private static int NodeAt(IReadOnlyList<TextNodeSpan> spans, int offset)
    {
        for (var index = 0; index < spans.Count; index++)
            if (offset >= spans[index].Start && offset < spans[index].Start + spans[index].Length) return index;
        return Math.Max(0, spans.Count - 1);
    }

    private static bool Contains(BlockPlan block, XElement element)
    {
        if (element == block.StartContainer || element == block.EndContainer || element == block.ElseContainer) return true;
        return block.TrueBody.Concat(block.FalseBody).Any(root => element == root || element.Ancestors().Contains(root));
    }

    private static void RemoveFootprint(BlockPlan block, Dictionary<XElement, XElement> map)
    {
        if (!map.TryGetValue(block.StartContainer, out var start) ||
            !map.TryGetValue(block.EndContainer, out var end) || start.Parent != end.Parent) return;
        RemoveAndEnsureCellStructure(start.Parent!.Elements().SkipWhile(element => element != start)
            .TakeWhile(element => element != end).Append(end).ToArray());
    }

    private static void RemoveAndEnsureCellStructure(IReadOnlyCollection<XElement> elements)
    {
        var cells = elements.Select(element => element.Parent)
            .Where(parent => parent?.Name == WordXml.TableCellElementName)
            .Cast<XElement>().Distinct().ToArray();
        foreach (var element in elements) element.Remove();
        foreach (var cell in cells)
        {
            // Range endpoints and other metadata do not supply cell block content.
            var w = WordXml.WordprocessingNamespace;
            var lastBlock = cell.Elements().LastOrDefault(element =>
                element.Name == WordXml.ParagraphElementName || element.Name == WordXml.TableElementName ||
                element.Name == w + "sdt" || element.Name == w + "customXml" || element.Name == w + "altChunk");
            if (lastBlock is null || lastBlock.Name == WordXml.TableElementName)
                cell.Add(new XElement(WordXml.ParagraphElementName));
        }
    }

    private void WarnMissing(BlockPlan block, MergeContext context)
    {
        if (!context.WarningKeys.Add(block.Location + "|" + block.StartExpression.Source)) return;
        context.Warnings.Add(new($"Missing structural path: {block.StartExpression.Source}",
            block.StartExpression.Source, block.StartExpression.Source, block.Location, Part(block.Location),
            _resolver.DataPath(block.StartExpression.Path!, context)));
    }

    private static string Part(string location) => location.Split(',', 2)[0];

    private bool RecordTypeError(BlockPlan block, MergeContext context, string detail)
    {
        var error = new TemplateException(detail, block.StartExpression.Source, block.StartExpression.Source,
            block.Location, code: "generation.invalid_data");
        return context.Record(error, _resolver.DataPath(block.StartExpression.Path!, context), block.Location);
    }
}
