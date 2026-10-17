using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;

namespace MonoReport.Services;

public sealed class DocxTemplatePlanner(TemplateLanguageParser parser)
{
    private readonly TemplateLanguageParser _parser = parser ?? throw new ArgumentNullException(nameof(parser));

    internal StoryPlan Plan(OpenXmlPart part, XDocument xml)
    {
        var partName = part.Uri.ToString();
        var diagnostics = new List<StructuralDiagnostic>();
        var paragraphs = new List<ParagraphPlan>();
        var paragraphNumber = 0;
        foreach (var paragraph in xml.Root!.Descendants(WordXml.ParagraphElementName))
        {
            paragraphNumber++;
            if (IsSpecialized(paragraph)) continue;
            var location = Location(xml.Root, paragraph, partName, paragraphNumber);
            foreach (var textNodes in EligibleTextGroups(paragraph))
            {
                var text = string.Concat(textNodes.Select(node => node.Value));
                var spans = new List<TextNodeSpan>();
                var offset = 0;
                foreach (var node in textNodes)
                {
                    spans.Add(new(node, offset, node.Value.Length));
                    offset += node.Value.Length;
                }
                var parsed = _parser.Parse(text, location);
                diagnostics.AddRange(parsed.Diagnostics);
                paragraphs.Add(new(paragraph, text, spans, parsed.Tokens, location));
            }
        }

        diagnostics.AddRange(RevisionDiagnostics(xml.Root, partName));
        var blocks = BuildBlocks(xml.Root, paragraphs, diagnostics);
        var dependencies = BuildDependencies(paragraphs, diagnostics, partName);
        return new(part, xml, paragraphs, blocks, dependencies, diagnostics);
    }

    private static IReadOnlyList<BlockPlan> BuildBlocks(XElement root, IReadOnlyList<ParagraphPlan> paragraphs,
        List<StructuralDiagnostic> diagnostics)
    {
        var markers = new List<Marker>();
        foreach (var paragraph in paragraphs)
        {
            foreach (var token in paragraph.Tokens.Where(token => token.Expression is { Kind: not TemplateExpressionKind.Field }))
            {
                var expression = token.Expression!;
                if (VisibleParagraphText(paragraph.Paragraph).Trim() != expression.Source)
                {
                    diagnostics.Add(Error("template.invalid_structure",
                        "Control expressions must be the only visible content in their marker paragraph or row.",
                        expression, paragraph.Location));
                    continue;
                }
                markers.Add(new(expression, paragraph));
            }
        }

        var top = new List<BlockPlan>();
        var stack = new Stack<OpenMarker>();
        foreach (var marker in markers)
        {
            switch (marker.Expression.Kind)
            {
                case TemplateExpressionKind.For:
                case TemplateExpressionKind.If:
                    stack.Push(new(marker, stack.TryPeek(out var parent) ? parent : null));
                    break;
                case TemplateExpressionKind.Else:
                    if (!stack.TryPeek(out var conditional) || conditional.Start.Expression.Kind != TemplateExpressionKind.If ||
                        conditional.Else is not null)
                        diagnostics.Add(Error("template.invalid_structure", "Else must occur once inside the nearest open If.",
                            marker.Expression, marker.Paragraph.Location));
                    else conditional.Else = marker;
                    break;
                case TemplateExpressionKind.EndFor:
                case TemplateExpressionKind.EndIf:
                    if (!stack.TryPop(out var opened) ||
                        (marker.Expression.Kind == TemplateExpressionKind.EndFor) !=
                        (opened.Start.Expression.Kind == TemplateExpressionKind.For))
                    {
                        diagnostics.Add(Error("template.invalid_structure", "Closing marker does not match the open block.",
                            marker.Expression, marker.Paragraph.Location));
                        break;
                    }
                    var block = CreateBlock(root, opened, marker, diagnostics);
                    if (block is null) break;
                    if (opened.Parent is null) top.Add(block);
                    else opened.Parent.Children.Add(block);
                    break;
            }
        }
        foreach (var unmatched in stack)
            diagnostics.Add(Error("template.invalid_structure", "Block has no matching closing marker.",
                unmatched.Start.Expression, unmatched.Start.Paragraph.Location));
        return top;
    }

    private static BlockPlan? CreateBlock(XElement root, OpenMarker opened, Marker end,
        List<StructuralDiagnostic> diagnostics)
    {
        var startParagraph = opened.Start.Paragraph.Paragraph;
        var endParagraph = end.Paragraph.Paragraph;
        var elseParagraph = opened.Else?.Paragraph.Paragraph;
        XElement startContainer;
        XElement endContainer;
        XElement? elseContainer;
        if (startParagraph.Parent == endParagraph.Parent &&
            (elseParagraph is null || elseParagraph.Parent == startParagraph.Parent))
        {
            startContainer = startParagraph;
            endContainer = endParagraph;
            elseContainer = elseParagraph;
        }
        else
        {
            var startRow = startParagraph.Ancestors(WordXml.TableRowElementName).FirstOrDefault();
            var endRow = endParagraph.Ancestors(WordXml.TableRowElementName).FirstOrDefault();
            var elseRow = elseParagraph?.Ancestors(WordXml.TableRowElementName).FirstOrDefault();
            if (startRow is null || endRow is null || startRow.Parent != endRow.Parent ||
                (elseRow is not null && elseRow.Parent != startRow.Parent) ||
                !MarkerOnlyRow(startRow, opened.Start.Expression.Source) ||
                !MarkerOnlyRow(endRow, end.Expression.Source) ||
                (elseRow is not null && !MarkerOnlyRow(elseRow, opened.Else!.Expression.Source)))
            {
                diagnostics.Add(Error("template.invalid_structure",
                    "Block markers must delimit sibling paragraphs or dedicated sibling rows under one parent.",
                    opened.Start.Expression, opened.Start.Paragraph.Location));
                return null;
            }
            startContainer = startRow;
            endContainer = endRow;
            elseContainer = elseRow;
        }

        var siblings = startContainer.Parent!.Elements().ToList();
        var startIndex = siblings.IndexOf(startContainer);
        var endIndex = siblings.IndexOf(endContainer);
        var elseIndex = elseContainer is null ? -1 : siblings.IndexOf(elseContainer);
        if (startIndex < 0 || endIndex <= startIndex || (elseIndex >= 0 && (elseIndex <= startIndex || elseIndex >= endIndex)))
        {
            diagnostics.Add(Error("template.invalid_structure", "Block markers are crossing or misordered.",
                opened.Start.Expression, opened.Start.Paragraph.Location));
            return null;
        }
        if (startContainer.Name == WordXml.TableRowElementName &&
            siblings.Skip(startIndex).Take(endIndex - startIndex + 1).Any(row =>
                row.Descendants(WordXml.VerticalMergeElementName).Any()))
        {
            diagnostics.Add(Error("template.invalid_structure",
                "Row blocks cannot contain or intersect vertical merges.", opened.Start.Expression,
                opened.Start.Paragraph.Location));
            return null;
        }
        var trueEnd = elseIndex >= 0 ? elseIndex : endIndex;
        var trueBody = siblings.Skip(startIndex + 1).Take(trueEnd - startIndex - 1).ToArray();
        var falseBody = elseIndex < 0 ? [] : siblings.Skip(elseIndex + 1).Take(endIndex - elseIndex - 1).ToArray();
        var result = new BlockPlan(opened.Start.Expression.Kind, opened.Start.Expression, startContainer, endContainer,
            elseContainer, trueBody, falseBody, opened.Start.Paragraph.Location, opened.Children.ToArray());
        return result;
    }

    private static IReadOnlyList<PlannedDependency> BuildDependencies(IReadOnlyList<ParagraphPlan> paragraphs,
        List<StructuralDiagnostic> diagnostics, string part)
    {
        var dependencies = new List<PlannedDependency>();
        var loops = new List<string>();
        var controls = new Stack<TemplateExpressionKind>();
        foreach (var paragraph in paragraphs)
        foreach (var token in paragraph.Tokens.Where(token => token.Kind == TemplateTokenKind.Expression))
        {
            var expression = token.Expression!;
            switch (expression.Kind)
            {
                case TemplateExpressionKind.Field:
                    if (expression.Path!.CurrentElement && loops.Count == 0)
                    {
                        diagnostics.Add(Error("template.invalid_context", "The current-element expression is valid only inside For.",
                            expression, paragraph.Location));
                        break;
                    }
                    if (TryCanonical(expression.Path, loops, out var fieldPath))
                        dependencies.Add(new("field", expression, fieldPath, "value", part, paragraph.Location));
                    else diagnostics.Add(Error("template.invalid_context", "The path moves beyond the available loop context.",
                        expression, paragraph.Location));
                    break;
                case TemplateExpressionKind.For:
                    if (!TryCanonical(expression.Path!, loops, out var arrayPath))
                    {
                        diagnostics.Add(Error("template.invalid_context", "The path moves beyond the available loop context.",
                            expression, paragraph.Location));
                        break;
                    }
                    dependencies.Add(new("for", expression, arrayPath, "array", part, paragraph.Location));
                    loops.Add(arrayPath + "[]");
                    controls.Push(TemplateExpressionKind.For);
                    break;
                case TemplateExpressionKind.If:
                    if (TryCanonical(expression.Path!, loops, out var booleanPath))
                        dependencies.Add(new("if", expression, booleanPath, "boolean", part, paragraph.Location));
                    else diagnostics.Add(Error("template.invalid_context", "The path moves beyond the available loop context.",
                        expression, paragraph.Location));
                    controls.Push(TemplateExpressionKind.If);
                    break;
                case TemplateExpressionKind.EndFor:
                    if (controls.TryPop(out var forKind) && forKind == TemplateExpressionKind.For && loops.Count != 0)
                        loops.RemoveAt(loops.Count - 1);
                    break;
                case TemplateExpressionKind.EndIf:
                    controls.TryPop(out _);
                    break;
            }
        }
        return dependencies;
    }

    private static bool TryCanonical(TemplatePath path, IReadOnlyList<string> loops, out string canonical)
    {
        if (path.CurrentElement)
        {
            canonical = loops.Count == 0 ? string.Empty : loops[^1];
            return loops.Count != 0;
        }
        string basis;
        if (path.Origin == TemplatePathOrigin.Root) basis = "$";
        else if (path.ParentDepth == 0) basis = loops.Count == 0 ? "$" : loops[^1];
        else
        {
            var index = loops.Count - 1 - path.ParentDepth;
            if (index < -1)
            {
                canonical = string.Empty;
                return false;
            }
            basis = index < 0 ? "$" : loops[index];
        }
        canonical = basis + "." + string.Join('.', path.Segments);
        return true;
    }

    private static IEnumerable<IReadOnlyList<XElement>> EligibleTextGroups(XElement paragraph)
    {
        var fieldDepth = 0;
        var current = new List<XElement>();
        foreach (var element in paragraph.Descendants())
        {
            if (element.Name == WordXml.FieldCharacterElementName)
            {
                if (current.Count != 0) { yield return current.ToArray(); current.Clear(); }
                var type = (string?)element.Attribute(WordXml.FieldCharacterTypeAttributeName);
                if (type == WordXml.FieldBeginValue) fieldDepth++;
                else if (type == WordXml.FieldEndValue && fieldDepth > 0) fieldDepth--;
                continue;
            }
            if (element.Name != WordXml.TextElementName) continue;
            var excluded = fieldDepth != 0 || element.Ancestors().TakeWhile(ancestor => ancestor != paragraph).Any(ancestor =>
                    ancestor.Name == WordXml.HyperlinkElementName ||
                    ancestor.Name == WordXml.SimpleFieldElementName ||
                    ancestor.Name == WordXml.DrawingElementName ||
                    ancestor.Name == WordXml.PictureElementName);
            if (excluded)
            {
                if (current.Count != 0) { yield return current.ToArray(); current.Clear(); }
                continue;
            }
            current.Add(element);
        }
        if (current.Count != 0) yield return current.ToArray();
    }

    private static string VisibleParagraphText(XElement paragraph) => string.Concat(
        paragraph.Descendants(WordXml.TextElementName).Where(text =>
            !text.Ancestors().TakeWhile(ancestor => ancestor != paragraph).Any(ancestor =>
                ancestor.Name == WordXml.DrawingElementName || ancestor.Name == WordXml.PictureElementName))
            .Select(text => text.Value));

    private static bool IsSpecialized(XElement paragraph) => paragraph.Ancestors().Any(ancestor =>
        ancestor.Name == WordXml.DrawingElementName || ancestor.Name == WordXml.PictureElementName);

    private static bool MarkerOnlyRow(XElement row, string marker) =>
        string.Concat(row.Descendants(WordXml.TextElementName).Where(text =>
            !text.Ancestors().Any(ancestor => ancestor.Name == WordXml.HyperlinkElementName ||
                ancestor.Name == WordXml.SimpleFieldElementName || ancestor.Name == WordXml.DrawingElementName ||
                ancestor.Name == WordXml.PictureElementName)).Select(text => text.Value)).Trim() == marker;

    private static IEnumerable<StructuralDiagnostic> RevisionDiagnostics(XElement root, string part)
    {
        var revisionNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "tblPrChange", "trPrChange",
            "tcPrChange", "sectPrChange", "numberingChange", "tblGridChange"
        };
        foreach (var revision in root.Descendants().Where(element =>
                     element.Name.Namespace == WordXml.WordprocessingNamespace && revisionNames.Contains(element.Name.LocalName)))
        {
            if (revision.Ancestors().Any(ancestor => ancestor.Name == WordXml.DrawingElementName ||
                    ancestor.Name == WordXml.PictureElementName)) continue;
            var paragraph = revision.AncestorsAndSelf(WordXml.ParagraphElementName).FirstOrDefault();
            var location = paragraph is null
                ? StructuralLocation(root, revision, part)
                : Location(root, paragraph, part,
                    root.Descendants(WordXml.ParagraphElementName).TakeWhile(item => item != paragraph).Count() + 1);
            yield return new("template.tracked_revision", "Tracked revisions must be accepted or rejected before processing.",
                Location: location + ", tracked revision " + revision.Name.LocalName);
        }
    }

    private static string StructuralLocation(XElement root, XElement element, string part)
    {
        var row = element.AncestorsAndSelf(WordXml.TableRowElementName).FirstOrDefault();
        var table = element.AncestorsAndSelf(WordXml.TableElementName).FirstOrDefault();
        if (table is null) return part + ", document properties";
        var tableNumber = root.Descendants(WordXml.TableElementName).TakeWhile(item => item != table).Count() + 1;
        if (row is null) return $"{part}, table {tableNumber}";
        var rowNumber = table.Elements(WordXml.TableRowElementName).TakeWhile(item => item != row).Count() + 1;
        var cell = element.AncestorsAndSelf(WordXml.TableCellElementName).FirstOrDefault();
        if (cell is null) return $"{part}, table {tableNumber}, row {rowNumber}";
        var cellNumber = row.Elements(WordXml.TableCellElementName).TakeWhile(item => item != cell).Count() + 1;
        return $"{part}, table {tableNumber}, row {rowNumber}, cell {cellNumber}";
    }

    private static string Location(XElement root, XElement paragraph, string part, int paragraphNumber)
    {
        var location = $"{part}, paragraph {paragraphNumber}";
        var row = paragraph.Ancestors(WordXml.TableRowElementName).FirstOrDefault();
        if (row is null) return location;
        var table = row.Ancestors(WordXml.TableElementName).First();
        var tableNumber = root.Descendants(WordXml.TableElementName).TakeWhile(item => item != table).Count() + 1;
        var rowNumber = table.Elements(WordXml.TableRowElementName).TakeWhile(item => item != row).Count() + 1;
        var cell = paragraph.Ancestors(WordXml.TableCellElementName).FirstOrDefault();
        var cellNumber = cell is null ? 0 : row.Elements(WordXml.TableCellElementName).TakeWhile(item => item != cell).Count() + 1;
        return $"{location}, table {tableNumber}, row {rowNumber}, cell {cellNumber}";
    }

    private static StructuralDiagnostic Error(string code, string detail, TemplateExpression expression, string location) =>
        new(code, detail, expression.Source, expression.Source, location);

    private sealed record Marker(TemplateExpression Expression, ParagraphPlan Paragraph);

    private sealed class OpenMarker(Marker start, OpenMarker? parent)
    {
        internal Marker Start { get; } = start;
        internal OpenMarker? Parent { get; } = parent;
        internal Marker? Else { get; set; }
        internal List<BlockPlan> Children { get; } = [];
    }
}
