using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;

namespace MonoReport.Services;

public sealed partial class DocxTemplatePlanner(TemplateLanguageParser parser)
{
    private readonly TemplateLanguageParser _parser = parser ?? throw new ArgumentNullException(nameof(parser));

    internal StoryPlan Plan(OpenXmlPart part, XDocument xml)
    {
        var partName = part.Uri.ToString();
        var diagnostics = new List<StructuralDiagnostic>();
        var paragraphs = new List<ParagraphPlan>();
        var sourceElements = xml.Root!.DescendantsAndSelf().ToArray();
        var sourceOrders = sourceElements.Select((element, index) => (element, index))
            .ToDictionary(item => item.element, item => item.index);
        var paragraphNumber = 0;
        var fieldDepth = 0;
        XElement? previousRoot = null;
        foreach (var paragraph in xml.Root!.Descendants(WordXml.ParagraphElementName))
        {
            paragraphNumber++;
            var storyRoot = IndependentRoot(paragraph);
            if (storyRoot != previousRoot) fieldDepth = 0;
            previousRoot = storyRoot;
            if (IsSpecialized(paragraph))
            {
                // Controls are opaque to syntax, but their field delimiters still enclose outside text.
                if (!paragraph.Ancestors().Any(ancestor => ancestor.Name == WordXml.DrawingElementName ||
                        ancestor.Name == WordXml.PictureElementName))
                    foreach (var character in paragraph.Descendants(WordXml.FieldCharacterElementName))
                        fieldDepth = UpdatedFieldDepth(fieldDepth, character);
                continue;
            }
            var location = Location(xml.Root, paragraph, partName, paragraphNumber);
            foreach (var textNodes in EligibleTextGroups(paragraph, fieldDepth, out fieldDepth))
            {
                var sourceOrder = (long)sourceOrders[textNodes[0]] << 32;
                var text = string.Concat(textNodes.Select(node => node.Value));
                var spans = new List<TextNodeSpan>();
                var offset = 0;
                foreach (var node in textNodes)
                {
                    spans.Add(new(node, offset, node.Value.Length, sourceOrders[node]));
                    offset += node.Value.Length;
                }
                paragraphs.Add(new(paragraph, text, spans, [], location, sourceOrder));
            }
        }

        paragraphs = ParseOutsideRaw(paragraphs, diagnostics);
        diagnostics.AddRange(RevisionDiagnostics(xml.Root, partName, sourceOrders));
        var blocks = BuildBlocks(xml.Root, paragraphs, diagnostics);
        ValidateRawDeletion(blocks, diagnostics, sourceOrders);
        var dependencies = BuildDependencies(paragraphs, diagnostics, partName);
        return new(part, xml, paragraphs, blocks, dependencies,
            diagnostics.OrderBy(item => item.SourceOrder).Select(item => item with
            {
                SourceContext = SourceContext(sourceElements, item.SourceOrder)
            }).ToArray());
    }

    internal DocxIdentityPlan InspectIdentities(StoryInspection package, IReadOnlyList<StoryPlan> stories)
    {
        var locations = new Dictionary<XElement, IdentityLocation>();
        foreach (var (part, xml) in package.PackageXml)
        {
            var blocks = stories.FirstOrDefault(story => story.Part == part)?.Blocks ?? [];
            var elements = xml.Root!.DescendantsAndSelf().ToArray();
            var paragraphNumbers = elements.Where(element => element.Name == WordXml.ParagraphElementName)
                .Select((element, index) => (element, index)).ToDictionary(item => item.element, item => item.index + 1);
            foreach (var (element, index) in elements.Select((element, index) => (element, index)))
            {
                var root = element.AncestorsAndSelf().FirstOrDefault(ancestor =>
                    ancestor.Name == WordXml.WordprocessingNamespace + "footnote" ||
                    ancestor.Name == WordXml.WordprocessingNamespace + "endnote" ||
                    ancestor.Name == WordXml.WordprocessingNamespace + "comment" ||
                    ancestor.Name == WordXml.WordprocessingNamespace + "docPart") ?? xml.Root;
                var paragraph = element.AncestorsAndSelf(WordXml.ParagraphElementName).FirstOrDefault();
                var location = paragraph is null ? StructuralLocation(xml.Root, element, part.Uri.ToString()) :
                    Location(xml.Root, paragraph, part.Uri.ToString(), paragraphNumbers[paragraph]);
                // Source XML ordinal also distinguishes auxiliary definitions with identical paragraph coordinates.
                locations.Add(element, new(element, part.Uri.ToString(), location + ", XML element " + (index + 1),
                    (long)index << 32, root, Ownership(element, blocks)));
            }
        }
        return ValidateIdentities(package, stories, locations);
    }

    private static string Bound(string? value) => value is null ? "" : value[..Math.Min(value.Length, 80)];

    private static string CommentIdentity(XElement element)
    {
        var value = (string?)element.Attribute(WordXml.WordprocessingNamespace + "id") ?? "";
        return int.TryParse(value, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var id)
            ? id.ToString(System.Globalization.CultureInfo.InvariantCulture) : value;
    }

    private static IReadOnlyList<StructuralOwner> Ownership(XElement element, IReadOnlyList<BlockPlan> blocks)
    {
        foreach (var block in blocks)
        {
            if (Within(element, block.StartContainer) || Within(element, block.EndContainer) ||
                (block.ElseContainer is not null && Within(element, block.ElseContainer)))
                return [new(block, element.AncestorsAndSelf().Contains(block.StartContainer) ? "marker-start" :
                    element.AncestorsAndSelf().Contains(block.EndContainer) ? "marker-end" : "marker-else")];
            var branch = block.TrueBody.Any(root => Within(element, root)) ? "true" :
                block.FalseBody.Any(root => Within(element, root)) ? "false" : null;
            if (branch is not null) return [new(block, branch), .. Ownership(element, block.Children)];
        }
        return [];
    }

    private static bool Within(XElement element, XElement root) => element == root || element.Ancestors().Contains(root);

    private static IReadOnlyList<BlockPlan> BuildBlocks(XElement root, IReadOnlyList<ParagraphPlan> paragraphs,
        List<StructuralDiagnostic> diagnostics)
    {
        var markers = new List<Marker>();
        foreach (var paragraph in paragraphs)
        {
            foreach (var token in paragraph.Tokens.Where(token => token.Expression is { Kind: not TemplateExpressionKind.Field }))
            {
                var expression = token.Expression!;
                if (!MarkerOnlyContainer(paragraph.Paragraph, expression.Source))
                {
                    diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidStructure,
                        "Control expressions must be the only visible content in their marker paragraph or row.",
                        expression, paragraph.Location, paragraph.PositionAt(token.Start)));
                    continue;
                }
                markers.Add(new(expression, paragraph, paragraph.PositionAt(token.Start)));
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
                case TemplateExpressionKind.Raw:
                    stack.Push(new(marker, stack.TryPeek(out var parent) ? parent : null));
                    break;
                case TemplateExpressionKind.Else:
                    if (!stack.TryPeek(out var conditional) || conditional.Start.Expression.Kind != TemplateExpressionKind.If ||
                        conditional.Else is not null)
                        diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidStructure, "Else must occur once inside the nearest open If.",
                            marker.Expression, marker.Paragraph.Location, marker.SourceOrder));
                    else conditional.Else = marker;
                    break;
                case TemplateExpressionKind.EndFor:
                case TemplateExpressionKind.EndIf:
                case TemplateExpressionKind.EndRaw:
                    if (!stack.TryPop(out var opened) || marker.Expression.Kind != (opened.Start.Expression.Kind switch
                        {
                            TemplateExpressionKind.For => TemplateExpressionKind.EndFor,
                            TemplateExpressionKind.If => TemplateExpressionKind.EndIf,
                            _ => TemplateExpressionKind.EndRaw
                        }))
                    {
                        diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidStructure, "Closing marker does not match the open block.",
                            marker.Expression, marker.Paragraph.Location, marker.SourceOrder));
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
            diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidStructure, "Block has no matching closing marker. Add a dedicated closing paragraph or row under the same parent.",
                unmatched.Start.Expression, unmatched.Start.Paragraph.Location,
                unmatched.Start.SourceOrder));
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
            if (startRow is null || endRow is null || startRow == endRow || startRow.Parent != endRow.Parent ||
                (elseRow is not null && elseRow.Parent != startRow.Parent) ||
                !MarkerOnlyRow(startRow, opened.Start.Expression.Source) ||
                !MarkerOnlyRow(endRow, end.Expression.Source) ||
                (elseRow is not null && !MarkerOnlyRow(elseRow, opened.Else!.Expression.Source)))
            {
                diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidStructure,
                    "Block markers must delimit sibling paragraphs or dedicated sibling rows under one parent.",
                    opened.Start.Expression, opened.Start.Paragraph.Location,
                    opened.Start.SourceOrder));
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
            diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidStructure, "Block markers are crossing or misordered.",
                opened.Start.Expression, opened.Start.Paragraph.Location,
                opened.Start.SourceOrder));
            return null;
        }
        if (opened.Start.Expression.Kind != TemplateExpressionKind.Raw &&
            startContainer.Name == WordXml.TableRowElementName &&
            siblings.Skip(startIndex).Take(endIndex - startIndex + 1).Any(row =>
                RowHasVerticalMerge(row)))
        {
            diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidStructure,
                "Row blocks cannot contain or intersect vertical merges.", opened.Start.Expression,
                opened.Start.Paragraph.Location, opened.Start.SourceOrder));
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
                        diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidContext, "The current-element expression is valid only inside For.",
                            expression, paragraph.Location, paragraph.PositionAt(token.Start)));
                        break;
                    }
                    if (TryCanonical(expression.Path, loops, out var fieldPath))
                        dependencies.Add(new("field", expression, fieldPath, "value", part, paragraph.Location,
                            paragraph.PositionAt(token.Start)));
                    else diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidContext, "The path moves beyond the available loop context.",
                        expression, paragraph.Location, paragraph.PositionAt(token.Start)));
                    break;
                case TemplateExpressionKind.For:
                    if (!TryCanonical(expression.Path!, loops, out var arrayPath))
                    {
                        diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidContext, "The path moves beyond the available loop context.",
                            expression, paragraph.Location, paragraph.SourceOrder + token.Start));
                        break;
                    }
                    dependencies.Add(new("for", expression, arrayPath, "array", part, paragraph.Location,
                        paragraph.SourceOrder + token.Start));
                    loops.Add(arrayPath + "[]");
                    controls.Push(TemplateExpressionKind.For);
                    break;
                case TemplateExpressionKind.If:
                    if (TryCanonical(expression.Path!, loops, out var booleanPath))
                        dependencies.Add(new("if", expression, booleanPath, "boolean", part, paragraph.Location,
                            paragraph.PositionAt(token.Start)));
                    else diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidContext, "The path moves beyond the available loop context.",
                        expression, paragraph.Location, paragraph.PositionAt(token.Start)));
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

    private static IReadOnlyList<IReadOnlyList<XElement>> EligibleTextGroups(XElement paragraph,
        int initialFieldDepth, out int finalFieldDepth)
    {
        var fieldDepth = initialFieldDepth;
        var groups = new List<IReadOnlyList<XElement>>();
        var current = new List<XElement>();
        foreach (var element in paragraph.Descendants())
        {
            if (element.Ancestors().TakeWhile(ancestor => ancestor != paragraph).Any(IsOpaque)) continue;
            if (IsOpaque(element))
            {
                Flush();
                if (element.Name == WordXml.HyperlinkElementName || element.Name == WordXml.WordprocessingNamespace + "sdt")
                    foreach (var fieldCharacter in element.Descendants(WordXml.FieldCharacterElementName))
                        fieldDepth = UpdatedFieldDepth(fieldDepth, fieldCharacter);
                continue;
            }
            if (element.Name == WordXml.FieldCharacterElementName)
            {
                Flush();
                fieldDepth = UpdatedFieldDepth(fieldDepth, element);
                continue;
            }
            if (element.Name == WordXml.TextElementName)
            {
                if (fieldDepth != 0) Flush();
                else current.Add(element);
                continue;
            }

            // Formatting and explicitly recognized invisible metadata preserve text continuity.
            // Every other content-level element is a boundary that remains opaque to the parser.
            // This includes symbols, tabs, breaks, references, and unsupported wrappers.
            if (element.Name == WordXml.RunElementName || IsInvisibleMetadata(element) || IsTextPropertyContainer(element) ||
                element.Ancestors().TakeWhile(ancestor => ancestor != paragraph).Any(IsTextPropertyContainer))
                continue;
            Flush();
        }
        Flush();
        finalFieldDepth = fieldDepth;
        return groups;

        void Flush()
        {
            if (current.Count == 0) return;
            groups.Add(current.ToArray());
            current.Clear();
        }
    }

    private static int UpdatedFieldDepth(int fieldDepth, XElement fieldCharacter)
    {
        var type = (string?)fieldCharacter.Attribute(WordXml.FieldCharacterTypeAttributeName);
        if (type == WordXml.FieldBeginValue) return fieldDepth + 1;
        return type == WordXml.FieldEndValue && fieldDepth > 0 ? fieldDepth - 1 : fieldDepth;
    }

    private static bool IsSpecialized(XElement paragraph) => paragraph.Ancestors().Any(ancestor =>
        ancestor.Name == WordXml.DrawingElementName || ancestor.Name == WordXml.PictureElementName ||
        ancestor.Name == WordXml.WordprocessingNamespace + "sdt");

    private static bool MarkerOnlyRow(XElement row, string marker) =>
        MarkerOnlyContainer(row, marker);

    private static bool MarkerOnlyContainer(XElement container, string marker) =>
        container.Descendants().All(element => IsAllowedMarkerElement(element, container)) &&
        string.Concat(container.Descendants(WordXml.TextElementName).Select(text => text.Value)).Trim() == marker;

    private static bool IsAllowedMarkerElement(XElement element, XElement container)
    {
        if (element.Ancestors().TakeWhile(ancestor => ancestor != container).Any(IsMarkerPropertyContainer)) return true;
        return (container.Name == WordXml.TableRowElementName &&
                element.Name == WordXml.TableCellElementName) ||
            element.Name == WordXml.ParagraphElementName ||
            element.Name == WordXml.RunElementName || element.Name == WordXml.TextElementName ||
            IsInvisibleMetadata(element) ||
            element.Name == WordXml.WordprocessingNamespace + "commentRangeStart" ||
            element.Name == WordXml.WordprocessingNamespace + "commentRangeEnd" ||
            element.Name == WordXml.WordprocessingNamespace + "commentReference" ||
            IsMarkerPropertyContainer(element);
    }

    private static bool IsInvisibleRangeMarker(XElement element) =>
        element.Name == WordXml.WordprocessingNamespace + "bookmarkStart" ||
        element.Name == WordXml.WordprocessingNamespace + "bookmarkEnd" ||
        element.Name == WordXml.WordprocessingNamespace + "permStart" ||
        element.Name == WordXml.WordprocessingNamespace + "permEnd";

    private static bool IsInvisibleMetadata(XElement element) => IsInvisibleRangeMarker(element) ||
        element.Name == WordXml.WordprocessingNamespace + "proofErr" ||
        element.Name == WordXml.WordprocessingNamespace + "lastRenderedPageBreak";

    private static string? SourceContext(IReadOnlyList<XElement> elements, long sourceOrder)
    {
        var index = (int)(sourceOrder >> 32);
        if (index < 0 || index >= elements.Count) return null;
        var source = elements[index];
        var paragraph = source.AncestorsAndSelf(WordXml.ParagraphElementName).FirstOrDefault();
        if (paragraph is null) return null;
        var text = new System.Text.StringBuilder();
        var offset = 0;
        foreach (var element in paragraph.Descendants())
        {
            if (element == source) offset = text.Length + (int)(sourceOrder & uint.MaxValue);
            if (element.Name == WordXml.TextElementName) text.Append(element.Value);
            else if (element.Name == WordXml.WordprocessingNamespace + "tab" ||
                     element.Name == WordXml.WordprocessingNamespace + "br") text.Append(' ');
        }
        var start = Math.Max(0, Math.Min(offset, text.Length) - 60);
        var length = Math.Min(160, text.Length - start);
        return (start > 0 ? "…" : "") + text.ToString(start, length) + (start + length < text.Length ? "…" : "");
    }

    private static bool RowHasVerticalMerge(XElement row) => row.Elements(WordXml.TableCellElementName)
        .SelectMany(cell => cell.Elements(WordXml.TableCellPropertiesElementName))
        .Any(properties => properties.Descendants(WordXml.VerticalMergeElementName).Any());

    private static bool IsTextPropertyContainer(XElement element) =>
        element.Name == WordXml.ParagraphPropertiesElementName ||
        element.Name == WordXml.RunPropertiesElementName;

    private static bool IsMarkerPropertyContainer(XElement element) =>
        element.Name == WordXml.TableRowPropertiesElementName ||
        element.Name == WordXml.TableCellPropertiesElementName ||
        element.Name == WordXml.ParagraphPropertiesElementName ||
        element.Name == WordXml.RunPropertiesElementName;

    private static bool IsOpaque(XElement element) =>
        element.Name == WordXml.HyperlinkElementName || element.Name == WordXml.SimpleFieldElementName ||
        element.Name == WordXml.DrawingElementName || element.Name == WordXml.PictureElementName ||
        element.Name == WordXml.WordprocessingNamespace + "sdt";

    // Qualified names are explicit: range base classes and suffixes also include non-revisions.
    private static readonly System.Collections.Frozen.FrozenSet<XName> RevisionNames =
        System.Collections.Frozen.FrozenSet.ToFrozenSet(
            new[]
            {
                // Content, cell changes, and deleted payloads.
                "ins", "del", "moveFrom", "moveTo", "cellIns", "cellDel", "cellMerge", "delText", "delInstrText",
                // Both ends of move and custom-XML revision ranges.
                "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
                "customXmlInsRangeStart", "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd",
                "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd", "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
                // Property revisions, including paragraph marks and table exceptions.
                "rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "tblGridChange", "trPrChange", "tcPrChange",
                "numberingChange", "tblPrExChange"
            }.Select(name => WordXml.WordprocessingNamespace + name).Concat(new[]
            {
                "conflictIns", "conflictDel", "customXmlConflictInsRangeStart", "customXmlConflictInsRangeEnd",
                "customXmlConflictDelRangeStart", "customXmlConflictDelRangeEnd"
            }.Select(name => XNamespace.Get("http://schemas.microsoft.com/office/word/2010/wordml") + name)));

    private static IEnumerable<StructuralDiagnostic> RevisionDiagnostics(XElement root, string part,
        IReadOnlyDictionary<XElement, int> sourceOrders)
    {
        foreach (var revision in root.Descendants().Where(element => RevisionNames.Contains(element.Name)))
        {
            if (revision.Ancestors().Any(ancestor => ancestor.Name == WordXml.DrawingElementName ||
                    ancestor.Name == WordXml.PictureElementName)) continue;
            var paragraph = revision.AncestorsAndSelf(WordXml.ParagraphElementName).FirstOrDefault();
            var location = paragraph is null
                ? StructuralLocation(root, revision, part)
                : Location(root, paragraph, part,
                    root.Descendants(WordXml.ParagraphElementName).TakeWhile(item => item != paragraph).Count() + 1);
            yield return new(DiagnosticCodes.TemplateTrackedRevision, "Tracked revisions must be accepted or rejected before processing.",
                Location: location + ", tracked revision " + revision.Name.LocalName,
                SourceOrder: (long)sourceOrders[revision] << 32);
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

    private static StructuralDiagnostic Error(string code, string detail, TemplateExpression expression,
        string location, long sourceOrder) =>
        new(code, detail, expression.Source, expression.Source, location, SourceOrder: sourceOrder, Explanation: detail);

    private sealed record Marker(TemplateExpression Expression, ParagraphPlan Paragraph, long SourceOrder);

    private sealed class OpenMarker(Marker start, OpenMarker? parent)
    {
        internal Marker Start { get; } = start;
        internal OpenMarker? Parent { get; } = parent;
        internal Marker? Else { get; set; }
        internal List<BlockPlan> Children { get; } = [];
    }
}
