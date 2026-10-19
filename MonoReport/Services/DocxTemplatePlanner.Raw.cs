using System.Xml.Linq;
using MonoReport.Models;

namespace MonoReport.Services;

public sealed partial class DocxTemplatePlanner
{
    private List<ParagraphPlan> ParseOutsideRaw(IReadOnlyList<ParagraphPlan> source,
        List<StructuralDiagnostic> diagnostics)
    {
        var result = new List<ParagraphPlan>();
        ParagraphPlan? opened = null;
        foreach (var paragraph in source)
        {
            if (opened is not null && IndependentRoot(opened.Paragraph) != IndependentRoot(paragraph.Paragraph))
                opened = null; // BuildBlocks diagnoses the unclosed node; never consume another independent root.
            if (opened is not null)
            {
                if (paragraph.Text.Trim() != "{/Raw}" ||
                    !RawContainers(opened.Paragraph, paragraph.Paragraph, out _, out _))
                    continue;
                opened = null;
            }

            var parsed = _parser.Parse(paragraph.Text, paragraph.Location);
            var planned = paragraph with { Tokens = parsed.Tokens };
            diagnostics.AddRange(parsed.Diagnostics.Select(item => item with
                { SourceOrder = planned.PositionAt((int)item.SourceOrder) }));
            result.Add(planned);
            if (parsed.Tokens.Any(token => token.Expression?.Kind == TemplateExpressionKind.Raw) &&
                MarkerOnlyContainer(paragraph.Paragraph, "{Raw}"))
                opened = planned;
        }
        return result;
    }

    private static XElement IndependentRoot(XElement element) => element.AncestorsAndSelf().FirstOrDefault(ancestor =>
        ancestor.Name == WordXml.WordprocessingNamespace + "footnote" ||
        ancestor.Name == WordXml.WordprocessingNamespace + "endnote" ||
        ancestor.Name == WordXml.WordprocessingNamespace + "comment" ||
        ancestor.Name == WordXml.WordprocessingNamespace + "docPart") ?? element.Document!.Root!;

    private static bool RawContainers(XElement start, XElement end, out XElement first, out XElement last)
    {
        first = start;
        last = end;
        if (IndependentRoot(start) != IndependentRoot(end) ||
            !MarkerOnlyContainer(start, "{Raw}") || !MarkerOnlyContainer(end, "{/Raw}")) return false;
        if (start != end && start.Parent == end.Parent) return true;
        var startRow = start.Ancestors(WordXml.TableRowElementName).FirstOrDefault();
        var endRow = end.Ancestors(WordXml.TableRowElementName).FirstOrDefault();
        if (startRow is null || endRow is null || startRow == endRow || startRow.Parent != endRow.Parent ||
            !MarkerOnlyRow(startRow, "{Raw}") || !MarkerOnlyRow(endRow, "{/Raw}")) return false;
        first = startRow;
        last = endRow;
        return true;
    }

    private static void ValidateRawDeletion(IReadOnlyList<BlockPlan> blocks, List<StructuralDiagnostic> diagnostics,
        IReadOnlyDictionary<XElement, int> orders)
    {
        var raw = AllBlocks(blocks).Where(block => block.Kind == TemplateExpressionKind.Raw).ToArray();
        var removed = raw.SelectMany(block => new[] { block.StartContainer, block.EndContainer }).ToHashSet();
        foreach (var block in raw)
        {
            string? problem = null;
            foreach (var delimiter in new[] { block.StartContainer, block.EndContainer })
            {
                if (delimiter.Descendants(WordXml.WordprocessingNamespace + "sectPr").Any())
                    problem = "a section boundary";
                if (delimiter.Parent?.Name == WordXml.TableCellElementName)
                {
                    var surviving = delimiter.Parent.Elements().Where(element => !removed.Contains(element) &&
                        element.Name != WordXml.TableCellPropertiesElementName).ToArray();
                    if (surviving.Length == 0 || surviving[^1].Name != WordXml.ParagraphElementName)
                        problem = "required cell block content or the final cell paragraph";
                }
                if (delimiter.Name != WordXml.TableRowElementName) continue;
                var rows = delimiter.Parent!.Elements(WordXml.TableRowElementName).ToArray();
                var survivingRows = rows.Where(row => !removed.Contains(row)).ToArray();
                if (survivingRows.Length == 0) problem = "all required table rows";
                var before = MergeOrigins(rows);
                var after = MergeOrigins(survivingRows);
                if (after.Any(pair => pair.Value is null || !before.TryGetValue(pair.Key, out var origin) || origin != pair.Value))
                    problem = "a surviving vertical merge chain";
            }
            if (problem is not null)
                diagnostics.Add(Error("template.invalid_structure",
                    $"Removing Raw delimiters would remove or damage {problem}. Move the delimiters to ordinary paragraphs/rows outside that structure in Word; retain required cell paragraphs and section boundaries.",
                    block.StartExpression, block.Location, (long)orders[block.StartContainer] << 32));
        }
    }

    // Compare surviving merge groups, rather than rejecting merges in a body that Raw leaves intact.
    private static Dictionary<XElement, XElement?> MergeOrigins(IEnumerable<XElement> rows)
    {
        var w = WordXml.WordprocessingNamespace;
        var previous = new Dictionary<(int Column, int Span), XElement?>();
        var result = new Dictionary<XElement, XElement?>();
        foreach (var row in rows)
        {
            var current = new Dictionary<(int Column, int Span), XElement?>();
            var column = (int?)row.Element(w + "trPr")?.Element(w + "gridBefore")?.Attribute(w + "val") ?? 0;
            foreach (var cell in row.Elements(WordXml.TableCellElementName))
            {
                var properties = cell.Element(WordXml.TableCellPropertiesElementName);
                var span = (int?)properties?.Element(w + "gridSpan")?.Attribute(w + "val") ?? 1;
                var merge = properties?.Element(WordXml.VerticalMergeElementName);
                if (merge is not null)
                {
                    var origin = (string?)merge.Attribute(w + "val") == "restart" ? cell :
                        previous.GetValueOrDefault((column, span));
                    current[(column, span)] = origin;
                    result[cell] = origin;
                }
                column += span;
            }
            previous = current;
        }
        return result;
    }
}
