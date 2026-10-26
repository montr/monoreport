using System.Collections.Immutable;
using System.Xml.Linq;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxTemplatePlanner
{
    private XlsxPreparedValue TextValue(XlsxCellPlan cell, string text, IReadOnlyList<XlsxTextEdit> edits)
    {
        if (cell.Source.RichString is not { } source) return new("inlineStr", text);
        var spans = ImmutableArray.CreateBuilder<XlsxTextSpan>();
        var sourceSpan = 0;
        var outputOffset = 0;
        var sourceOffset = 0;
        void Advance(int position)
        {
            while (sourceSpan < source.Spans.Length && source.Spans[sourceSpan].Start + source.Spans[sourceSpan].Text.Length <= position) sourceSpan++;
        }
        void Append(string value, XElement? properties)
        {
            if (value.Length == 0) return;
            spans.Add(new(outputOffset, value, properties is null ? null : new XElement(properties)));
            outputOffset = checked(outputOffset + value.Length);
        }
        void Literal(int end)
        {
            Advance(sourceOffset);
            while (sourceOffset < end && sourceSpan < source.Spans.Length)
            {
                var span = source.Spans[sourceSpan];
                var count = Math.Min(end - sourceOffset, span.Start + span.Text.Length - sourceOffset);
                Append(span.Text.Substring(sourceOffset - span.Start, count), span.Properties);
                sourceOffset += count;
                Advance(sourceOffset);
            }
        }
        // The same parsed-token edits drive visible output, style inheritance and annotation relocation.
        var deltas = new int[edits.Count + 1];
        for (var index = 0; index < edits.Count; index++)
        {
            var edit = edits[index];
            Literal(edit.Start);
            Advance(edit.Start);
            Append(edit.Text, sourceSpan < source.Spans.Length ? source.Spans[sourceSpan].Properties : null);
            sourceOffset = edit.Start + edit.Length;
            deltas[index + 1] = checked(deltas[index] + edit.Text.Length - edit.Length);
        }
        Literal(cell.Source.Text!.Length);
        var phonetics = ImmutableArray.CreateBuilder<XElement>();
        foreach (var hint in source.Phonetics)
        {
            var start = (int)hint.Attribute("sb")!;
            var end = (int)hint.Attribute("eb")!;
            // Find the first edit not wholly before the annotation; no character rescans or whole-string diff.
            var low = 0;
            var high = edits.Count;
            while (low < high)
            {
                var mid = low + (high - low) / 2;
                if (edits[mid].Start + edits[mid].Length <= start) low = mid + 1; else high = mid;
            }
            if (low < edits.Count && edits[low].Start < end) continue;
            var copy = new XElement(hint);
            copy.SetAttributeValue("sb", checked(start + deltas[low]));
            copy.SetAttributeValue("eb", checked(end + deltas[low]));
            phonetics.Add(copy);
        }
        return new("inlineStr", text)
        {
            RichString = new(text, spans.ToImmutable(), phonetics.ToImmutable(),
                source.PhoneticProperties is null ? null : new XElement(source.PhoneticProperties))
        };
    }
}
