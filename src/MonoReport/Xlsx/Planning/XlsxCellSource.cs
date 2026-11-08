using System.Collections.Immutable;
using System.Xml.Linq;

namespace MonoReport.Xlsx.Planning;

internal sealed record XlsxCellSource(
    SourceCellId Id,
    XElement Xml,
    string? Text,
    string? NativeType,
    bool HasFormula,
    bool RichText,
    XElement? SharedString = null)
{
    internal XlsxRichString? RichString { get; init; }
}

internal sealed record XlsxTextSpan(int Start, string Text, XElement? Properties);

internal sealed record XlsxRichString(
    string Text,
    ImmutableArray<XlsxTextSpan> Spans,
    ImmutableArray<XElement> Phonetics,
    XElement? PhoneticProperties);

internal sealed record XlsxTextEdit(int Start, int Length, string Text);

// RetentionArea is the reference domain: authored crop, or the full worksheet for identity retention.
// It is never a stored-content extent and must not be emitted as a synthetic dimension/PrintArea.