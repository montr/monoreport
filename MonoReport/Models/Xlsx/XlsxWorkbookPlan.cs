using System.Collections.Immutable;
using System.Xml.Linq;

namespace MonoReport.Models.Xlsx;

internal enum XlsxRangeImageKind { Cell, Rectangle, Empty, Disjoint, Ambiguous }
// Complete requires PrintArea containment and a compatible image for every requested source address.
internal sealed record XlsxRangeImage(XlsxRangeImageKind Kind, XlsxRectangle? Rectangle, bool Complete);
internal sealed record XlsxAxisSpan(int First, int Last, int OutputFirst,
    VerticalInstancePath? Vertical = null, HorizontalInstancePath? Horizontal = null)
{
    internal int OutputLast => checked(OutputFirst + Last - First);
}
internal sealed record XlsxSheetGeometry(XlsxSheetSource Source, XlsxRectangle OutputArea, int RowOffset, int ColumnOffset)
{
    internal bool Expanded { get; init; }
    internal ImmutableArray<XlsxAxisSpan> RowSpans { get; init; } = [];
    internal ImmutableArray<XlsxAxisSpan> ColumnSpans { get; init; } = [];
    internal ImmutableArray<XlsxContentOwner> ContentOwners { get; init; } = [];
    internal ImmutableArray<XlsxDisplacedSource> DisplacedSources { get; init; } = [];
}
internal sealed record XlsxContentOwner(SourceCellId Source, VerticalInstancePath? Vertical,
    HorizontalInstancePath? Horizontal, XlsxAddress Output);
// A destination occupied by another retained source cannot remain a referent for this physical source.
internal sealed record XlsxDisplacedSource(SourceCellId Source, VerticalInstancePath? Vertical,
    HorizontalInstancePath? Horizontal, XlsxAddress Output);
internal sealed record XlsxCellPlacement(CellInstanceId Id, XlsxAddress Output, XlsxCellSource Source, XlsxPreparedValue? Value)
{
    internal string? Formula { get; init; }
    internal bool Suppressed { get; init; }
}
internal sealed record XlsxRowPlan(XElement Properties, int OutputRow, ImmutableArray<XlsxCellPlacement> Cells);
internal sealed record XlsxSheetOutput(XlsxSheetGeometry Geometry, XElement Settings, ImmutableArray<XlsxRowPlan> Rows);
internal sealed record XlsxWorkbookPlan(ImmutableArray<XlsxSheetOutput> Sheets, ImmutableArray<XElement> Names);
