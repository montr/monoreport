using System.Collections.Immutable;
using System.Xml.Linq;

namespace MonoReport.Models.Xlsx;

internal enum XlsxRangeImageKind { Cell, Rectangle, Empty, Disjoint, Ambiguous }
internal sealed record XlsxRangeImage(XlsxRangeImageKind Kind, XlsxRectangle? Rectangle, bool Complete);
internal sealed record XlsxSheetGeometry(XlsxSheetSource Source, XlsxRectangle OutputArea, int RowOffset, int ColumnOffset);
internal sealed record XlsxCellPlacement(CellInstanceId Id, XlsxAddress Output, XlsxCellSource Source, XlsxPreparedValue? Value);
internal sealed record XlsxRowPlan(XElement Properties, int OutputRow, ImmutableArray<XlsxCellPlacement> Cells);
internal sealed record XlsxSheetOutput(XlsxSheetGeometry Geometry, XElement Settings, ImmutableArray<XlsxRowPlan> Rows);
internal sealed record XlsxWorkbookPlan(ImmutableArray<XlsxSheetOutput> Sheets, ImmutableArray<XElement> Names);
