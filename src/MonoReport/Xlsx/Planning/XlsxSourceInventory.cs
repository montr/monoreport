using System.Collections.Immutable;
using System.Xml.Linq;

namespace MonoReport.Xlsx.Planning;

// RetentionArea is the reference domain: authored crop, or the full worksheet for identity retention.
// It is never a stored-content extent and must not be emitted as a synthetic dimension/PrintArea.
internal sealed record XlsxSheetSource(
    uint Id,
    string Name,
    string Part,
    XDocument Xml,
    XlsxRectangle? PrintArea,
    XlsxRectangle RetentionArea,
    ImmutableArray<XlsxCellSource> Cells,
    ImmutableArray<string> Features,
    int PrintAreaNameIndex);

internal sealed record XlsxSourceInventory(
    ImmutableArray<XlsxSheetSource> Sheets,
    bool Date1904,
    ImmutableArray<XElement> Names,
    ImmutableArray<(string Name, string Description)> WorkbookFeatures,
    XDocument? Styles)
{
    internal ImmutableArray<int> DisposableWorkbookNodes { get; init; } = [];
    internal string WorkbookPart { get; init; } = "/xl/workbook.xml";
}