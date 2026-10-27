using System.Collections.Immutable;
using System.Xml.Linq;

namespace MonoReport.Models.Xlsx;

internal readonly record struct XlsxAddress(int Row, int Column)
{
    public override string ToString()
    {
        var column = Column;
        var letters = string.Empty;
        while (column > 0) { column--; letters = (char)('A' + column % 26) + letters; column /= 26; }
        return letters + Row.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }
}

internal readonly record struct XlsxRectangle(XlsxAddress First, XlsxAddress Last)
{
    internal int Height => checked(Last.Row - First.Row + 1);
    internal int Width => checked(Last.Column - First.Column + 1);
    internal bool Contains(XlsxAddress cell) => cell.Row >= First.Row && cell.Row <= Last.Row &&
        cell.Column >= First.Column && cell.Column <= Last.Column;
    public override string ToString() => First + ":" + Last;
}

internal readonly record struct SourceCellId(uint SheetId, XlsxAddress Address);
internal sealed record VerticalInstancePath(int RegionId, ImmutableArray<int> LogicalRows)
{
    public bool Equals(VerticalInstancePath? other) => other is not null &&
        RegionId == other.RegionId && LogicalRows.AsSpan().SequenceEqual(other.LogicalRows.AsSpan());

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(RegionId);
        foreach (var index in LogicalRows.AsSpan()) hash.Add(index);
        return hash.ToHashCode();
    }
}
internal sealed record HorizontalInstancePath(int NodeId, int ItemIndex, HorizontalInstancePath? Parent)
{
    internal int Level { get; init; }
}
internal sealed record CellInstanceId(SourceCellId Source, VerticalInstancePath? Vertical, HorizontalInstancePath? Horizontal);

internal sealed record XlsxCellSource(SourceCellId Id, XElement Xml, string? Text, string? NativeType,
    bool HasFormula, bool RichText, XElement? SharedString = null)
{
    internal XlsxRichString? RichString { get; init; }
}
internal sealed record XlsxTextSpan(int Start, string Text, XElement? Properties);
internal sealed record XlsxRichString(string Text, ImmutableArray<XlsxTextSpan> Spans, ImmutableArray<XElement> Phonetics, XElement? PhoneticProperties);
internal sealed record XlsxTextEdit(int Start, int Length, string Text);
// RetentionArea is the reference domain: authored crop, or the full worksheet for identity retention.
// It is never a stored-content extent and must not be emitted as a synthetic dimension/PrintArea.
internal sealed record XlsxSheetSource(uint Id, string Name, string Part, XDocument Xml,
    XlsxRectangle? PrintArea, XlsxRectangle RetentionArea, ImmutableArray<XlsxCellSource> Cells, ImmutableArray<string> Features, int PrintAreaNameIndex);
internal sealed record XlsxSourceInventory(ImmutableArray<XlsxSheetSource> Sheets, bool Date1904,
    ImmutableArray<XElement> Names, ImmutableArray<(string Name, string Description)> WorkbookFeatures, XDocument? Styles)
{
    internal ImmutableArray<int> DisposableWorkbookNodes { get; init; } = [];
    internal string WorkbookPart { get; init; } = "/xl/workbook.xml";
}

internal sealed record XlsxCellPlan(XlsxCellSource Source, ImmutableArray<TemplateToken> Tokens, string Location)
{
    internal XlsxParsedFormula? Formula { get; init; }
    internal uint? SharedFormulaIndex { get; init; }
    internal bool NativeField => Tokens is [{ Kind: TemplateTokenKind.Expression, Start: 0, Expression.Formatter: null } token] &&
        token.Length == Source.Text?.Length;
}
internal enum XlsxControlKind { Vertical, Header, Geometry }
internal sealed record XlsxControlMarker(XlsxCellSource Cell, XlsxControlKind Kind, int Level,
    TemplateExpression? Collection, string? CollectionPath, string Location);
internal sealed record XlsxSheetTemplate(XlsxSheetSource Source, ImmutableArray<XlsxCellPlan> Cells,
    ImmutableArray<XlsxControlMarker> Controls)
{
    internal XlsxHorizontalNode? Horizontal { get; init; }
    internal ImmutableArray<XlsxVerticalRegion> Regions { get; init; } = [];
}
internal sealed record XlsxHorizontalNode(int Id, int Level, int FirstColumn, int LastColumn,
    ImmutableArray<XlsxHorizontalNode> Children);
internal sealed record XlsxVerticalBlock(int FirstRow, int LastRow);
internal sealed record XlsxVerticalNode(int Level, XlsxVerticalBlock? Header, XlsxVerticalBlock? Body,
    XlsxVerticalBlock? Footer, ImmutableArray<XlsxVerticalNode> Children);
internal sealed record XlsxVerticalRegion(int Id, int FirstRow, int LastRow, string Collection, XlsxVerticalNode Root)
{
    internal int Level => Root.Level;
}
internal sealed record XlsxTemplatePlan(XlsxSourceInventory Source, ImmutableArray<XlsxSheetTemplate> Sheets,
    ImmutableArray<PlannedDependency> Dependencies, IReadOnlyList<TemplateRequirement> Requirements,
    IReadOnlyList<TemplateOccurrence> Occurrences)
{
    internal ImmutableArray<XlsxDefinedName> DefinedNames { get; init; } = [];
    internal ImmutableArray<XlsxSharedFormulaGroup> SharedFormulas { get; init; } = [];
}
internal sealed record XlsxPreparedValue(string? Type, string? Value)
{
    internal XlsxRichString? RichString { get; init; }
}
