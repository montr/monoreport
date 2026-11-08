using System.Collections.Immutable;
using MonoReport.Core.Templates;

namespace MonoReport.Xlsx.Planning;

internal sealed record XlsxCellPlan(XlsxCellSource Source, ImmutableArray<TemplateToken> Tokens, string Location)
{
    internal XlsxParsedFormula? Formula { get; init; }
    internal uint? SharedFormulaIndex { get; init; }

    internal bool NativeField =>
        Tokens is [{Kind: TemplateTokenKind.Expression, Start: 0, Expression.Formatter: null} token] &&
        token.Length == Source.Text?.Length;
}

internal enum XlsxControlKind
{
    Vertical,
    Header,
    Geometry
}

internal sealed record XlsxControlMarker(
    XlsxCellSource Cell,
    XlsxControlKind Kind,
    int Level,
    TemplateExpression? Collection,
    string? CollectionPath,
    string Location);

internal sealed record XlsxSheetTemplate(
    XlsxSheetSource Source,
    ImmutableArray<XlsxCellPlan> Cells,
    ImmutableArray<XlsxControlMarker> Controls)
{
    internal XlsxHorizontalNode? Horizontal { get; init; }
    internal ImmutableArray<XlsxVerticalRegion> Regions { get; init; } = [];
}