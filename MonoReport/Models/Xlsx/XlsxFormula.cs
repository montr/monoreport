using System.Collections.Immutable;

namespace MonoReport.Models.Xlsx;

internal sealed record XlsxFormulaPoint(XlsxAddress Address, bool RowAbsolute, bool ColumnAbsolute)
{
    public override string ToString()
    {
        var text = Address.ToString(); var at = 0;
        while (at < text.Length && char.IsAsciiLetter(text[at])) at++;
        return (ColumnAbsolute ? "$" : "") + text[..at] + (RowAbsolute ? "$" : "") + text[at..];
    }
}
internal sealed record XlsxFormulaReference(int Start, int Length, string Prefix,
    XlsxFormulaPoint First, XlsxFormulaPoint Last, bool IsRange)
{
    internal uint? TargetSheetId { get; init; }
    internal int? NameIndex { get; init; }
}
internal sealed record XlsxParsedFormula(string Text, ImmutableArray<XlsxFormulaReference> References);
internal sealed record XlsxDefinedName(int Index, string Name, uint? ScopeSheetId, XlsxParsedFormula Formula, string Location);
internal sealed record XlsxSharedFormulaGroup(uint Index, SourceCellId Master, XlsxRectangle Range,
    string Formula, ImmutableArray<SourceCellId> Members);

