namespace MonoReport.Xlsx.Planning;

internal sealed record XlsxPreparedValue(string? Type, string? Value)
{
    internal XlsxRichString? RichString { get; init; }
}