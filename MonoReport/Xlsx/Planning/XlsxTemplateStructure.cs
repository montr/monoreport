using System.Collections.Immutable;

namespace MonoReport.Xlsx.Planning;

internal sealed record XlsxHorizontalNode(
    int Id,
    int Level,
    int FirstColumn,
    int LastColumn,
    ImmutableArray<XlsxHorizontalNode> Children);

internal sealed record XlsxVerticalBlock(int FirstRow, int LastRow);

internal sealed record XlsxVerticalNode(
    int Level,
    XlsxVerticalBlock? Header,
    XlsxVerticalBlock? Body,
    XlsxVerticalBlock? Footer,
    ImmutableArray<XlsxVerticalNode> Children);

internal sealed record XlsxVerticalRegion(int Id, int FirstRow, int LastRow, string Collection, XlsxVerticalNode Root)
{
    internal int Level => Root.Level;
}