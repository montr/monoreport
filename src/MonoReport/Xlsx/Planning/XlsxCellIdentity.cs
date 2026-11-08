using System;
using System.Collections.Immutable;

namespace MonoReport.Xlsx.Planning;

internal readonly record struct SourceCellId(uint SheetId, XlsxAddress Address);

internal sealed record VerticalInstancePath(int RegionId, ImmutableArray<int> LogicalRows)
{
    public bool Equals(VerticalInstancePath? other)
    {
        return other is not null &&
               RegionId == other.RegionId && LogicalRows.AsSpan().SequenceEqual(other.LogicalRows.AsSpan());
    }

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

internal sealed record CellInstanceId(
    SourceCellId Source,
    VerticalInstancePath? Vertical,
    HorizontalInstancePath? Horizontal);