using System.Collections.Immutable;
using System.Text.Json;

namespace MonoReport.Models.Xlsx;

internal sealed record XlsxLogicalRow(string? Key, int Level, int FirstIndex, ImmutableArray<JsonElement> Records);
internal sealed record XlsxBoundCollection(int Level, string Path, ImmutableArray<JsonElement> Items);
internal sealed record XlsxBoundSheet(JsonElement Root, ImmutableArray<XlsxBoundCollection> Headers,
    ImmutableDictionary<int, ImmutableArray<XlsxLogicalRow>> Rows);
