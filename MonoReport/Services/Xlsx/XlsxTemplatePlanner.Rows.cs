using System.Collections.Immutable;
using System.Text.Json;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxTemplatePlanner
{
    internal DataScope RowScope(JsonElement root, XlsxLogicalRow row, string collection, int? tupleIndex)
    {
        var tuple = tupleIndex ?? 0;
        var values = new List<JsonElement> { row.Records[tuple] };
        var paths = new List<string> { collection + "[" + (row.FirstIndex + tuple) + "]" };
        for (var parent = row.Parent; parent is not null; parent = parent.Parent)
        {
            values.Add(parent.Records[0]);
            paths.Add(collection + "[" + parent.FirstIndex + "]");
        }
        return new(root, values, paths);
    }

    internal DataScope HeaderScope(JsonElement root, ImmutableArray<XlsxBoundCollection> headers, IReadOnlyList<int> indices)
    {
        var values = new List<JsonElement>(); var paths = new List<string>();
        for (var i = indices.Count - 1; i >= 0; i--)
        {
            values.Add(headers[i].Items[indices[i]]);
            paths.Add(headers[i].Path + "[" + indices[i] + "]");
        }
        values.Add(root); paths.Add("$");
        return new(root, values, paths);
    }

    internal XlsxBoundSheet BindSheet(XlsxSheetTemplate sheet, JsonElement root, XlsxOperationContext context)
    {
        if (root.ValueKind != JsonValueKind.Object) throw InvalidRows("$", "Supply one JSON object as the XLSX report data.");
        var headers = sheet.Controls.Where(marker => marker.Kind == XlsxControlKind.Header)
            .GroupBy(marker => marker.Level).OrderBy(group => group.Key)
            .Select(group => new XlsxBoundCollection(group.Key, group.First().CollectionPath!, Collection(group.First(), root, context)))
            .ToImmutableArray();
        var rows = ImmutableDictionary.CreateBuilder<int, ImmutableArray<XlsxLogicalRow>>();
        foreach (var region in sheet.Regions)
        {
            var marker = sheet.Controls.Single(marker => marker.Kind == XlsxControlKind.Vertical && marker.Cell.Id.Address.Row == region.FirstRow);
            var items = Collection(marker, root, context);
            var intersects = sheet.Horizontal is { } tree && sheet.Cells.Any(cell =>
                cell.Source.Id.Address.Row >= region.FirstRow && cell.Source.Id.Address.Row <= region.LastRow &&
                cell.Source.Id.Address.Column >= tree.FirstColumn && cell.Source.Id.Address.Column <= tree.LastColumn);
            try
            {
                var grouped = GroupRecords(items, region.Collection, intersects);
                ValidateArity(grouped, intersects ? RecordArity(headers.Select(header => header.Items.Length)) : 1, region.Collection);
                rows.Add(region.Id, BindHierarchy(grouped));
            }
            catch (TemplateException error)
            {
                context.Report(new StructuralDiagnostic(error.Code, error.Message, marker.Cell.Text, Location: marker.Location, Explanation: error.Message).ToException(), error.Location);
                context.ThrowIfErrors();
            }
        }
        return new(root, headers, rows.ToImmutable());
    }

    private ImmutableArray<JsonElement> Collection(XlsxControlMarker marker, JsonElement root, XlsxOperationContext context)
    {
        void Reject(string remedy)
        {
            context.Report(new StructuralDiagnostic(DiagnosticCodes.GenerationInvalidData, remedy, marker.Cell.Text,
                Location: marker.Location, Explanation: remedy).ToException(), marker.CollectionPath);
            context.ThrowIfErrors();
        }
        var resolution = _resolver.Resolve(marker.Collection!.Path!, new DataScope(root, [root], ["$"]));
        if (resolution.State == ResolutionState.Null) return [];
        if (resolution.State == ResolutionState.Missing)
        {
            if (context.Options.MissingValues == MissingValueBehavior.Error) Reject("Supply the missing collection " + marker.CollectionPath + ".");
            if (context.Options.MissingValues == MissingValueBehavior.Warning)
                context.Warnings.Add(new("A collection is missing; no repetitions were generated.", marker.Cell.Text!, marker.Cell.Text!, marker.Location, marker.Location.Split(',', 2)[0], marker.CollectionPath) { Code = DiagnosticCodes.XlsxMissingValue });
            return [];
        }
        if (resolution.Value.ValueKind != JsonValueKind.Array) Reject("Supply an array for " + marker.CollectionPath + ".");
        return resolution.Value.EnumerateArray().ToImmutableArray();
    }

    internal int RecordArity(IEnumerable<int> counts)
    {
        var sizes = counts.ToArray();
        if (sizes.Any(count => count < 0)) throw InvalidRows("$", "Use nonnegative horizontal cardinalities.");
        if (sizes.Contains(0)) return 1;
        long product = 1;
        try
        {
            foreach (var count in sizes)
            {
                if (count < 0) throw new OverflowException();
                product = checked(product * count);
            }
            return checked((int)Math.Max(1, product));
        }
        catch (OverflowException) { throw InvalidRows("$", "Reduce horizontal collection cardinalities; the Cartesian product exceeds supported indexing."); }
    }

    internal void ValidateArity(ImmutableArray<XlsxLogicalRow> rows, int arity, string path)
    {
        foreach (var row in rows)
            if (row.Records.Length != arity)
                throw InvalidRows(path + "[" + row.FirstIndex + "]", "Supply exactly " + arity +
                    " records per contiguous $path group in horizontal Cartesian order; an empty product requires one descriptor.");
    }

    internal ImmutableArray<XlsxLogicalRow> GroupRows(JsonElement records, string dataPath, bool requirePath)
    {
        if (records.ValueKind != JsonValueKind.Array) throw InvalidRows(dataPath, "Supply an array of logical-row records.");
        return GroupRecords(records.EnumerateArray(), dataPath, requirePath);
    }

    private static ImmutableArray<XlsxLogicalRow> GroupRecords(IEnumerable<JsonElement> records, string dataPath, bool requirePath)
    {
        var rows = ImmutableArray.CreateBuilder<XlsxLogicalRow>();
        var closed = new HashSet<string>(StringComparer.Ordinal);
        var group = ImmutableArray.CreateBuilder<JsonElement>();
        string? key = null;
        var level = 0;
        var first = 0;
        var index = 0;
        foreach (var record in records)
        {
            var path = dataPath + "[" + index + "]";
            if (record.ValueKind != JsonValueKind.Object || !record.TryGetProperty("$level", out var value) ||
                value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out var currentLevel))
                throw InvalidRows(path, "Supply an integer $level on every logical-row record.");
            string? currentKey = null;
            if (record.TryGetProperty("$path", out var property))
            {
                if (property.ValueKind != JsonValueKind.String) throw InvalidRows(path, "Supply a string $path grouping key.");
                try { currentKey = property.GetString(); }
                catch (InvalidOperationException) { throw InvalidRows(path, "Supply valid Unicode in the $path grouping key."); }
            }
            if (requirePath && currentKey is null) throw InvalidRows(path, "Supply $path to group the horizontal intersection records.");
            if (group.Count > 0 && (currentKey is null || currentKey != key))
            {
                rows.Add(new(key, level, first, group.ToImmutable()));
                if (key is not null) closed.Add(key);
                group.Clear();
            }
            if (group.Count == 0)
            {
                if (currentKey is not null && closed.Contains(currentKey))
                    throw InvalidRows(path, "Keep each $path group contiguous; a closed path must not reappear. Do not sort records in the engine.");
                first = index; key = currentKey; level = currentLevel;
            }
            else if (currentLevel != level) throw InvalidRows(path, "Use the same integer $level throughout a $path group.");
            group.Add(record); index++;
        }
        if (group.Count > 0) rows.Add(new(key, level, first, group.ToImmutable()));
        return rows.ToImmutable();
    }

    private static TemplateException InvalidRows(string path, string remedy) =>
        new StructuralDiagnostic(DiagnosticCodes.GenerationInvalidData, remedy, Location: path, Explanation: remedy).ToException();
}
