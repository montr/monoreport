using System.Collections.Immutable;
using System.Globalization;
using System.Text.RegularExpressions;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxTemplatePlanner
{
    private ImmutableArray<XlsxControlMarker> ReadControls(XlsxSheetSource sheet, XlsxOperationContext context)
    {
        var controls = ImmutableArray.CreateBuilder<XlsxControlMarker>();
        foreach (var cell in sheet.Cells)
        {
            if (context.LimitReached) break;
            if (cell.Text is not { } text || !ControlPrefix().IsMatch(text)) continue;
            var location = Location(sheet, cell.Id.Address);
            var match = Control().Match(text);
            if (!match.Success || !int.TryParse(match.Groups["level"].Value, NumberStyles.AllowLeadingSign,
                    CultureInfo.InvariantCulture, out var level))
            {
                context.Report(ControlError(location, "Use an integer level and a valid collection path in a complete control cell."));
                continue;
            }
            var vertical = match.Groups["kind"].Value.StartsWith("l", StringComparison.OrdinalIgnoreCase);
            var hasCollection = match.Groups["collection"].Success;
            TemplateExpression? collection = null;
            string? canonical = null;
            if (hasCollection)
            {
                var parsed = _language.Parse("{For:" + match.Groups["collection"].Value + "}", location);
                if (parsed.Diagnostics.Count != 0 || parsed.Tokens is not [{ Expression.Path: { Origin: TemplatePathOrigin.Current, ParentDepth: 0 } } token])
                {
                    context.Report(ControlError(location, "Use a root collection property path after the control level."));
                    continue;
                }
                canonical = "$." + string.Join('.', token.Expression!.Path!.Segments);
                collection = token.Expression with { Source = text };
            }
            var kind = vertical ? XlsxControlKind.Vertical : hasCollection ? XlsxControlKind.Header : XlsxControlKind.Geometry;
            var address = cell.Id.Address;
            var validPlacement = kind == XlsxControlKind.Geometry
                ? address.Row < sheet.PrintArea.First.Row && address.Column >= sheet.PrintArea.First.Column && address.Column <= sheet.PrintArea.Last.Column
                : hasCollection && address.Column < sheet.PrintArea.First.Column && address.Row >= sheet.PrintArea.First.Row && address.Row <= sheet.PrintArea.Last.Row;
            if (!validPlacement)
            {
                context.Report(ControlError(location, "Place collection controls in one column left of PrintArea and geometry controls above its report columns."));
                continue;
            }
            controls.Add(new(cell, kind, level, collection, canonical, location));
        }
        var rowControls = controls.Where(marker => marker.Kind != XlsxControlKind.Geometry).ToArray();
        if (rowControls.Select(marker => marker.Cell.Id.Address.Column).Distinct().Skip(1).Any())
            context.Report(ControlError(rowControls[0].Location, "Use one unambiguous control column left of PrintArea."));
        foreach (var level in controls.Where(marker => marker.Kind == XlsxControlKind.Header).GroupBy(marker => marker.Level))
            if (level.Select(marker => marker.CollectionPath).Distinct(StringComparer.Ordinal).Skip(1).Any())
                context.Report(ControlError(level.First().Location, "Associate each horizontal header level with one case-sensitive root collection."));
        context.ThrowIfErrors();
        return controls.ToImmutable();
    }

    private static Dictionary<int, ImmutableArray<string>> CellScopes(ImmutableArray<XlsxControlMarker> controls)
    {
        var result = new Dictionary<int, ImmutableArray<string>>();
        var headers = new SortedDictionary<int, XlsxControlMarker>();
        foreach (var marker in controls.Where(marker => marker.Kind == XlsxControlKind.Header).OrderBy(marker => marker.Cell.Id.Address.Row))
        {
            var row = marker.Cell.Id.Address.Row;
            foreach (var key in headers.Keys.Where(key => key >= marker.Level).ToArray()) headers.Remove(key);
            headers.Add(marker.Level, marker);
            result[row] = headers.Values.Reverse().Select(item => item.CollectionPath!).ToImmutableArray();
        }
        return result;
    }

    private static long Order(XlsxAddress address) => ((long)address.Row << 40) + ((long)address.Column << 24);
    private static string Location(XlsxSheetSource sheet, XlsxAddress address) => sheet.Part + ", sheet '" + sheet.Name + "', " + address;
    private static TemplateException ControlError(string location, string explanation) =>
        new StructuralDiagnostic(DiagnosticCodes.TemplateInvalidStructure, explanation, Location: location, Explanation: explanation).ToException();

    [GeneratedRegex(@"^\s*\$(?:level|l|header|h)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex ControlPrefix();
    [GeneratedRegex(@"^\$(?<kind>level|l|header|h)\((?<level>[+-]?[0-9]+)\)(?:\.(?<collection>[^\s]+))?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex Control();
}
