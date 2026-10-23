using System.Collections.Immutable;
using System.Globalization;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxFormulaReferences
{
    internal ImmutableArray<XlsxDefinedName> ReadNames(XlsxSourceInventory source, XlsxOperationContext context)
    {
        var result = ImmutableArray.CreateBuilder<XlsxDefinedName>();
        var printAreas = source.Sheets.Select(sheet => sheet.PrintAreaNameIndex).ToHashSet();
        var sheets = source.Sheets.ToDictionary(sheet => sheet.Name, sheet => sheet.Id, StringComparer.OrdinalIgnoreCase);
        var keys = new HashSet<(uint, string)>();
        for (var index = 0; index < source.Names.Length && !context.LimitReached; index++)
        {
            if (printAreas.Contains(index)) continue;
            var xml = source.Names[index];
            var name = (string?)xml.Attribute("name") ?? string.Empty;
            var location = "/xl/workbook.xml, defined name '" + name + "'";
            try
            {
                if (name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase))
                {
                    var remedy = "Unsupported built-in defined name '" + name + "'. Remove print titles or other unsupported built-in names; only PrintArea is currently supported.";
                    throw new StructuralDiagnostic(DiagnosticCodes.XlsxUnsupportedFeature,
                        remedy, Location: location, Explanation: remedy).ToException();
                }
                uint? scope = null;
                if (xml.Attribute("localSheetId") is { } local)
                {
                    if (!int.TryParse(local.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var sheetIndex) || sheetIndex < 0 || sheetIndex >= source.Sheets.Length)
                        throw Unsafe(location, "Correct localSheetId to the worksheet's zero-based workbook position for defined name '" + name + "'.");
                    scope = source.Sheets[sheetIndex].Id;
                    location += ", sheet '" + source.Sheets[sheetIndex].Name + "'";
                }
                if (name.Length is < 1 or > 255 || !IdentifierStart(name[0]) || !name.All(IdentifierPart) ||
                    name.Equals("R", StringComparison.OrdinalIgnoreCase) || name.Equals("C", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || name.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ||
                    System.Text.RegularExpressions.Regex.IsMatch(name, "^[Rr][0-9]*[Cc][0-9]*$") ||
                    TryPoint(name, 0, out _, out var after) && after == name.Length)
                    throw Unsafe(location, "Use a valid simple identifier for defined name '" + name + "', distinct from A1 cells and reserved literals.");
                if (!keys.Add((scope ?? 0, name.ToUpperInvariant())))
                    throw Unsafe(location, "Defined name '" + name + "' is duplicated in the same scope. Keep one definition per workbook or worksheet scope; names are case-insensitive.");
                if (xml.Name.LocalName != "definedName" || xml.Attributes().Any(attribute =>
                    attribute.Name.NamespaceName.Length != 0 || attribute.Name.LocalName is not ("name" or "localSheetId" or "hidden" or "comment")) ||
                    xml.Attribute("hidden") is { Value: not ("0" or "1" or "true" or "false") })
                    throw Unsafe(location, "Remove unsupported function/macro metadata from defined name '" + name + "'; retain an ordinary cell/range name.");
                XlsxParsedFormula formula;
                try { formula = Read(xml.Value, location); }
                catch (TemplateException error)
                {
                    throw Unsafe(location, "Defined name '" + name + "' has unsupported definition " + xml.Value + ". " +
                        "Use one explicit same-workbook worksheet-qualified absolute A1 cell or rectangle; formula-valued, recursive, external and 3-D definitions are unsupported. " + error.Message);
                }
                if (formula.References is not [var reference] || reference.Start != 0 || reference.Length != formula.Text.Length ||
                    reference.Prefix.Length == 0 || !reference.First.RowAbsolute || !reference.First.ColumnAbsolute ||
                    !reference.Last.RowAbsolute || !reference.Last.ColumnAbsolute)
                    throw Unsafe(location, "Defined name '" + name + "' has context-dependent or unsupported definition " + xml.Value +
                        ". Use one explicit same-workbook worksheet-qualified absolute A1 cell or rectangle; relative/mixed, formula-valued and multi-area definitions are unsupported.");
                formula = ResolveSheets(formula, scope ?? source.Sheets[0].Id, sheets, location);
                result.Add(new(index, name, scope, formula, location));
            }
            catch (TemplateException error) { context.Report(error); }
        }
        context.ThrowIfErrors();
        return result.ToImmutable();
    }

    internal XlsxFormulaReference ResolveName(IReadOnlyDictionary<(uint, string), XlsxDefinedName> names,
        IReadOnlyDictionary<string, uint> sheets, uint sourceSheetId, string prefix, string name, string location)
    {
        var scope = sourceSheetId;
        if (prefix.Length > 0)
        {
            var sheet = prefix[..^1];
            if (sheet[0] == '\'') sheet = sheet[1..^1].Replace("''", "'", StringComparison.Ordinal);
            if (!sheets.TryGetValue(sheet, out scope))
                throw Unsafe(location, "Name operand " + prefix + name + " has an unresolved worksheet. Correct the worksheet name or use an existing same-workbook sheet and a retained A1 operand.");
        }
        var key = name.ToUpperInvariant();
        if (!names.TryGetValue((scope, key), out var found) && !names.TryGetValue((0, key), out found))
            throw Unsafe(location, "Name operand " + prefix + name + " is unresolved in its worksheet/workbook scope. Define it as one explicit absolute same-workbook A1 cell/rectangle or use a retained A1 operand directly.");
        return found.Formula.References[0] with { NameIndex = found.Index };
    }
}
