using System.Collections.Immutable;
using System.Globalization;
using System.Text;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

// A bounded operand scanner, not an Excel grammar or evaluator. Only reference spans are rewritten.
internal sealed partial class XlsxFormulaReferences
{
    internal XlsxParsedFormula ResolveSheets(XlsxParsedFormula formula, uint sourceSheetId,
        IReadOnlyDictionary<string, uint> sheetIds, string location) => formula with
    {
        References = formula.References.Select(reference =>
        {
            if (reference.NameIndex is not null) return reference;
            var target = sourceSheetId;
            if (reference.Prefix.Length > 0)
            {
                var name = reference.Prefix[..^1];
                if (name[0] == '\'') name = name[1..^1].Replace("''", "'", StringComparison.Ordinal);
                if (!sheetIds.TryGetValue(name, out target))
                    throw Unsafe(location, "Reference " + formula.Text.Substring(reference.Start, reference.Length) +
                        " names an unresolved worksheet '" + name + "'. Correct the worksheet name or use an existing same-workbook sheet and a retained A1 operand.");
            }
            return reference with { TargetSheetId = target };
        }).ToImmutableArray()
    };

    internal string Copy(XlsxParsedFormula formula, int rowDelta, int columnDelta, string location) => Rewrite(formula, reference =>
    {
        if (reference.NameIndex is not null) return formula.Text.Substring(reference.Start, reference.Length);
        XlsxFormulaPoint MovePoint(XlsxFormulaPoint point)
        {
            var row = (long)point.Address.Row + (point.RowAbsolute ? 0 : rowDelta);
            var column = (long)point.Address.Column + (point.ColumnAbsolute ? 0 : columnDelta);
            if (row is < 1 or > 1048576 || column is < 1 or > 16384) throw Unsafe(location, "Re-author the copied reference inside Excel worksheet dimensions.");
            return point with { Address = new((int)row, (int)column) };
        }
        return reference.Prefix + MovePoint(reference.First) + (reference.IsRange ? ":" + MovePoint(reference.Last) : string.Empty);
    });

    internal XlsxParsedFormula Read(string text, string location, Func<string, string, XlsxFormulaReference>? resolveName = null)
    {
        var references = ImmutableArray.CreateBuilder<XlsxFormulaReference>();
        var depth = 0;
        for (var i = 0; i < text.Length;)
        {
            var character = text[i];
            if (char.IsWhiteSpace(character)) { i++; continue; }
            if (character == '"') { Quoted(text, ref i, '"', location); continue; }
            var start = i; var prefix = string.Empty;
            if (character == '\'')
            {
                Quoted(text, ref i, '\'', location);
                RejectSheetRange(text, start, i, location);
                if (i == text.Length || text[i++] != '!') throw Unsafe(location, "A quoted sheet reference requires ! and an explicit A1 or simple name operand.");
                prefix = text[start..i];
                if (prefix.IndexOfAny(['[', ']']) >= 0) throw UnsupportedSheet(text, start, location, true);
                if (prefix.Contains(':')) throw UnsupportedSheet(text, start, location, false);
            }
            else if (char.IsLetter(character) || character is '_' or '\\')
            {
                while (i < text.Length && (char.IsLetterOrDigit(text[i]) || text[i] is '_' or '.' or '\\')) i++;
                RejectSheetRange(text, start, i, location);
                var word = text[start..i]; var next = i;
                while (next < text.Length && char.IsWhiteSpace(text[next])) next++;
                if (next < text.Length && text[next] == '(')
                {
                    if (word.Split('.')[^1].Equals("INDIRECT", StringComparison.OrdinalIgnoreCase) ||
                        word.Split('.')[^1].Equals("OFFSET", StringComparison.OrdinalIgnoreCase))
                        throw Unsafe(location, "Replace constructed INDIRECT/OFFSET references with explicit retained A1 operands.");
                    continue;
                }
                if (i < text.Length && text[i] == '!') prefix = text[start..++i];
                else if (TryPoint(text, start, out _, out var end) && end >= i) i = start;
                else if (word.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || word.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) continue;
                else if (resolveName is not null) i = start;
                else throw Unsafe(location, "Replace unresolved names, structured or whole-axis references with explicit A1 cells/rectangles.");
            }
            else if (character == '[') throw UnsupportedSheet(text, start, location, true);
            if (prefix.Length > 0 || character == '$' || i == start && IdentifierStart(character))
            {
                if (!TryPoint(text, i, out var first, out var after) || after < text.Length && IdentifierPart(text[after]))
                {
                    var nameStart = i;
                    if (resolveName is null || i == text.Length || !IdentifierStart(text[i]))
                        throw Unsafe(location, "Use finite A1 cell/rectangle operands within Excel dimensions.");
                    while (i < text.Length && IdentifierPart(text[i])) i++;
                    if (i < text.Length && text[i] is '(' or '!' or '[')
                        throw Unsafe(location, "Use a simple name operand; structured references and sheet-qualified function calls are unsupported.");
                    references.Add(resolveName(prefix, text[nameStart..i]) with { Start = start, Length = i - start, Prefix = prefix });
                    continue;
                }
                i = after; var last = first; var range = i < text.Length && text[i] == ':';
                if (range)
                {
                    if (!TryPoint(text, i + 1, out last, out after)) throw Unsafe(location, "Use a finite A1 range endpoint on the same sheet.");
                    i = after;
                }
                if (i < text.Length && (char.IsAsciiLetterOrDigit(text[i]) || text[i] is '_' or '.' or '$' or '!'))
                    throw Unsafe(location, "Separate A1 operands from identifiers; the reference boundary is ambiguous.");
                references.Add(new(start, i - start, prefix, first!, last!, range)); continue;
            }
            if (char.IsAsciiDigit(character) || character == '.' && i + 1 < text.Length && char.IsAsciiDigit(text[i + 1]))
            {
                while (i < text.Length && char.IsAsciiDigit(text[i])) i++;
                if (i < text.Length && text[i] == '.') { i++; while (i < text.Length && char.IsAsciiDigit(text[i])) i++; }
                if (i < text.Length && text[i] is 'e' or 'E')
                {
                    i++; if (i < text.Length && text[i] is '+' or '-') i++;
                    var exponent = i; while (i < text.Length && char.IsAsciiDigit(text[i])) i++;
                    if (i == exponent) throw Unsafe(location, "Repair the malformed numeric exponent.");
                }
                if (i < text.Length && (text[i] == '.' || char.IsAsciiLetter(text[i]))) throw Unsafe(location, "Repair the malformed numeric token.");
                continue;
            }
            if (character == '(') depth++;
            else if (character == ')' && --depth < 0) throw Unsafe(location, "Balance formula parentheses.");
            else if (!"+-*/^&%=<>,()".Contains(character)) throw Unsafe(location, "Replace unsupported formula tokens (including arrays, spills, implicit intersection or external references) with ordinary functions over explicit A1 operands.");
            i++;
        }
        if (depth != 0) throw Unsafe(location, "Balance formula parentheses.");
        if (text.Length == 0) throw Unsafe(location, "Supply a nonempty scalar formula; shared followers require M3 normalization.");
        return new(text, references.ToImmutable());
    }

    internal string Rewrite(XlsxParsedFormula formula, Func<XlsxFormulaReference, string> transform)
    {
        var result = new StringBuilder(formula.Text);
        foreach (var reference in formula.References.Reverse()) result.Remove(reference.Start, reference.Length).Insert(reference.Start, transform(reference));
        return result.ToString();
    }

    private static bool TryPoint(string text, int at, out XlsxFormulaPoint? point, out int end)
    {
        point = null; end = at; var i = at;
        var columnAbsolute = i < text.Length && text[i] == '$'; if (columnAbsolute) i++;
        var letters = i; var column = 0;
        while (i < text.Length && char.ToUpperInvariant(text[i]) is >= 'A' and <= 'Z')
        {
            if (i - letters >= 3) return false;
            column = column * 26 + char.ToUpperInvariant(text[i++]) - 'A' + 1;
        }
        if (i == letters) return false;
        var rowAbsolute = i < text.Length && text[i] == '$'; if (rowAbsolute) i++;
        var digits = i; while (i < text.Length && char.IsAsciiDigit(text[i])) i++;
        if (i == digits || column > 16384 || !int.TryParse(text[digits..i], NumberStyles.None, CultureInfo.InvariantCulture, out var row) || row is < 1 or > 1048576) return false;
        point = new(new(row, column), rowAbsolute, columnAbsolute); end = i; return true;
    }

    private static bool IdentifierStart(char character) => char.IsLetter(character) || character is '_' or '\\';
    private static bool IdentifierPart(char character) => char.IsLetterOrDigit(character) || character is '_' or '.' or '\\';

    private static void Quoted(string text, ref int at, char quote, string location)
    {
        at++;
        while (at < text.Length)
        {
            if (text[at++] != quote) continue;
            if (at < text.Length && text[at] == quote) { at++; continue; }
            return;
        }
        throw Unsafe(location, "Close the formula string/sheet quotation.");
    }

    private static void RejectSheetRange(string text, int start, int at, string location)
    {
        if (at == text.Length || text[at] != ':') return;
        var end = at + 1;
        if (end < text.Length && text[end] == '\'') Quoted(text, ref end, '\'', location);
        else while (end < text.Length && (char.IsLetterOrDigit(text[end]) || text[end] is '_' or '.' or '\\')) end++;
        if (end < text.Length && text[end] == '!') throw UnsupportedSheet(text, start, location, false);
    }

    private static TemplateException UnsupportedSheet(string text, int start, string location, bool external)
    {
        var end = start;
        var quoted = false;
        while (end < text.Length)
        {
            var character = text[end];
            if (character == '\'') quoted = !quoted;
            if (!quoted && (char.IsWhiteSpace(character) || "+-*/^&%=<>,()".Contains(character))) break;
            end++;
        }
        return Unsafe(location, "Reference " + text[start..end] + (external
            ? " uses an external workbook. Move the required values into this workbook and use an explicit same-workbook A1 cell or range."
            : " uses an unsupported 3-D worksheet range. Replace it with explicit same-workbook A1 operands for each required worksheet."));
    }

    private static TemplateException Unsafe(string location, string remedy) => new StructuralDiagnostic(DiagnosticCodes.XlsxUnsafeFormula,
        remedy, Location: location, Explanation: remedy).ToException();
}
