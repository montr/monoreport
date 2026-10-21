using System.Collections.Immutable;
using System.Globalization;
using System.Text;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

// A bounded operand scanner, not an Excel grammar or evaluator. Only reference spans are rewritten.
internal sealed class XlsxFormulaReferences
{
    internal string Copy(XlsxParsedFormula formula, int rowDelta, int columnDelta, string location) => Rewrite(formula, reference =>
    {
        XlsxFormulaPoint MovePoint(XlsxFormulaPoint point)
        {
            var row = (long)point.Address.Row + (point.RowAbsolute ? 0 : rowDelta);
            var column = (long)point.Address.Column + (point.ColumnAbsolute ? 0 : columnDelta);
            if (row is < 1 or > 1048576 || column is < 1 or > 16384) throw Unsafe(location, "Re-author the copied reference inside Excel worksheet dimensions.");
            return point with { Address = new((int)row, (int)column) };
        }
        return reference.Prefix + MovePoint(reference.First) + (reference.IsRange ? ":" + MovePoint(reference.Last) : string.Empty);
    });

    internal XlsxParsedFormula Read(string text, string location)
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
                if (i == text.Length || text[i++] != '!') throw Unsafe(location, "A quoted sheet reference requires ! and an explicit A1 operand.");
                prefix = text[start..i];
                if (prefix.IndexOfAny(['[', ']', ':']) >= 0) throw Unsafe(location, "Replace external/3-D references with explicit retained local A1 operands.");
            }
            else if (char.IsAsciiLetter(character) || character is '_' or '\\')
            {
                while (i < text.Length && (char.IsAsciiLetterOrDigit(text[i]) || text[i] is '_' or '.' or '\\')) i++;
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
                else throw Unsafe(location, "Replace unresolved names, structured or whole-axis references with explicit A1 cells/rectangles; name binding requires M3.");
            }
            if (prefix.Length > 0 || character == '$' || i == start && char.IsAsciiLetter(character))
            {
                if (!TryPoint(text, i, out var first, out var after)) throw Unsafe(location, "Use finite A1 cell/rectangle operands within Excel dimensions.");
                i = after; var last = first; var range = i < text.Length && text[i] == ':';
                if (range)
                {
                    if (!TryPoint(text, i + 1, out last, out after)) throw Unsafe(location, "Use a finite A1 range endpoint on the same sheet.");
                    i = after;
                    if (last!.Address.Row < first!.Address.Row || last.Address.Column < first.Address.Column)
                        throw Unsafe(location, "Use increasing rectangular A1 range endpoints.");
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

    private static TemplateException Unsafe(string location, string remedy) => new StructuralDiagnostic(DiagnosticCodes.XlsxUnsafeFormula,
        remedy, Location: location, Explanation: remedy).ToException();
}
