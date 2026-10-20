using System.Text;

// Research only: bounded reference-token scanner, not Excel syntax validation or evaluation.
sealed record CellPoint(int Row, int Col, bool RowAbsolute, bool ColAbsolute)
{
    public override string ToString() => (ColAbsolute ? "$" : "") + Letters(Col) + (RowAbsolute ? "$" : "") + Row;
    private static string Letters(int col)
    {
        var result = "";
        for (; col > 0; col = (col - 1) / 26) result = (char)('A' + (col - 1) % 26) + result;
        return result;
    }
}

sealed record ReferenceToken(int Start, int Length, string Prefix, CellPoint First, CellPoint Last, bool IsRange);

sealed class InternalReferences
{
    public IReadOnlyList<ReferenceToken> Read(string formula, ISet<string>? names = null)
    {
        var refs = new List<ReferenceToken>();
        var depth = 0;
        for (var i = 0; i < formula.Length;)
        {
            var ch = formula[i];
            if (char.IsWhiteSpace(ch)) { i++; continue; }
            if (ch == '"') { Quoted(formula, ref i, '"'); continue; }
            var start = i;
            var prefix = "";
            string? word = null;
            if (ch == '\'')
            {
                Quoted(formula, ref i, '\'');
                if (i == formula.Length || formula[i++] != '!') throw new FormatException("Quoted sheet requires !.");
                prefix = formula[start..i];
                if (prefix.IndexOfAny(['[', ']', ':']) >= 0) throw new NotSupportedException("External/3-D references require a separate capability.");
            }
            else if (char.IsLetter(ch) || ch is '_' or '\\')
            {
                while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] is '_' or '.' or '\\')) i++;
                word = formula[start..i];
                var next = i;
                while (next < formula.Length && char.IsWhiteSpace(formula[next])) next++;
                if (next < formula.Length && formula[next] == '(')
                {
                    if (word.Equals("INDIRECT", StringComparison.OrdinalIgnoreCase) || word.Equals("OFFSET", StringComparison.OrdinalIgnoreCase))
                        throw new NotSupportedException("Constructed references are unsafe under geometry changes.");
                    // Function names such as LOG10 resemble references. Leave all ordinary names opaque.
                    continue;
                }
                if (i < formula.Length && formula[i] == '!')
                {
                    prefix = formula[start..++i];
                    word = null;
                }
                else if (i < formula.Length && formula[i] == ':')
                {
                    if (!TryCell(formula, start, out _, out _)) throw new NotSupportedException("Whole-axis/3-D reference is outside the bounded capability.");
                    i = start;
                }
                else if (i < formula.Length && formula[i] == '[') throw new NotSupportedException("Structured reference is outside the bounded capability.");
                else if (TryCell(formula, start, out _, out var end) && end >= i) i = start;
                else
                {
                    if (word.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || word.Equals("FALSE", StringComparison.OrdinalIgnoreCase) || names?.Contains(word) == true)
                        continue;
                    throw new NotSupportedException("Unresolved name: " + word);
                }
            }
            if (prefix.Length != 0 || ch == '$' || i == start && char.IsLetter(ch))
            {
                if (!TryCell(formula, i, out var first, out var after)) throw new NotSupportedException("Only bounded A1 cell references are supported here.");
                i = after;
                var last = first;
                var isRange = i < formula.Length && formula[i] == ':';
                if (isRange)
                {
                    if (!TryCell(formula, i + 1, out last, out after)) throw new NotSupportedException("Range endpoint is not a local bounded A1 cell.");
                    i = after;
                }
                if (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] is '_' or '.' or '$')) throw new FormatException("Reference token boundary is ambiguous.");
                refs.Add(new(start, i - start, prefix, first!, last!, isRange));
                continue;
            }
            if (char.IsDigit(ch) || ch == '.' && i + 1 < formula.Length && char.IsDigit(formula[i + 1]))
            {
                while (i < formula.Length && (char.IsDigit(formula[i]) || formula[i] == '.')) i++;
                if (i < formula.Length && formula[i] is 'e' or 'E')
                {
                    i++;
                    if (i < formula.Length && formula[i] is '+' or '-') i++;
                    var exponent = i;
                    while (i < formula.Length && char.IsDigit(formula[i])) i++;
                    if (i == exponent) throw new FormatException("Malformed numeric exponent.");
                }
                continue;
            }
            if (ch == '(') depth++;
            else if (ch == ')' && --depth < 0) throw new FormatException("Unbalanced parentheses.");
            else if (!"+-*/^&%=<>,()".Contains(ch)) throw new NotSupportedException("Unsupported formula token: " + ch);
            i++;
        }
        if (depth != 0) throw new FormatException("Unbalanced parentheses.");
        return refs;
    }

    public string Rewrite(string formula, Func<ReferenceToken, string> transform, ISet<string>? names = null)
    {
        var refs = Read(formula, names);
        var text = new StringBuilder(formula);
        foreach (var token in refs.Reverse()) text.Remove(token.Start, token.Length).Insert(token.Start, transform(token));
        return text.ToString();
    }

    private static bool TryCell(string source, int at, out CellPoint? point, out int end)
    {
        var i = at;
        var ca = i < source.Length && source[i] == '$';
        if (ca) i++;
        var col = 0;
        var letters = i;
        while (i < source.Length && char.ToUpperInvariant(source[i]) is >= 'A' and <= 'Z')
        {
            if (i - letters >= 3) { point = null; end = at; return false; }
            col = col * 26 + char.ToUpperInvariant(source[i++]) - 'A' + 1;
        }
        var ra = i < source.Length && source[i] == '$';
        if (ra) i++;
        var digits = i;
        while (i < source.Length && char.IsAsciiDigit(source[i])) i++;
        end = i;
        if (letters == (ra ? digits - 1 : digits) || digits == i || col is < 1 or > 16384 ||
            !int.TryParse(source[digits..i], out var row) || row is < 1 or > 1048576)
        { point = null; return false; }
        point = new(row, col, ra, ca);
        return true;
    }

    private static void Quoted(string text, ref int at, char quote)
    {
        at++;
        while (at < text.Length)
        {
            if (text[at++] != quote) continue;
            if (at < text.Length && text[at] == quote) { at++; continue; }
            return;
        }
        throw new FormatException("Unterminated quoted text.");
    }
}
