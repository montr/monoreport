using System.Globalization;
using System.Text.RegularExpressions;
using MonoReport.Models;

namespace MonoReport.Services;

public sealed partial class TemplateLanguageParser
{
    private static readonly HashSet<string> TextFormatters =
        new(["upper", "lower", "caps", "firstcap", "spellout"], StringComparer.Ordinal);

    internal TemplateParseResult Parse(string text, string location)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(location);

        var tokens = new List<TemplateToken>();
        var diagnostics = new List<StructuralDiagnostic>();
        for (var index = 0; index < text.Length;)
        {
            if (text[index] == '{')
            {
                if (index + 1 < text.Length && text[index + 1] == '{')
                {
                    tokens.Add(new(TemplateTokenKind.LiteralOpenBrace, index, 2));
                    index += 2;
                    continue;
                }

                var end = text.IndexOf('}', index + 1);
                if (end < 0)
                {
                    diagnostics.Add(Error("template.unmatched_brace", "Template expression has no closing brace.",
                        text[index..], location, index));
                    break;
                }

                var source = text[index..(end + 1)];
                var content = text[(index + 1)..end];
                if (content.Contains('{', StringComparison.Ordinal))
                {
                    diagnostics.Add(Error("template.invalid_expression", "Nested opening braces are not valid inside an expression.",
                        source, location, index));
                }
                else if (TryParseExpression(content, source, location, out var expression, out var diagnostic))
                {
                    tokens.Add(new(TemplateTokenKind.Expression, index, end - index + 1, expression));
                }
                else
                {
                    diagnostics.Add(diagnostic! with { SourceOrder = index });
                }
                index = end + 1;
                continue;
            }

            if (text[index] == '}')
            {
                if (index + 1 < text.Length && text[index + 1] == '}')
                {
                    tokens.Add(new(TemplateTokenKind.LiteralCloseBrace, index, 2));
                    index += 2;
                    continue;
                }
                diagnostics.Add(Error("template.unmatched_brace", "Template text contains an unmatched closing brace.",
                    "}", location, index));
            }
            index++;
        }

        return new(tokens, diagnostics);
    }

    private static bool TryParseExpression(string content, string source, string location,
        out TemplateExpression? expression, out StructuralDiagnostic? diagnostic)
    {
        expression = null;
        diagnostic = null;
        if (content.Length == 0 || content != content.Trim())
            return Invalid(source, location, "Expression content cannot be empty or surrounded by whitespace.", out diagnostic);

        if (content is "EndFor" or "Else" or "EndIf")
        {
            expression = new(content switch
            {
                "EndFor" => TemplateExpressionKind.EndFor,
                "Else" => TemplateExpressionKind.Else,
                _ => TemplateExpressionKind.EndIf
            }, source);
            return true;
        }

        if (content.StartsWith("For:", StringComparison.Ordinal) ||
            content.StartsWith("If:", StringComparison.Ordinal))
        {
            var isFor = content.StartsWith("For:", StringComparison.Ordinal);
            var pathText = content[(isFor ? 4 : 3)..];
            if (!TryParsePath(pathText, allowCurrentElement: false, out var path))
                return Invalid(source, location, $"'{pathText}' is not a valid template path.", out diagnostic);
            expression = new(isFor ? TemplateExpressionKind.For : TemplateExpressionKind.If, source, path);
            return true;
        }

        var separator = FindWhitespace(content);
        var pathSource = separator < 0 ? content : content[..separator];
        if (!TryParsePath(pathSource, allowCurrentElement: true, out var scalarPath))
            return Invalid(source, location, $"'{pathSource}' is not a valid template path.", out diagnostic);

        TemplateFormatter? formatter = null;
        if (separator >= 0)
        {
            var formatterSource = content[separator..].Trim();
            if (!TryParseFormatter(formatterSource, out formatter, out var reason))
                return Invalid(source, location, reason!, out diagnostic);
        }
        expression = new(TemplateExpressionKind.Field, source, scalarPath, formatter);
        return true;
    }

    private static bool Invalid(string source, string location, string message,
        out StructuralDiagnostic? diagnostic)
    {
        diagnostic = Error("template.invalid_expression", message, source, location);
        return false;
    }

    private static bool TryParsePath(string source, bool allowCurrentElement, out TemplatePath? path)
    {
        path = null;
        if (allowCurrentElement && source == ".")
        {
            path = new(TemplatePathOrigin.Current, 0, true, []);
            return true;
        }

        var origin = TemplatePathOrigin.Current;
        var parents = 0;
        if (source.StartsWith("$.", StringComparison.Ordinal))
        {
            origin = TemplatePathOrigin.Root;
            source = source[2..];
        }
        else
        {
            while (source.StartsWith("../", StringComparison.Ordinal))
            {
                parents++;
                source = source[3..];
            }
            if (source.StartsWith("./", StringComparison.Ordinal)) source = source[2..];
        }

        var segments = source.Split('.', StringSplitOptions.None);
        if (segments.Length == 0 || segments.Any(segment => !IdentifierRegex().IsMatch(segment))) return false;
        path = new(origin, parents, false, segments);
        return true;
    }

    private static bool TryParseFormatter(string source, out TemplateFormatter? formatter, out string? reason)
    {
        formatter = null;
        reason = null;
        if (TextFormatters.Contains(source))
        {
            formatter = new(source);
            return true;
        }

        foreach (var name in new[] { "number", "date" })
        {
            var prefix = name + ":\"";
            if (!source.StartsWith(prefix, StringComparison.Ordinal) || source.Length <= prefix.Length ||
                source[^1] != '"' || source[prefix.Length..^1].Contains('"', StringComparison.Ordinal)) continue;
            var argument = source[prefix.Length..^1];
            if (argument.Length == 0)
            {
                reason = $"Formatter '{name}' requires a non-empty pattern.";
                return false;
            }
            try
            {
                _ = name == "number"
                    ? 0m.ToString(argument, CultureInfo.InvariantCulture)
                    : DateTimeOffset.UnixEpoch.ToString(argument, CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                reason = $"Formatter '{name}' has an invalid .NET format pattern.";
                return false;
            }
            formatter = new(name, argument);
            return true;
        }

        reason = $"'{source}' is not a supported formatter or contains more than one formatter.";
        return false;
    }

    private static int FindWhitespace(string value)
    {
        for (var index = 0; index < value.Length; index++)
            if (char.IsWhiteSpace(value[index])) return index;
        return -1;
    }

    private static StructuralDiagnostic Error(string code, string detail, string expression, string location,
        int sourceOrder = int.MaxValue) =>
        new(code, detail, expression, expression, location, SourceOrder: sourceOrder);

    [GeneratedRegex(@"^[\p{L}_][\p{L}\p{Nd}_]*$", RegexOptions.CultureInvariant)]
    private static partial Regex IdentifierRegex();
}
