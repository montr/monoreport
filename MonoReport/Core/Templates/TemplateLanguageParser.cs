using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Formatting;

namespace MonoReport.Core.Templates;

public sealed class TemplateLanguageParser(IServiceProvider provider, FormatterParameterBinder parameterBinder)
{
    private readonly IKeyedServiceProvider _formatters = provider as IKeyedServiceProvider ??
                                                         throw new ArgumentException(
                                                             "The service provider must support keyed services.",
                                                             nameof(provider));

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
                    tokens.Add(new TemplateToken(TemplateTokenKind.LiteralOpenBrace, index, 2));
                    index += 2;
                    continue;
                }

                var end = text.IndexOf('}', index + 1);
                if (end < 0)
                {
                    diagnostics.Add(Error(DiagnosticCodes.TemplateUnmatchedBrace,
                        "Template expression has no closing brace.",
                        text[index..], location, index));
                    break;
                }

                var source = text[index..(end + 1)];
                var content = text[(index + 1)..end];
                if (content.Contains('{', StringComparison.Ordinal))
                    diagnostics.Add(Error(DiagnosticCodes.TemplateInvalidExpression,
                        "Nested opening braces are not valid inside an expression.",
                        source, location, index));
                else if (TryParseExpression(content, source, location, out var expression, out var diagnostic))
                    tokens.Add(new TemplateToken(TemplateTokenKind.Expression, index, end - index + 1, expression));
                else
                    diagnostics.Add(diagnostic! with {SourceOrder = index});
                index = end + 1;
                continue;
            }

            if (text[index] == '}')
            {
                if (index + 1 < text.Length && text[index + 1] == '}')
                {
                    tokens.Add(new TemplateToken(TemplateTokenKind.LiteralCloseBrace, index, 2));
                    index += 2;
                    continue;
                }

                diagnostics.Add(Error(DiagnosticCodes.TemplateUnmatchedBrace,
                    "Template text contains an unmatched closing brace.",
                    "}", location, index));
            }

            index++;
        }

        return new TemplateParseResult(tokens, diagnostics);
    }

    private bool TryParseExpression(string content, string source, string location,
        out TemplateExpression? expression, out StructuralDiagnostic? diagnostic)
    {
        expression = null;
        diagnostic = null;
        if (content.Length == 0 || content != content.Trim())
            return Invalid(source, location, "Expression content cannot be empty or surrounded by whitespace.",
                out diagnostic);

        if (content is "/For" or "Else" or "/If" or "Raw" or "/Raw")
        {
            expression = new TemplateExpression(content switch
            {
                "/For" => TemplateExpressionKind.EndFor,
                "Else" => TemplateExpressionKind.Else,
                "/If" => TemplateExpressionKind.EndIf,
                "Raw" => TemplateExpressionKind.Raw,
                _ => TemplateExpressionKind.EndRaw
            }, source);
            return true;
        }

        if (content.StartsWith("For:", StringComparison.Ordinal) ||
            content.StartsWith("If:", StringComparison.Ordinal))
        {
            var isFor = content.StartsWith("For:", StringComparison.Ordinal);
            var pathText = content[(isFor ? 4 : 3)..];
            if (!TryParsePath(pathText, false, out var path))
                return Invalid(source, location, $"'{pathText}' is not a valid template path.", out diagnostic);
            expression = new TemplateExpression(isFor ? TemplateExpressionKind.For : TemplateExpressionKind.If, source,
                path);
            return true;
        }

        var separator = FindWhitespace(content);
        var pathSource = separator < 0 ? content : content[..separator];
        if (!TryParsePath(pathSource, true, out var scalarPath))
            return Invalid(source, location, $"'{pathSource}' is not a valid template path.", out diagnostic);

        TemplateFormatter? formatter = null;
        if (separator >= 0)
        {
            var formatterSource = content[separator..].Trim();
            if (!TryParseFormatter(formatterSource, out formatter, out var reason))
                return Invalid(source, location, reason!, out diagnostic, formatterSource.Split(null, 2)[0]);
        }

        expression = new TemplateExpression(TemplateExpressionKind.Field, source, scalarPath, formatter);
        return true;
    }

    private static bool Invalid(string source, string location, string message,
        out StructuralDiagnostic? diagnostic, string? formatter = null)
    {
        diagnostic = Error(DiagnosticCodes.TemplateInvalidExpression, message, source, location, formatter: formatter);
        return false;
    }

    private static bool TryParsePath(string source, bool allowCurrentElement, out TemplatePath? path)
    {
        path = null;
        if (allowCurrentElement && source == ".")
        {
            path = new TemplatePath(TemplatePathOrigin.Current, 0, true, []);
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

        var segments = source.Split('.');
        if (segments.Length == 0 || segments.Any(segment => !TemplateIdentifier.IsValid(segment))) return false;
        path = new TemplatePath(origin, parents, false, segments);
        return true;
    }

    private bool TryParseFormatter(string source, out TemplateFormatter? formatter, out string? reason)
    {
        formatter = null;
        reason = null;
        var cursor = 0;
        while (cursor < source.Length && !char.IsWhiteSpace(source[cursor])) cursor++;
        var name = source[..cursor];
        if (!TemplateIdentifier.IsValid(name))
        {
            reason = $"'{source}' is not a valid formatter invocation.";
            return false;
        }

        var selected = _formatters.GetKeyedService<IFormatter>(name);
        if (selected is null)
        {
            reason = $"Formatter '{name}' is not registered.";
            return false;
        }

        var metadata = selected.Metadata;
        var supplied = new List<KeyValuePair<string, FormatterOperand>>();
        string literal;
        var firstArgument = true;
        if (cursor < source.Length && char.IsWhiteSpace(source[cursor]))
        {
            while (cursor < source.Length && char.IsWhiteSpace(source[cursor])) cursor++;
            if (cursor < source.Length && source[cursor] == '"')
            {
                if (metadata.PositionalParameter is null)
                {
                    reason = $"Formatter '{name}' does not accept a positional argument.";
                    return false;
                }

                if (!ReadLiteral(source, ref cursor, out literal))
                {
                    reason = $"Formatter '{name}' requires double-quoted string arguments.";
                    return false;
                }

                supplied.Add(new KeyValuePair<string, FormatterOperand>(metadata.PositionalParameter,
                    new FormatterOperand(FormatterOperandKind.Literal, literal)));
                firstArgument = false;
            }

            while (cursor < source.Length)
            {
                if (!firstArgument && !char.IsWhiteSpace(source[cursor]))
                {
                    reason = "Formatter arguments must be separated by whitespace.";
                    return false;
                }

                if (!firstArgument)
                    while (cursor < source.Length && char.IsWhiteSpace(source[cursor]))
                        cursor++;
                if (cursor == source.Length) break;
                firstArgument = false;
                var start = cursor;
                while (cursor < source.Length &&
                       (char.IsLetterOrDigit(source[cursor]) || source[cursor] == '_')) cursor++;
                var parameter = source[start..cursor];
                if (!TemplateIdentifier.IsValid(parameter) || cursor >= source.Length || source[cursor++] != '=')
                {
                    reason = $"Formatter '{name}' arguments must use name=\"value\" syntax.";
                    return false;
                }

                if (cursor >= source.Length || source[cursor++] != '"' ||
                    !ReadLiteralBody(source, ref cursor, out literal))
                {
                    reason = $"Formatter '{name}' requires double-quoted string arguments.";
                    return false;
                }

                supplied.Add(new KeyValuePair<string, FormatterOperand>(parameter,
                    new FormatterOperand(FormatterOperandKind.Literal, literal)));
            }
        }

        var validation = parameterBinder.Bind(name, selected, supplied, out var bound);
        if (validation.Errors.Count > 0)
        {
            reason = validation.Errors[0];
            return false;
        }

        formatter = new TemplateFormatter(name, bound);
        return true;
    }

    private static bool ReadLiteral(string source, ref int cursor, out string value)
    {
        cursor++;
        return ReadLiteralBody(source, ref cursor, out value);
    }

    private static bool ReadLiteralBody(string source, ref int cursor, out string value)
    {
        var start = cursor;
        while (cursor < source.Length && source[cursor] != '"' && source[cursor] is not ('{' or '}')) cursor++;
        if (cursor >= source.Length || source[cursor] != '"')
        {
            value = string.Empty;
            return false;
        }

        value = source[start..cursor];
        cursor++;
        return true;
    }

    private static int FindWhitespace(string value)
    {
        for (var index = 0; index < value.Length; index++)
            if (char.IsWhiteSpace(value[index]))
                return index;
        return -1;
    }

    private static StructuralDiagnostic Error(string code, string detail, string expression, string location,
        int sourceOrder = int.MaxValue, string? formatter = null)
    {
        return new StructuralDiagnostic(code, detail, expression, expression, location, formatter, sourceOrder);
    }
}
