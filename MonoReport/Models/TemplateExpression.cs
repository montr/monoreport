namespace MonoReport.Models;

internal enum TemplateExpressionKind
{
    Field,
    For,
    EndFor,
    If,
    Else,
    EndIf
}

internal enum TemplatePathOrigin
{
    Current,
    Root
}

internal sealed record TemplatePath(
    TemplatePathOrigin Origin,
    int ParentDepth,
    bool CurrentElement,
    IReadOnlyList<string> Segments);

internal sealed record TemplateFormatter(string Name, string? Argument = null);

internal sealed record TemplateExpression(
    TemplateExpressionKind Kind,
    string Source,
    TemplatePath? Path = null,
    TemplateFormatter? Formatter = null);

internal enum TemplateTokenKind
{
    Expression,
    LiteralOpenBrace,
    LiteralCloseBrace
}

internal sealed record TemplateToken(
    TemplateTokenKind Kind,
    int Start,
    int Length,
    TemplateExpression? Expression = null);

internal sealed record TemplateParseResult(
    IReadOnlyList<TemplateToken> Tokens,
    IReadOnlyList<StructuralDiagnostic> Diagnostics);
