using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class TemplateLanguageParserTests
{
    private readonly TemplateLanguageParser _parser = new();

    [Fact]
    public void Parses_paths_controls_formatters_and_strict_escapes()
    {
        var result = _parser.Parse("{{{Customer.Name upper}}} {For:Orders} {../Title} {$.Company} {/For}", "part");

        Assert.Empty(result.Diagnostics);
        Assert.Equal(new[]
        {
            TemplateTokenKind.LiteralOpenBrace, TemplateTokenKind.Expression,
            TemplateTokenKind.LiteralCloseBrace, TemplateTokenKind.Expression,
            TemplateTokenKind.Expression, TemplateTokenKind.Expression, TemplateTokenKind.Expression
        }, result.Tokens.Select(token => token.Kind));
        var field = result.Tokens[1].Expression!;
        Assert.Equal(TemplateExpressionKind.Field, field.Kind);
        Assert.Equal(new[] { "Customer", "Name" }, field.Path!.Segments);
        Assert.Equal("upper", field.Formatter!.Name);
        Assert.Equal(1, result.Tokens[4].Expression!.Path!.ParentDepth);
        Assert.Equal(TemplatePathOrigin.Root, result.Tokens[5].Expression!.Path!.Origin);
    }

    [Theory]
    [InlineData("{.}", ".")]
    [InlineData("{./Name}", "Name")]
    [InlineData("{Имя_1}", "Имя_1")]
    public void Parses_current_and_unicode_paths(string template, string segment)
    {
        var result = _parser.Parse(template, "part");

        Assert.Empty(result.Diagnostics);
        var path = Assert.Single(result.Tokens).Expression!.Path!;
        if (template == "{.}") Assert.True(path.CurrentElement);
        else Assert.Equal(segment, Assert.Single(path.Segments));
    }

    [Theory]
    [InlineData("{")]
    [InlineData("}")]
    [InlineData("{Name unknown}")]
    [InlineData("{Name upper lower}")]
    [InlineData("{For:Items upper}")]
    [InlineData("{Name number:\"\"}")]
    [InlineData("{bad-name}")]
    public void Rejects_malformed_or_unsupported_syntax(string template)
    {
        Assert.NotEmpty(_parser.Parse(template, "part").Diagnostics);
    }

    [Fact]
    public void Parses_exactly_one_pattern_formatter()
    {
        var result = _parser.Parse("{Amount number:\"0.00\"}{When date:\"yyyy-MM-dd\"}", "part");

        Assert.Empty(result.Diagnostics);
        Assert.Equal("0.00", result.Tokens[0].Expression!.Formatter!.Argument);
        Assert.Equal("yyyy-MM-dd", result.Tokens[1].Expression!.Formatter!.Argument);
    }

    [Fact]
    public void Escapes_adjacent_expressions_and_triple_braces_have_unambiguous_precedence()
    {
        var result = _parser.Parse("{{Name}}|{{{Name}}}|{A}{B}|}}", "part");
        Assert.Empty(result.Diagnostics);
        Assert.Equal(new[]
        {
            TemplateTokenKind.LiteralOpenBrace, TemplateTokenKind.LiteralCloseBrace,
            TemplateTokenKind.LiteralOpenBrace, TemplateTokenKind.Expression, TemplateTokenKind.LiteralCloseBrace,
            TemplateTokenKind.Expression, TemplateTokenKind.Expression, TemplateTokenKind.LiteralCloseBrace
        }, result.Tokens.Select(item => item.Kind));
    }
}
