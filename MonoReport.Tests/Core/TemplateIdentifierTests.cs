using TemplateIdentifier = MonoReport.Core.Templates.TemplateIdentifier;
using Xunit;

namespace MonoReport.Tests.Core;

public sealed class TemplateIdentifierTests
{
    [Theory]
    [InlineData("word")]
    [InlineData("_private")]
    [InlineData("Éclair2")]
    [InlineData("слово_٢")]
    public void Accepts_unicode_letters_decimal_digits_and_underscores(string identifier)
    {
        Assert.True(TemplateIdentifier.IsValid(identifier));
    }

    [Theory]
    [InlineData("")]
    [InlineData("9word")]
    [InlineData("bad-key")]
    [InlineData("bad.key")]
    [InlineData("word name")]
    [InlineData("word-")]
    public void Rejects_values_outside_template_identifier_syntax(string identifier)
    {
        Assert.False(TemplateIdentifier.IsValid(identifier));
    }

    [Fact]
    public void Rejects_null()
    {
        Assert.False(TemplateIdentifier.IsValid(null));
    }
}