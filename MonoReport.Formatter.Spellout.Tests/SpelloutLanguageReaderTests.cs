using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Xunit;

namespace MonoReport.Formatter.Spellout.Tests;

public sealed class SpelloutLanguageReaderTests
{
    [Fact]
    public void Reads_both_embedded_language_configurations()
    {
        var languages = new SpelloutLanguageReader().ReadEmbeddedLanguages();

        Assert.Equal(new[] {"en", "ru"}, languages.Keys.OrderBy(key => key, StringComparer.Ordinal));
    }

    [Fact]
    public void Discovers_additional_embedded_languages_and_their_configured_currencies()
    {
        var reader = new SpelloutLanguageReader(typeof(SpelloutLanguageReaderTests).Assembly);
        var languages = reader.ReadEmbeddedLanguages();

        Assert.Equal(new[] {"en", "test"}, languages.Keys.OrderBy(key => key, StringComparer.Ordinal));
        Assert.Contains("TOK", languages["test"].Currencies.Keys);
    }

    [Fact]
    public void Requires_the_language_identifier_to_match_its_resource_filename()
    {
        var configuration = EnglishConfiguration();
        configuration["language"] = "custom";

        Assert.Equal("custom", ReadEmbeddedLanguage("Languages.spell-custom.json", configuration).Language);
        Assert.Throws<FormatException>(() => ReadEmbeddedLanguage("Languages.spell-other.json", configuration));
    }

    [Theory]
    [InlineData("zero")]
    [InlineData("minus")]
    [InlineData("genderVariants")]
    public void Rejects_null_required_values(string property)
    {
        var configuration = EnglishConfiguration();
        configuration[property] = null;

        Assert.Throws<FormatException>(() => Read(configuration));
    }

    [Theory]
    [InlineData("zero")]
    [InlineData("minus")]
    [InlineData("genderVariants")]
    [InlineData("ones")]
    [InlineData("scales")]
    [InlineData("currencies")]
    public void Rejects_missing_required_values(string property)
    {
        var configuration = EnglishConfiguration();
        configuration.Remove(property);

        Assert.Throws<FormatException>(() => Read(configuration));
    }

    [Fact]
    public void Rejects_null_word_entries()
    {
        var configuration = EnglishConfiguration();
        configuration["ones"]![1] = null;

        Assert.Throws<FormatException>(() => Read(configuration));
    }

    [Fact]
    public void Rejects_null_currency_units()
    {
        var configuration = EnglishConfiguration();
        configuration["currencies"]!["EUR"]!["major"] = null;

        Assert.Throws<FormatException>(() => Read(configuration));
    }

    [Fact]
    public void Rejects_null_gender_variant_tables()
    {
        var configuration = EnglishConfiguration();
        configuration["genderVariants"]!["f"] = null;

        Assert.Throws<FormatException>(() => Read(configuration));
    }

    private static JsonObject EnglishConfiguration()
    {
        var language = new SpelloutLanguageReader().ReadEmbeddedLanguages()["en"];
        return JsonSerializer.SerializeToNode(language, new JsonSerializerOptions(JsonSerializerDefaults.Web))!
            .AsObject();
    }

    private static SpelloutLanguage Read(JsonObject configuration)
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(configuration.ToJsonString()));
        return new SpelloutLanguageReader().Read(stream);
    }

    private static SpelloutLanguage ReadEmbeddedLanguage(string resourceName, JsonObject configuration)
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(configuration.ToJsonString()));
        return new SpelloutLanguageReader().ReadEmbeddedLanguage(resourceName, stream);
    }
}