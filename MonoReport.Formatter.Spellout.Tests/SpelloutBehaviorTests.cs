using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Core.Templates;
using Xunit;
using System.IO;

namespace MonoReport.Formatter.Spellout.Tests;

public sealed class SpelloutBehaviorTests
{
    [Theory]
    [InlineData("0", "Ноль рублей 00 копеек")]
    [InlineData("-1.005", "Минус один рубль 01 копейка")]
    [InlineData("2.5", "Два рубля 50 копеек")]
    [InlineData("5.05", "Пять рублей 05 копеек")]
    [InlineData("11.11", "Одиннадцать рублей 11 копеек")]
    [InlineData("21.21", "Двадцать один рубль 21 копейка")]
    [InlineData("22.22", "Двадцать два рубля 22 копейки")]
    [InlineData("25.25", "Двадцать пять рублей 25 копеек")]
    [InlineData("1000", "Одна тысяча рублей 00 копеек")]
    [InlineData("2000", "Две тысячи рублей 00 копеек")]
    [InlineData("11000", "Одиннадцать тысяч рублей 00 копеек")]
    [InlineData("21000", "Двадцать одна тысяча рублей 00 копеек")]
    public void Formats_rubles_with_rounding_and_russian_grammatical_forms(string value, string expected)
    {
        Assert.Equal(expected, RenderMoney(value, "ru-RU"));
    }

    [Fact]
    public void Accepts_the_upper_boundary_and_rejects_the_first_rounded_out_of_range_value()
    {
        Assert.EndsWith("99 копеек", RenderMoney("999999999999999999.994", "ru-RU"));
        var error = Assert.Throws<TemplateException>(() => RenderMoney("999999999999999999.995", "ru-RU"));
        Assert.Equal("generation.invalid_format", error.Code);
        error = Assert.Throws<TemplateException>(() => RenderMoney("-1000000000000000000", "ru-RU"));
        Assert.Equal("generation.invalid_format", error.Code);
    }

    [Theory]
    [InlineData("ru-RU", "Один")]
    [InlineData("en-US", "One")]
    [InlineData("tr-TR", "One")]
    public void Defaults_to_number_and_selects_language_from_culture_with_english_fallback(string culture,
        string expected)
    {
        Assert.Equal(expected, Render("1", culture));
    }

    [Fact]
    public void Supports_money_rounding_currency_pluralization_and_omitting_minor_units()
    {
        Assert.Equal("Minus one dollar 01 cent", Render("-1.005", "en-US",
            "{Value spellout mode=\"money\" language=\"en\" currency=\"USD\"}"));
        Assert.Equal("Two dollars", Render("1.995", "en-US",
            "{Value spellout mode=\"money\" language=\"en\" currency=\"USD\" includeMinor=\"false\"}"));
        Assert.Equal("Ноль рублей 00 копеек", RenderMoney("0", "ru-RU"));
        Assert.Equal("One ruble 00 copecks", Render("1", "tr-TR",
            "{Value spellout mode=\"money\" currency=\"RUB\"}"));
    }

    [Theory]
    [InlineData("ru-RU", "1000000000000000", "Один квадриллион")]
    [InlineData("ru-RU", "2000000000000000", "Два квадриллиона")]
    [InlineData("ru-RU", "5000000000000000", "Пять квадриллионов")]
    [InlineData("en-US", "1000000000000000", "One quadrillion")]
    public void Supports_quadrillion_scale_and_rejects_the_10_to_18_boundary(string culture, string value,
        string expected)
    {
        Assert.Equal(expected, Render(value, culture));
        var error = Assert.Throws<TemplateException>(() => Render("1000000000000000000", culture));
        Assert.Equal("generation.invalid_format", error.Code);
    }

    [Fact]
    public void Supports_integer_number_mode_gender_and_capitalization()
    {
        Assert.Equal("двадцать одна", Render("21", "en-US",
            "{Value spellout mode=\"number\" language=\"ru\" gender=\"f\" capitalize=\"false\"}"));
        Assert.Equal("Minus twenty one", Render("-21", "en-US",
            "{Value spellout mode=\"number\" language=\"en\"}"));
    }

    [Theory]
    [InlineData("{Value spellout mode=\"money\" currency=\"RUB\" gender=\"f\"}")]
    [InlineData("{Value spellout mode=\"number\" currency=\"USD\"}")]
    [InlineData("{Value spellout currency=\"USD\"}")]
    [InlineData("{Value spellout mode=\"money\"}")]
    [InlineData("{Value spellout mode=\"money\" currency=\"RUB\" language=\"fr\"}")]
    [InlineData("{Value spellout mode=\"money\" currency=\"JPY\"}")]
    public void Rejects_invalid_values_and_inapplicable_parameter_combinations(string expression)
    {
        var error = Assert.Throws<TemplateException>(() => Render("1", "en-US", expression));
        Assert.Equal("template.invalid_expression", error.Code);
    }

    [Fact]
    public void Number_mode_rejects_fractional_values()
    {
        var error = Assert.Throws<TemplateException>(() => Render("1.5", "en-US",
            "{Value spellout mode=\"number\"}"));
        Assert.Equal("generation.invalid_format", error.Code);
    }

    [Fact]
    public void Unsupported_json_values_return_the_generation_diagnostic()
    {
        var error = Assert.Throws<TemplateException>(() => Render("true"));
        Assert.Equal("generation.invalid_format", error.Code);
        Assert.Equal("spellout", Assert.Single(error.Errors).Formatter);
    }

    private static string Render(string value, string? culture = null)
    {
        return Render(value, culture, "{Value spellout}");
    }

    private static string RenderMoney(string value, string? culture = null)
    {
        return Render(value, culture, "{Value spellout mode=\"money\" currency=\"RUB\"}");
    }

    private static string Render(string value, string? culture, string expression)
    {
        using var provider = new ServiceCollection().AddMonoReport().AddMonoReportSpelloutFormatter()
            .BuildServiceProvider();
        using var input = new MemoryStream(CreateDocx(expression));
        using var data = JsonDocument.Parse("{\"Value\":" + value + "}");
        var result = provider.GetRequiredService<DocumentGenerator>().Generate(input, data,
            new GenerationOptions {Culture = CultureInfo.GetCultureInfo(culture ?? "en-US")});
        using var documentStream = new MemoryStream(result.Document);
        using var document = WordprocessingDocument.Open(documentStream, false);
        return document.MainDocumentPart!.Document!.Body!.InnerText;
    }

    internal static byte[] CreateDocx(string expression)
    {
        using var output = new MemoryStream();
        using (var document = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text(expression)))));
        }

        return output.ToArray();
    }
}
