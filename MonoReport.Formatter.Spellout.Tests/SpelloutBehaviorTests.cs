using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using MonoReport.Core.Formatting;
using Xunit;

namespace MonoReport.Formatter.Spellout.Tests;

public sealed class SpelloutBehaviorTests
{
    [Theory]
    [InlineData("0", "Ноль рублей 00 копеек")]
    [InlineData("-1.005", "Минус один рубль 01 копейка")]
    [InlineData("21.21", "Двадцать один рубль 21 копейка")]
    [InlineData("2000", "Две тысячи рублей 00 копеек")]
    [InlineData("21000", "Двадцать одна тысяча рублей 00 копеек")]
    public void Formats_ruble_money_with_rounding_and_russian_grammatical_forms(string value, string expected)
    {
        Assert.Equal(expected, CreateService().Format(decimal.Parse(value, CultureInfo.InvariantCulture),
            "money", "ru", "RUB", "m", true, true));
    }

    [Theory]
    [InlineData("ru-RU", "1000000000000000", "Один квадриллион")]
    [InlineData("ru-RU", "2000000000000000", "Два квадриллиона")]
    [InlineData("ru-RU", "5000000000000000", "Пять квадриллионов")]
    [InlineData("en-US", "1000000000000000", "One quadrillion")]
    [InlineData("en-US", "1000000", "One million")]
    [InlineData("en-US", "21000000", "Twenty-one million")]
    public void Formats_quadrillion_scale_names_and_declensions(string culture, string value, string expected)
    {
        var service = CreateService();
        Assert.Equal(expected, service.Format(decimal.Parse(value, CultureInfo.InvariantCulture),
            "number", culture == "ru-RU" ? "ru" : "en", null, "m", true, true));
    }

    [Fact]
    public void Rejects_both_signs_of_the_10_to_18_boundary()
    {
        var service = CreateService();
        Assert.Throws<FormatException>(() => service.Format(1_000_000_000_000_000_000m,
            "number", "en", null, "m", true, true));
        Assert.Throws<FormatException>(() => service.Format(-1_000_000_000_000_000_000m,
            "number", "en", null, "m", true, true));
    }

    [Fact]
    public void Money_rounding_accepts_values_below_the_limit_and_rejects_values_that_round_to_it()
    {
        var service = CreateService();
        Assert.EndsWith("99 cents", service.Format(999_999_999_999_999_999.994m,
            "money", "en", "USD", "m", true, false));
        Assert.EndsWith("99 cents", service.Format(-999_999_999_999_999_999.994m,
            "money", "en", "USD", "m", true, false));
        Assert.Throws<FormatException>(() => service.Format(999_999_999_999_999_999.995m,
            "money", "en", "USD", "m", true, false));
        Assert.Throws<FormatException>(() => service.Format(-999_999_999_999_999_999.995m,
            "money", "en", "USD", "m", true, false));
    }

    [Fact]
    public void Money_rounding_carries_into_the_major_units()
    {
        Assert.Equal("two dollars 00 cents", CreateService().Format(1.995m,
            "money", "en", "USD", "m", true, false));
    }

    [Fact]
    public void Number_mode_rejects_fractional_values()
    {
        Assert.Throws<FormatException>(() => CreateService().Format(1.25m,
            "number", "en", null, "m", true, false));
    }

    [Theory]
    [InlineData("1", "one dollar")]
    [InlineData("2", "two dollars")]
    [InlineData("11", "eleven dollars")]
    [InlineData("21", "twenty-one dollars")]
    [InlineData("22", "twenty-two dollars")]
    [InlineData("101", "one hundred one dollars")]
    [InlineData("1001", "one thousand one dollars")]
    public void One_other_currency_forms_use_the_whole_unit_amount(string value, string expected)
    {
        Assert.Equal(expected, CreateService().Format(decimal.Parse(value, CultureInfo.InvariantCulture),
            "money", "en", "USD", "m", false, false));
    }

    [Theory]
    [InlineData("1", "один рубль")]
    [InlineData("2", "два рубля")]
    [InlineData("11", "одиннадцать рублей")]
    [InlineData("21", "двадцать один рубль")]
    [InlineData("22", "двадцать два рубля")]
    [InlineData("101", "сто один рубль")]
    public void One_few_many_selects_currency_forms_using_ruled_endings(string value, string expected)
    {
        Assert.Equal(expected, CreateService().Format(decimal.Parse(value, CultureInfo.InvariantCulture),
            "money", "ru", "RUB", "m", false, false));
    }

    [Fact]
    public void Formats_euros_and_neuter_gender()
    {
        using var value = JsonDocument.Parse("2.01");
        var formatter = CreateFormatter();
        Assert.Equal("Two euros 01 cent", formatter.Format(value.RootElement,
            Arguments(("mode", "money"), ("currency", "EUR")),
            new FormatterContext(CultureInfo.GetCultureInfo("en-US"))));
        Assert.Equal("Одно", CreateService().Format(1m, "number", "ru", null, "n", true, true));
    }

    [Fact]
    public void Formatter_resolves_test_language_and_currency_from_embedded_configuration()
    {
        using var value = JsonDocument.Parse("21");
        var service = new SpelloutService(new SpelloutLanguageReader(typeof(SpelloutBehaviorTests).Assembly));
        var formatter = new SpelloutFormatter(new InvariantDecimalParser(), service);
        var parameters = Arguments(("mode", "money"), ("language", "test"), ("currency", "TOK"),
            ("includeMinor", "false"), ("capitalize", "false"));

        Assert.Empty(formatter.Validate(parameters).Errors);
        Assert.Equal("twenty-one tokens", formatter.Format(value.RootElement, parameters,
            new FormatterContext(CultureInfo.InvariantCulture)));
        Assert.NotEmpty(formatter.Validate(Arguments(("mode", "money"), ("language", "test"),
            ("currency", "USD"))).Errors);
    }

    [Theory]
    [InlineData("ru-RU", "Один")]
    [InlineData("tr-TR", "One")]
    public void Formatter_defaults_to_number_and_uses_the_context_language(string culture, string expected)
    {
        using var value = JsonDocument.Parse("1");
        var formatter = CreateFormatter();

        var actual = formatter.Format(value.RootElement, Arguments(),
            new FormatterContext(CultureInfo.GetCultureInfo(culture)));

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Formatter_rejects_money_without_currency_and_currency_in_number_mode()
    {
        var formatter = CreateFormatter();
        Assert.Equal("currency is required when mode is money.",
            Assert.Single(formatter.Validate(Arguments(("mode", "money"))).Errors));
        Assert.Equal("currency applies only to money mode.",
            Assert.Single(formatter.Validate(Arguments(("currency", "EUR"))).Errors));
    }

    [Fact]
    public void Formatter_direct_formatting_revalidates_parameters()
    {
        using var value = JsonDocument.Parse("1");
        var error = Assert.Throws<FormatException>(() => CreateFormatter().Format(value.RootElement,
            Arguments(("mode", "money")), new FormatterContext(CultureInfo.InvariantCulture)));

        Assert.Equal("currency is required when mode is money.", error.Message);
    }

    private static SpelloutService CreateService()
    {
        return new SpelloutService(new SpelloutLanguageReader());
    }

    private static SpelloutFormatter CreateFormatter()
    {
        return new SpelloutFormatter(new InvariantDecimalParser(), CreateService());
    }

    private static IReadOnlyDictionary<string, FormatterOperand> Arguments(params (string Name, string Value)[] values)
    {
        return values.ToDictionary(value => value.Name,
            value => new FormatterOperand(FormatterOperandKind.Literal, value.Value), StringComparer.Ordinal);
    }
}
