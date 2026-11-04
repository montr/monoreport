using System;
using System.Collections.Generic;
using System.Globalization;

namespace MonoReport.Formatter.Spellout;

internal sealed class SpelloutService(IReadOnlyDictionary<string, SpelloutLanguage> languages)
{
    private const decimal Limit = 1_000_000_000_000_000_000m;

    public string Format(decimal value, string mode, string languageCode, string? currency, string gender,
        bool includeMinor, bool capitalize)
    {
        if (!languages.TryGetValue(languageCode, out var language))
            throw new ArgumentException($"Unsupported spellout language: {languageCode}", nameof(languageCode));
        if (mode is not ("money" or "number"))
            throw new ArgumentException($"Unsupported spellout mode: {mode}", nameof(mode));

        SpelloutCurrency? currencyData = null;
        if (mode == "money")
        {
            if (currency is null || !language.Currencies.TryGetValue(currency, out currencyData))
                throw new ArgumentException($"Unsupported currency: {currency}", nameof(currency));
            value = decimal.Round(value, 2, MidpointRounding.AwayFromZero);
        }
        else if (decimal.Truncate(value) != value)
            throw new FormatException("Number spellout supports integers only.");
        else if (gender is not ("m" or "f" or "n"))
            throw new ArgumentException($"Unsupported gender: {gender}", nameof(gender));

        if (value <= -Limit || value >= Limit)
            throw new FormatException("Spellout supports absolute values below 10^18.");

        var absolute = Math.Abs(value);
        var whole = (long)decimal.Truncate(absolute);
        var minor = mode == "money" ? (int)((absolute - whole) * 100m) : 0;
        var words = SpellInteger(whole, language, mode == "money" ? currencyData!.Major.Gender : gender);
        var parts = new List<string>();
        if (value < 0) parts.Add(language.Minus);
        parts.Add(words);
        if (mode == "money")
        {
            parts.Add(ChooseForm(whole, currencyData!.Major.Forms, language.PluralRule));
            if (includeMinor)
            {
                parts.Add(minor.ToString("00", CultureInfo.InvariantCulture));
                parts.Add(ChooseForm(minor, currencyData.Minor.Forms, language.PluralRule));
            }
        }
        var result = string.Join(" ", parts);
        return capitalize ? char.ToUpperInvariant(result[0]) + result[1..] : result;
    }

    private static string SpellInteger(long value, SpelloutLanguage language, string gender)
    {
        if (value == 0) return language.Zero;
        var groups = new List<string>();
        for (var scale = 0; value != 0; scale++, value /= 1000)
        {
            var chunk = (int)(value % 1000);
            if (chunk == 0) continue;
            var groupGender = scale == 0 ? gender : language.Scales[scale - 1].Gender;
            var words = SpellGroup(chunk, groupGender, language);
            if (scale > 0)
                words += " " + ChooseForm(chunk, language.Scales[scale - 1].Forms, language.PluralRule);
            groups.Insert(0, words);
        }
        return string.Join(" ", groups);
    }

    private static string SpellGroup(int chunk, string gender, SpelloutLanguage language)
    {
        var words = new List<string>();
        if (chunk / 100 != 0) words.Add(language.Hundreds[chunk / 100]);
        var rest = chunk % 100;
        if (rest is >= 10 and <= 19) words.Add(language.Teens[rest - 10]);
        else
        {
            if (rest >= 20) words.Add(language.Tens[rest / 10]);
            var digit = rest % 10;
            if (digit != 0)
            {
                var word = language.Ones[digit];
                if (language.GenderVariants.TryGetValue(gender, out var variants) &&
                    variants.TryGetValue(digit.ToString(CultureInfo.InvariantCulture), out var variant))
                    word = variant;
                words.Add(word);
            }
        }
        return string.Join(" ", words);
    }

    private static string ChooseForm(long value, string[] forms, string rule)
    {
        if (rule == "english") return value == 1 ? forms[0] : forms[2];
        var lastHundred = value % 100;
        if (lastHundred is >= 11 and <= 14) return forms[2];
        return (value % 10) switch { 1 => forms[0], >= 2 and <= 4 => forms[1], _ => forms[2] };
    }
}
