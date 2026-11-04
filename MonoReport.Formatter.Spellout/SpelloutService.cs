using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace MonoReport.Formatter.Spellout;

internal static class Modes
{
    public const string Number = "number";
    public const string Money = "money";
}

internal sealed class SpelloutService
{
    private const string FallbackLanguageCode = "en";
    
    private const decimal Limit = 1_000_000_000_000_000_000m;
    
    private readonly IReadOnlyDictionary<string, SpelloutLanguage> _languages;

    public SpelloutService(SpelloutLanguageReader reader)
    {
        ArgumentNullException.ThrowIfNull(reader);
        _languages = reader.ReadEmbeddedLanguages();
        if (!_languages.ContainsKey(FallbackLanguageCode))
            throw new FormatException("Spellout requires an English language resource for culture fallback.");
    }

    public string ResolveLanguage(CultureInfo culture)
    {
        var code = culture.TwoLetterISOLanguageName;
        return _languages.ContainsKey(code) ? code : FallbackLanguageCode;
    }

    public bool SupportsLanguage(string languageCode)
    {
        return _languages.ContainsKey(languageCode);
    }

    public bool SupportsCurrency(string languageCode, string currencyCode)
    {
        return _languages.TryGetValue(languageCode, out var language) && language.Currencies.ContainsKey(currencyCode);
    }

    public bool SupportsCurrency(string currencyCode)
    {
        return _languages.Values.Any(language => language.Currencies.ContainsKey(currencyCode));
    }

    public string Format(decimal value, string mode, string languageCode, string? currency, string gender,
        bool includeMinor, bool capitalize)
    {
        if (!_languages.TryGetValue(languageCode, out var language))
            throw new ArgumentException($"Unsupported spellout language: {languageCode}", nameof(languageCode));
        if (mode is not (Modes.Money or Modes.Number))
            throw new ArgumentException($"Unsupported spellout mode: {mode}", nameof(mode));

        SpelloutCurrency? currencyData = null;
        if (mode == Modes.Money)
        {
            if (currency is null || !language.Currencies.TryGetValue(currency, out currencyData))
                throw new ArgumentException($"Unsupported currency: {currency}", nameof(currency));
            value = decimal.Round(value, 2, MidpointRounding.AwayFromZero);
        }
        else if (decimal.Truncate(value) != value)
        {
            throw new FormatException("Number spellout supports integers only.");
        }
        else if (gender is not (Genders.Masculine or Genders.Feminine or Genders.Neuter))
        {
            throw new ArgumentException($"Unsupported gender: {gender}", nameof(gender));
        }

        if (value <= -Limit || value >= Limit)
            throw new FormatException("Spellout supports absolute values below 10^18.");

        var absolute = Math.Abs(value);
        var whole = (long) decimal.Truncate(absolute);
        var minor = mode == Modes.Money ? (int) ((absolute - whole) * 100m) : 0;
        var words = SpellInteger(whole, language, mode == Modes.Money ? currencyData!.Major.Gender : gender);
        var parts = new List<string>();
        if (value < 0) parts.Add(language.Minus);
        parts.Add(words);
        if (mode == Modes.Money)
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
            var chunk = (int) (value % 1000);
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
        
        if (rest is >= 10 and <= 19)
        {
            words.Add(language.Teens[rest - 10]);
        }
        else
        {
            var digit = rest % 10;
            if (rest >= 20 && digit != 0)
            {
                words.Add(language.Tens[rest / 10] + language.TensAndOnesSeparator + GetOne(digit, gender, language));
            }
            else
            {
                if (rest >= 20) words.Add(language.Tens[rest / 10]);
                if (digit != 0) words.Add(GetOne(digit, gender, language));
            }
        }

        return string.Join(" ", words);
    }

    private static string GetOne(int digit, string gender, SpelloutLanguage language)
    {
        if (language.GenderVariants.TryGetValue(gender, out var variants) &&
            variants.TryGetValue(digit.ToString(CultureInfo.InvariantCulture), out var variant))
            return variant;
        return language.Ones[digit];
    }

    private static string ChooseForm(long value, string[] forms, string rule)
    {
        if (rule == PluralRules.OneOther) return value == 1 ? forms[0] : forms[2];
        var lastHundred = value % 100;
        if (lastHundred is >= 11 and <= 14) return forms[2];
        return (value % 10) switch {1 => forms[0], >= 2 and <= 4 => forms[1], _ => forms[2]};
    }
}