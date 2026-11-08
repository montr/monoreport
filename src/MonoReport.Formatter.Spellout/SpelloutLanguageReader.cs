using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;

namespace MonoReport.Formatter.Spellout;

internal sealed class SpelloutLanguageReader
{
    private static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web);
    private readonly Assembly _assembly;

    public SpelloutLanguageReader() : this(typeof(SpelloutLanguageReader).Assembly)
    {
    }

    internal SpelloutLanguageReader(Assembly assembly)
    {
        ArgumentNullException.ThrowIfNull(assembly);
        _assembly = assembly;
    }

    public IReadOnlyDictionary<string, SpelloutLanguage> ReadEmbeddedLanguages()
    {
        var languages = new Dictionary<string, SpelloutLanguage>(StringComparer.Ordinal);
        foreach (var resource in _assembly.GetManifestResourceNames())
        {
            if (GetResourceLanguage(resource) is null) continue;
            using var stream = _assembly.GetManifestResourceStream(resource)
                               ?? throw new InvalidOperationException($"Unable to load spellout resource: {resource}.");
            var language = ReadEmbeddedLanguage(resource, stream);
            if (!languages.TryAdd(language.Language, language))
                throw new FormatException($"Duplicate spellout language identifier: {language.Language}.");
        }

        return languages;
    }

    internal SpelloutLanguage ReadEmbeddedLanguage(string resourceName, Stream stream)
    {
        var fileLanguage = GetResourceLanguage(resourceName)
                           ?? throw new FormatException($"Invalid spellout resource filename: {resourceName}.");
        var language = Read(stream);
        if (!StringComparer.Ordinal.Equals(language.Language, fileLanguage))
            throw new FormatException(
                $"Spellout language '{language.Language}' does not match resource '{resourceName}'.");
        return language;
    }

    public SpelloutLanguage Read(Stream stream)
    {
        SpelloutLanguage? language;
        try
        {
            language = JsonSerializer.Deserialize<SpelloutLanguage>(stream, Options);
        }
        catch (JsonException exception)
        {
            throw new FormatException("Invalid spellout language configuration.", exception);
        }

        if (language is null) throw new FormatException("Empty spellout language configuration.");
        if (!ValidLanguage(language))
            throw new FormatException("Invalid spellout language configuration.");
        return language;
    }

    private static bool ValidLanguage(SpelloutLanguage language)
    {
        return !string.IsNullOrWhiteSpace(language.Language) &&
               language.PluralRule is PluralRules.OneFewMany or PluralRules.OneOther &&
               !string.IsNullOrWhiteSpace(language.Zero) && !string.IsNullOrWhiteSpace(language.Minus) &&
               language.TensAndOnesSeparator is not null &&
               ValidWordTable(language.Ones, 1) && ValidWordTable(language.Teens, 0) &&
               ValidWordTable(language.Tens, 2) && ValidWordTable(language.Hundreds, 1) &&
               ValidGenderVariants(language.GenderVariants) &&
               language.Scales is {Length: >= 5} && language.Scales.All(ValidUnit) &&
               HasCurrencies(language.Currencies);
    }

    private static bool ValidWordTable(string[]? words, int firstUsedIndex)
    {
        return words is {Length: 10} && words.All(word => word is not null) &&
               words.Skip(firstUsedIndex).All(word => !string.IsNullOrWhiteSpace(word));
    }

    private static bool ValidGenderVariants(Dictionary<string, Dictionary<string, string>>? variants)
    {
        return variants is not null && variants.All(gender =>
            gender.Key is Genders.Masculine or Genders.Feminine or Genders.Neuter && gender.Value is not null &&
            gender.Value.All(variant => variant.Key.Length == 1 && variant.Key[0] is >= '1' and <= '9' &&
                                        !string.IsNullOrEmpty(variant.Value)));
    }

    private static bool HasCurrencies(Dictionary<string, SpelloutCurrency>? currencies)
    {
        return currencies is {Count: > 0} && currencies.All(currency =>
            !string.IsNullOrWhiteSpace(currency.Key) && ValidCurrency(currency.Value));
    }

    private static bool ValidCurrency(SpelloutCurrency? currency)
    {
        return currency is not null && ValidUnit(currency.Major) && ValidUnit(currency.Minor);
    }

    private static bool ValidUnit(SpelloutUnit? unit)
    {
        return unit is {Forms.Length: 3} && unit.Forms.All(form => !string.IsNullOrEmpty(form)) &&
               unit.Gender is Genders.Masculine or Genders.Feminine or Genders.Neuter;
    }

    private static string? GetResourceLanguage(string resource)
    {
        const string prefix = "spell-";
        const string extension = ".json";
        if (!resource.EndsWith(extension, StringComparison.Ordinal)) return null;

        var marker = resource.LastIndexOf(prefix, StringComparison.Ordinal);
        if (marker < 0 || (marker > 0 && resource[marker - 1] != '.')) return null;
        var start = marker + prefix.Length;
        var length = resource.Length - start - extension.Length;
        if (length <= 0 || resource.AsSpan(start, length).Contains('.'))
            throw new FormatException($"Invalid spellout resource filename: {resource}.");
        return resource.Substring(start, length);
    }
}