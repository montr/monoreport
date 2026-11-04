using System;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace MonoReport.Formatter.Spellout;

internal sealed class SpelloutLanguageReader
{
    private static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web);

    public SpelloutLanguage Read(Stream stream)
    {
        var language = JsonSerializer.Deserialize<SpelloutLanguage>(stream, Options)
            ?? throw new FormatException("Empty spellout language configuration.");
        if (language.Language is not ("ru" or "en") || language.PluralRule is not ("slavic" or "english") ||
            language.Ones?.Length != 10 || language.Teens?.Length != 10 ||
            language.Tens?.Length != 10 || language.Hundreds?.Length != 10 ||
            language.Scales is not { Length: >= 5 } ||
            language.Currencies is null || language.Currencies.Count == 0 ||
            language.Currencies.Any(p => !ValidUnit(p.Value.Major) || !ValidUnit(p.Value.Minor)) ||
            language.Scales.Any(s => !ValidUnit(s)))
            throw new FormatException("Invalid spellout language configuration.");
        return language;
    }

    private static bool ValidUnit(SpelloutUnit? unit) =>
        unit is { Forms.Length: 3 } && unit.Forms.All(s => s is not null) &&
        unit.Gender is "m" or "f" or "n";
}
