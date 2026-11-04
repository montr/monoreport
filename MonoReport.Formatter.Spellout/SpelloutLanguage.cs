using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace MonoReport.Formatter.Spellout;

internal sealed class SpelloutLanguage
{
    public required string Language { get; init; }
    public required string Zero { get; init; }
    public required string Minus { get; init; }
    public required string PluralRule { get; init; }
    public required string[] Ones { get; init; }
    public required string[] Teens { get; init; }
    public required string[] Tens { get; init; }
    public required string[] Hundreds { get; init; }
    public Dictionary<string, Dictionary<string, string>> GenderVariants { get; init; } = new();
    public required SpelloutUnit[] Scales { get; init; }
    public required Dictionary<string, SpelloutCurrency> Currencies { get; init; }
}

internal sealed class SpelloutUnit
{
    public required string[] Forms { get; init; }
    public string Gender { get; init; } = "m";
}

internal sealed class SpelloutCurrency
{
    public required SpelloutUnit Major { get; init; }
    public required SpelloutUnit Minor { get; init; }
}
