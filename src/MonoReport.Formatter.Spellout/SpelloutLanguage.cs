using System.Collections.Generic;

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
    public required Dictionary<string, Dictionary<string, string>> GenderVariants { get; init; }
    public string TensAndOnesSeparator { get; init; } = " ";
    public required SpelloutUnit[] Scales { get; init; }
    public required Dictionary<string, SpelloutCurrency> Currencies { get; init; }
}

// ReSharper disable once ClassNeverInstantiated.Global
internal sealed class SpelloutUnit
{
    public required string[] Forms { get; init; }
    public string Gender { get; init; } = Genders.Masculine;
}

// ReSharper disable once ClassNeverInstantiated.Global
internal sealed class SpelloutCurrency
{
    public required SpelloutUnit Major { get; init; }
    public required SpelloutUnit Minor { get; init; }
}

internal static class Genders
{
    public const string Masculine = "m";
    public const string Feminine = "f";
    public const string Neuter = "n";
}

internal static class PluralRules
{
    public const string OneFewMany = "one-few-many";
    public const string OneOther = "one-other";
}