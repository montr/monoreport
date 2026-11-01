using System.Text.Json;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Formatter.Spellout;

internal sealed class SpelloutFormatter(InvariantDecimalParser decimalParser, RubSpellOutFormatter spellout) : IFormatter
{
    public FormatterMetadata Metadata { get; } = new([]);
    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters) => FormatterValidationResult.Valid;
    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters, FormatterContext context) =>
        spellout.Format(decimalParser.Parse(value));
}
