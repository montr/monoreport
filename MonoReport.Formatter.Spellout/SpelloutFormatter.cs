using System.Text.Json;
using System.Collections.Generic;
using MonoReport.Core.Formatting;

namespace MonoReport.Formatter.Spellout;

internal sealed class SpelloutFormatter(InvariantDecimalParser decimalParser, RubSpellOutFormatter spellout)
    : IFormatter
{
    public FormatterMetadata Metadata { get; } = new([]);

    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
    {
        return FormatterValidationResult.Valid;
    }

    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
        FormatterContext context)
    {
        return spellout.Format(decimalParser.Parse(value));
    }
}
