using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;

namespace MonoReport.Core.Formatting;

internal sealed class NumberFormatter(InvariantDecimalParser decimalParser) : IFormatter
{
    private static readonly FormatterMetadata MetadataValue =
        new([new FormatterParameter("pattern", typeof(string), true)], "pattern");

    public FormatterMetadata Metadata => MetadataValue;

    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
    {
        return PatternValidation.Validate(parameters, "number",
            sample => 0m.ToString(sample, CultureInfo.InvariantCulture));
    }

    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
        FormatterContext context)
    {
        return decimalParser.Parse(value).ToString(parameters["pattern"].Literal!, context.Culture);
    }
}