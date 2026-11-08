using System.Collections.Generic;
using System.Text.Json;

namespace MonoReport.Core.Formatting;

internal sealed class UpperFormatter : IFormatter
{
    public FormatterMetadata Metadata { get; } = new([]);

    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
    {
        return FormatterValidationResult.Valid;
    }

    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
        FormatterContext context)
    {
        return context.Culture.TextInfo.ToUpper(FormatterScalarText.Read(value));
    }
}