using System.Text.Json;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace MonoReport.Core.Formatting;

internal sealed partial class CapsFormatter : IFormatter
{
    public FormatterMetadata Metadata { get; } = new([]);

    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
    {
        return FormatterValidationResult.Valid;
    }

    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
        FormatterContext context)
    {
        return WordInitialRegex().Replace(FormatterScalarText.Read(value),
            match => context.Culture.TextInfo.ToUpper(match.Value));
    }

    [GeneratedRegex(@"\b\p{L}")]
    private static partial Regex WordInitialRegex();
}