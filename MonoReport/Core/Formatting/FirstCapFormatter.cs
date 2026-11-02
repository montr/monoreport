using System.Text.Json;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace MonoReport.Core.Formatting;

internal sealed partial class FirstCapFormatter : IFormatter
{
    public FormatterMetadata Metadata { get; } = new([]);

    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
    {
        return FormatterValidationResult.Valid;
    }

    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
        FormatterContext context)
    {
        var text = FormatterScalarText.Read(value);
        var match = LetterRegex().Match(text);
        return match.Success
            ? text[..match.Index] + context.Culture.TextInfo.ToUpper(match.Value) + text[(match.Index + match.Length)..]
            : text;
    }

    [GeneratedRegex(@"\p{L}")]
    private static partial Regex LetterRegex();
}