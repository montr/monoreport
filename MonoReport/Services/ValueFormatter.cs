using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using MonoReport.Models;

namespace MonoReport.Services;

public sealed partial class ValueFormatter(RubSpellOutFormatter spellOut)
{
    private readonly RubSpellOutFormatter _spellOut = spellOut ?? throw new ArgumentNullException(nameof(spellOut));

    internal string Format(JsonElement value, TemplateExpression expression, MergeOptions options, string location)
    {
        var formatter = expression.Formatter;
        var active = formatter?.Name ?? "value";
        try
        {
            var empty = value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined ||
                value.ValueKind == JsonValueKind.String && value.GetString() == string.Empty;
            if (empty) return string.Empty;
            if (formatter?.Name is "number" or "spellout")
            {
                var number = ParseDecimal(value);
                return formatter.Name == "number"
                    ? number.ToString(formatter.Argument!, options.Culture)
                    : _spellOut.Format(number);
            }
            if (formatter?.Name == "date") return ParseDate(value).ToString(formatter.Argument!, options.Culture);

            var text = ScalarText(value);
            return formatter?.Name switch
            {
                null => text,
                "upper" => options.Culture.TextInfo.ToUpper(text),
                "lower" => options.Culture.TextInfo.ToLower(text),
                "caps" => WordInitialRegex().Replace(text, match => options.Culture.TextInfo.ToUpper(match.Value)),
                "firstcap" => FirstCap(text, options.Culture),
                _ => throw new FormatException("Unsupported formatter.")
            };
        }
        catch (Exception exception) when (exception is FormatException or ArgumentException or OverflowException or InvalidOperationException)
        {
            var code = value.ValueKind is JsonValueKind.Object or JsonValueKind.Array
                ? DiagnosticCodes.GenerationInvalidData : DiagnosticCodes.GenerationInvalidFormat;
            throw new TemplateException($"Cannot format '{expression.Source}' using '{active}'. {exception.Message}",
                expression.Source, expression.Source, location, active, exception, code);
        }
    }

    private static string ScalarText(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.String => value.GetString()!,
        JsonValueKind.Number => value.GetRawText(),
        JsonValueKind.True => "true",
        JsonValueKind.False => "false",
        _ => throw new FormatException("An object or array cannot be rendered as a scalar value.")
    };

    private static decimal ParseDecimal(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Number => value.GetDecimal(),
        JsonValueKind.String => decimal.Parse(value.GetString()!, NumberStyles.Float, CultureInfo.InvariantCulture),
        _ => throw new FormatException("A numeric formatter requires a JSON number or invariant numeric string.")
    };

    private static DateTimeOffset ParseDate(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.String) throw new FormatException("The date formatter requires an ISO 8601 string.");
        var text = value.GetString()!;
        if (DateOnlyRegex().IsMatch(text))
        {
            var date = DateOnly.ParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None);
            return new(date.ToDateTime(TimeOnly.MinValue), TimeSpan.Zero);
        }
        if (!DateTimeRegex().IsMatch(text)) throw new FormatException("The date value is not a supported ISO 8601 representation.");
        return DateTimeOffset.ParseExact(text,
            ["yyyy-MM-dd'T'HH:mm:ssK", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK"],
            CultureInfo.InvariantCulture, DateTimeStyles.None);
    }

    private static string FirstCap(string text, CultureInfo culture)
    {
        var match = LetterRegex().Match(text);
        return match.Success ? text[..match.Index] + culture.TextInfo.ToUpper(match.Value) + text[(match.Index + match.Length)..] : text;
    }

    [GeneratedRegex(@"\b\p{L}")]
    private static partial Regex WordInitialRegex();
    [GeneratedRegex(@"\p{L}")]
    private static partial Regex LetterRegex();
    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}$", RegexOptions.CultureInvariant)]
    private static partial Regex DateOnlyRegex();
    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})$", RegexOptions.CultureInvariant)]
    private static partial Regex DateTimeRegex();
}
