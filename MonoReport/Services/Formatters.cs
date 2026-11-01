using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using MonoReport.Models;

namespace MonoReport.Services;

internal sealed class UpperFormatter : IFormatter
{
    public FormatterMetadata Metadata { get; } = new([]);
    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters) => FormatterValidationResult.Valid;
    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters, FormatterContext context) =>
        context.Culture.TextInfo.ToUpper(FormatterScalarText.Read(value));
}

internal sealed class LowerFormatter : IFormatter
{
    public FormatterMetadata Metadata { get; } = new([]);
    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters) => FormatterValidationResult.Valid;
    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters, FormatterContext context) =>
        context.Culture.TextInfo.ToLower(FormatterScalarText.Read(value));
}

internal sealed partial class CapsFormatter : IFormatter
{
    public FormatterMetadata Metadata { get; } = new([]);
    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters) => FormatterValidationResult.Valid;
    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters, FormatterContext context) =>
        WordInitialRegex().Replace(FormatterScalarText.Read(value), match => context.Culture.TextInfo.ToUpper(match.Value));
    [GeneratedRegex(@"\b\p{L}")] private static partial Regex WordInitialRegex();
}

internal sealed partial class FirstCapFormatter : IFormatter
{
    public FormatterMetadata Metadata { get; } = new([]);
    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters) => FormatterValidationResult.Valid;
    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters, FormatterContext context)
    {
        var text = FormatterScalarText.Read(value);
        var match = LetterRegex().Match(text);
        return match.Success ? text[..match.Index] + context.Culture.TextInfo.ToUpper(match.Value) + text[(match.Index + match.Length)..] : text;
    }
    [GeneratedRegex(@"\p{L}")] private static partial Regex LetterRegex();
}

internal static class FormatterScalarText
{
    internal static string Read(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.String => value.GetString()!, JsonValueKind.Number => value.GetRawText(),
        JsonValueKind.True => "true", JsonValueKind.False => "false",
        _ => throw new FormatException("An object or array cannot be rendered as a scalar value.")
    };
}

internal sealed class NumberFormatter(InvariantDecimalParser decimalParser) : IFormatter
{
    private static readonly FormatterMetadata MetadataValue = new([new("pattern", typeof(string), true)], "pattern");
    public FormatterMetadata Metadata => MetadataValue;
    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters) =>
        PatternValidation.Validate(parameters, "number", sample => 0m.ToString(sample, CultureInfo.InvariantCulture));
    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters, FormatterContext context) =>
        decimalParser.Parse(value).ToString(parameters["pattern"].Literal!, context.Culture);
}

internal sealed partial class DateFormatter : IFormatter
{
    private static readonly FormatterMetadata MetadataValue = new([new("pattern", typeof(string), true)], "pattern");
    public FormatterMetadata Metadata => MetadataValue;
    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters) =>
        PatternValidation.Validate(parameters, "date", sample => DateTimeOffset.UnixEpoch.ToString(sample, CultureInfo.InvariantCulture));
    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters, FormatterContext context) =>
        ParseDate(value).ToString(parameters["pattern"].Literal!, context.Culture);

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
        return DateTimeOffset.ParseExact(text, ["yyyy-MM-dd'T'HH:mm:ssK", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK"],
            CultureInfo.InvariantCulture, DateTimeStyles.None);
    }
    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}$", RegexOptions.CultureInvariant)] private static partial Regex DateOnlyRegex();
    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})$", RegexOptions.CultureInvariant)] private static partial Regex DateTimeRegex();
}

public sealed class InvariantDecimalParser
{
    public decimal Parse(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Number => value.GetDecimal(),
        JsonValueKind.String => decimal.Parse(value.GetString()!, NumberStyles.Float, CultureInfo.InvariantCulture),
        _ => throw new FormatException("A numeric formatter requires a JSON number or invariant numeric string.")
    };
}

internal static class PatternValidation
{
    internal static FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters,
        string name, Func<string, string> format)
    {
        if (!parameters.TryGetValue("pattern", out var operand) || operand.Kind != FormatterOperandKind.Literal)
            return FormatterValidationResult.Valid;
        if (operand.Literal!.Length == 0) return new([$"Formatter '{name}' requires a non-empty pattern."]);
        try { _ = format(operand.Literal); return FormatterValidationResult.Valid; }
        catch (FormatException) { return new([$"Formatter '{name}' has an invalid .NET format pattern."]); }
    }
}
