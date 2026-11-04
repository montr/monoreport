using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace MonoReport.Core.Formatting;

internal sealed partial class DateFormatter : IFormatter
{
    private static readonly FormatterMetadata MetadataValue =
        new([new FormatterParameter("pattern", typeof(string), true)], "pattern");

    public FormatterMetadata Metadata => MetadataValue;

    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
    {
        return PatternValidation.Validate(parameters, "date",
            sample => DateTimeOffset.UnixEpoch.ToString(sample, CultureInfo.InvariantCulture));
    }

    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
        FormatterContext context)
    {
        return ParseDate(value).ToString(parameters["pattern"].Literal!, context.Culture);
    }

    private static DateTimeOffset ParseDate(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.String)
            throw new FormatException("The date formatter requires an ISO 8601 string.");
        var text = value.GetString()!;
        if (DateOnlyRegex().IsMatch(text))
        {
            var date = DateOnly.ParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            return new DateTimeOffset(date.ToDateTime(TimeOnly.MinValue), TimeSpan.Zero);
        }

        if (!DateTimeRegex().IsMatch(text))
            throw new FormatException("The date value is not a supported ISO 8601 representation.");
        return DateTimeOffset.ParseExact(text, ["yyyy-MM-dd'T'HH:mm:ssK", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK"],
            CultureInfo.InvariantCulture, DateTimeStyles.None);
    }

    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}$", RegexOptions.CultureInvariant)]
    private static partial Regex DateOnlyRegex();

    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})$",
        RegexOptions.CultureInvariant)]
    private static partial Regex DateTimeRegex();
}