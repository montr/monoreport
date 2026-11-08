using System;
using System.Globalization;
using System.Text.Json;

namespace MonoReport.Core.Formatting;

public sealed class InvariantDecimalParser
{
    public decimal Parse(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.Number => value.GetDecimal(),
            JsonValueKind.String => decimal.Parse(value.GetString()!, NumberStyles.Float, CultureInfo.InvariantCulture),
            _ => throw new FormatException("A numeric formatter requires a JSON number or invariant numeric string.")
        };
    }
}