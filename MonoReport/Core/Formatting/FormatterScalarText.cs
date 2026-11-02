using System.Text.Json;
using System;

namespace MonoReport.Core.Formatting;

internal static class FormatterScalarText
{
    internal static string Read(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString()!, JsonValueKind.Number => value.GetRawText(),
            JsonValueKind.True => "true", JsonValueKind.False => "false",
            _ => throw new FormatException("An object or array cannot be rendered as a scalar value.")
        };
    }
}