using System;
using System.Buffers;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Templates;

namespace MonoReport.Core.Formatting;

public sealed class ValueFormatter(IServiceProvider provider)
{
    private readonly IKeyedServiceProvider _formatters = provider as IKeyedServiceProvider ??
                                                         throw new ArgumentException(
                                                             "The service provider must support keyed services.",
                                                             nameof(provider));

    internal string Format(JsonElement value, TemplateExpression expression, GenerationOptions options, string location,
        out int removedCharacterCount)
    {
        removedCharacterCount = 0;
        var registration = expression.Formatter;
        var active = registration?.Name ?? "value";
        try
        {
            var formatted = Format(value, registration?.Name, registration?.Arguments,
                new FormatterContext(options.Culture));
            return SanitizeXmlText(formatted, out removedCharacterCount);
        }
        catch (Exception exception) when (exception is FormatException or ArgumentException or OverflowException
                                              or InvalidOperationException)
        {
            var code = value.ValueKind is JsonValueKind.Object or JsonValueKind.Array
                ? DiagnosticCodes.GenerationInvalidData
                : DiagnosticCodes.GenerationInvalidFormat;
            throw new TemplateException($"Cannot format '{expression.Source}' using '{active}'. {exception.Message}",
                expression.Source, expression.Source, location, active, exception, code);
        }
    }

    internal string SanitizeXmlText(string value, out int removedCharacterCount)
    {
        ArgumentNullException.ThrowIfNull(value);
        removedCharacterCount = 0;
        StringBuilder? sanitized = null;
        for (var offset = 0; offset < value.Length;)
        {
            var status = Rune.DecodeFromUtf16(value.AsSpan(offset), out var rune, out var consumed);
            if (status != OperationStatus.Done)
            {
                sanitized ??= new StringBuilder(value.Length).Append(value, 0, offset);
                removedCharacterCount++;
                offset++;
                continue;
            }

            var codePoint = rune.Value;
            var allowed = codePoint is 0x9 or 0xA or 0xD ||
                          codePoint is >= 0x20 and <= 0xD7FF ||
                          codePoint is >= 0xE000 and <= 0xFFFD ||
                          codePoint is >= 0x10000 and <= 0x10FFFF;
            if (!allowed)
            {
                sanitized ??= new StringBuilder(value.Length).Append(value, 0, offset);
                removedCharacterCount++;
            }
            else if (sanitized is not null)
            {
                sanitized.Append(value, offset, consumed);
            }

            offset += consumed;
        }

        return sanitized?.ToString() ?? value;
    }

    public string Format(JsonElement value, string? name, IReadOnlyDictionary<string, FormatterOperand>? parameters,
        FormatterContext context)
    {
        var empty = value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined ||
                    (value.ValueKind == JsonValueKind.String && value.GetString() == string.Empty);
        if (empty) return string.Empty;
        if (name is null) return FormatterScalarText.Read(value);
        var formatter = _formatters.GetRequiredKeyedService<IFormatter>(name);
        return formatter.Format(value, parameters!, context);
    }
}
