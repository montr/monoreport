using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;

namespace MonoReport.Services;

public sealed class ValueFormatter(IServiceProvider provider)
{
    private readonly IKeyedServiceProvider _formatters = provider as IKeyedServiceProvider ?? throw new ArgumentException("The service provider must support keyed services.", nameof(provider));

    internal string Format(JsonElement value, TemplateExpression expression, MergeOptions options, string location)
    {
        var registration = expression.Formatter;
        var active = registration?.Name ?? "value";
        try
        {
            var empty = value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined ||
                value.ValueKind == JsonValueKind.String && value.GetString() == string.Empty;
            if (empty) return string.Empty;
            if (registration is null) return FormatterScalarText.Read(value);
            var formatter = _formatters.GetRequiredKeyedService<IFormatter>(registration.Name);
            return formatter.Format(value, registration.Arguments, new(options.Culture));
        }
        catch (Exception exception) when (exception is FormatException or ArgumentException or OverflowException or InvalidOperationException)
        {
            var code = value.ValueKind is JsonValueKind.Object or JsonValueKind.Array
                ? DiagnosticCodes.GenerationInvalidData : DiagnosticCodes.GenerationInvalidFormat;
            throw new TemplateException($"Cannot format '{expression.Source}' using '{active}'. {exception.Message}",
                expression.Source, expression.Source, location, active, exception, code);
        }
    }
}
