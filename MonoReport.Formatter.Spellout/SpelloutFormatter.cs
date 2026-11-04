using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using MonoReport.Core.Formatting;

namespace MonoReport.Formatter.Spellout;

internal sealed class SpelloutFormatter(InvariantDecimalParser decimalParser, SpelloutService service) : IFormatter
{
    public FormatterMetadata Metadata { get; } = new([
        new FormatterParameter("mode", typeof(string), false),
        new FormatterParameter("language", typeof(string), false),
        new FormatterParameter("currency", typeof(string), false),
        new FormatterParameter("gender", typeof(string), false),
        new FormatterParameter("includeMinor", typeof(bool), false),
        new FormatterParameter("capitalize", typeof(bool), false)
    ]);

    public FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters)
    {
        var errors = new List<string>();
        foreach (var (name, operand) in parameters)
        {
            if (operand.Kind == FormatterOperandKind.Reference)
            {
                errors.Add($"Spellout parameter '{name}' must be a literal.");
                continue;
            }
            var value = operand.Literal;
            var allowed = name switch
            {
                "mode" => value is "money" or "number",
                "language" => value is "ru" or "en",
                "currency" => value is "RUB" or "USD" or "EUR",
                "gender" => value is "m" or "f" or "n",
                "includeMinor" or "capitalize" => bool.TryParse(value, out _),
                _ => false
            };
            if (!allowed) errors.Add($"Invalid spellout parameter '{name}': '{value}'.");
        }
        var mode = parameters.TryGetValue("mode", out var modeOperand) &&
                   modeOperand.Kind == FormatterOperandKind.Literal
            ? modeOperand.Literal
            : "number";
        if (mode == "money" && !parameters.ContainsKey("currency"))
            errors.Add("currency is required when mode is money.");
        if (mode == "number" && parameters.ContainsKey("currency"))
            errors.Add("currency applies only to money mode.");
        if (mode == "number" && parameters.ContainsKey("includeMinor"))
            errors.Add("includeMinor applies only to money mode.");
        if (mode == "money" && parameters.ContainsKey("gender"))
            errors.Add("gender applies only to number mode.");
        return errors.Count == 0 ? FormatterValidationResult.Valid : new FormatterValidationResult(errors);
    }

    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
        FormatterContext context)
    {
        var error = Validate(parameters);
        if (error.Errors.Count != 0) throw new FormatException(string.Join("; ", error.Errors));
        var mode = Get(parameters, "mode") ?? "number";
        var language = Get(parameters, "language") ?? ResolveLanguage(context.Culture);
        var currency = Get(parameters, "currency");
        var gender = Get(parameters, "gender") ?? "m";
        var includeMinor = GetBoolean(parameters, "includeMinor", true);
        var capitalize = GetBoolean(parameters, "capitalize", true);
        return service.Format(decimalParser.Parse(value), mode, language, currency, gender, includeMinor, capitalize);
    }

    private static string ResolveLanguage(CultureInfo culture) =>
        culture.TwoLetterISOLanguageName == "ru" ? "ru" : "en";

    private static string? Get(IReadOnlyDictionary<string, FormatterOperand> parameters, string name)
    {
        if (!parameters.TryGetValue(name, out var operand)) return null;
        if (operand.Kind == FormatterOperandKind.Reference)
            throw new NotSupportedException($"Unresolved spellout parameter reference: {name}.");
        return operand.Literal;
    }

    private static bool GetBoolean(IReadOnlyDictionary<string, FormatterOperand> parameters, string name, bool fallback) =>
        bool.Parse(Get(parameters, name) ?? fallback.ToString());
}
