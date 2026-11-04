using System;
using System.Collections.Generic;
using System.Text.Json;
using MonoReport.Core.Formatting;

namespace MonoReport.Formatter.Spellout;

internal sealed class SpelloutFormatter(InvariantDecimalParser decimalParser, SpelloutService service) : IFormatter
{
    private static class Parameters
    {
        public const string Mode = "mode";
        public const string Language = "language";
        public const string Currency = "currency";
        public const string Gender = "gender";
        public const string IncludeMinor = "includeMinor";
        public const string Capitalize = "capitalize";
    }

    public FormatterMetadata Metadata { get; } = new([
        new FormatterParameter(Parameters.Mode, typeof(string), false),
        new FormatterParameter(Parameters.Language, typeof(string), false),
        new FormatterParameter(Parameters.Currency, typeof(string), false),
        new FormatterParameter(Parameters.Gender, typeof(string), false),
        new FormatterParameter(Parameters.IncludeMinor, typeof(bool), false),
        new FormatterParameter(Parameters.Capitalize, typeof(bool), false)
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
                Parameters.Mode => value is Modes.Money or Modes.Number,
                Parameters.Language => value is not null && service.SupportsLanguage(value),
                Parameters.Currency => value is not null,
                Parameters.Gender => value is Genders.Masculine or Genders.Feminine or Genders.Neuter,
                Parameters.IncludeMinor or Parameters.Capitalize => bool.TryParse(value, out _),
                _ => false
            };
            if (!allowed) errors.Add($"Invalid spellout parameter '{name}': '{value}'.");
        }
        var mode = parameters.TryGetValue(Parameters.Mode, out var modeOperand) &&
                   modeOperand.Kind == FormatterOperandKind.Literal
            ? modeOperand.Literal
            : Modes.Number;
        if (mode == Modes.Money && !parameters.ContainsKey(Parameters.Currency))
            errors.Add($"{Parameters.Currency} is required when {Parameters.Mode} is {Modes.Money}.");
        if (mode == Modes.Number && parameters.ContainsKey(Parameters.Currency))
            errors.Add($"{Parameters.Currency} applies only to {Modes.Money} mode.");
        if (mode == Modes.Number && parameters.ContainsKey(Parameters.IncludeMinor))
            errors.Add($"{Parameters.IncludeMinor} applies only to {Modes.Money} mode.");
        if (mode == Modes.Money && parameters.ContainsKey(Parameters.Gender))
            errors.Add($"{Parameters.Gender} applies only to {Modes.Number} mode.");
        if (parameters.TryGetValue(Parameters.Currency, out var currencyOperand) &&
            currencyOperand.Kind == FormatterOperandKind.Literal && currencyOperand.Literal is { } currency)
        {
            var explicitLanguage = parameters.TryGetValue(Parameters.Language, out var languageOperand) &&
                                   languageOperand.Kind == FormatterOperandKind.Literal &&
                                   languageOperand.Literal is { } language && service.SupportsLanguage(language)
                ? language
                : null;
            if (explicitLanguage is not null && !service.SupportsCurrency(explicitLanguage, currency) ||
                explicitLanguage is null && !service.SupportsCurrency(currency))
                errors.Add($"Unsupported spellout currency '{currency}'.");
        }
        return errors.Count == 0 ? FormatterValidationResult.Valid : new FormatterValidationResult(errors);
    }

    public string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
        FormatterContext context)
    {
        var error = Validate(parameters);
        if (error.Errors.Count != 0) throw new FormatException(string.Join("; ", error.Errors));
        var mode = Get(parameters, Parameters.Mode) ?? Modes.Number;
        var language = Get(parameters, Parameters.Language) ?? service.ResolveLanguage(context.Culture);
        var currency = Get(parameters, Parameters.Currency);
        var gender = Get(parameters, Parameters.Gender) ?? Genders.Masculine;
        var includeMinor = GetBoolean(parameters, Parameters.IncludeMinor, true);
        var capitalize = GetBoolean(parameters, Parameters.Capitalize, true);
        return service.Format(decimalParser.Parse(value), mode, language, currency, gender, includeMinor, capitalize);
    }

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
