using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;

namespace MonoReport.Core.Formatting;

public interface IFormatter
{
    FormatterMetadata Metadata { get; }
    FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters);

    string Format(JsonElement value, IReadOnlyDictionary<string, FormatterOperand> parameters,
        FormatterContext context);
}

public sealed record FormatterMetadata(
    IReadOnlyList<FormatterParameter> Parameters,
    string? PositionalParameter = null);

public sealed record FormatterParameter(string Name, Type Type, bool Required);

public enum FormatterOperandKind
{
    Literal,
    Reference
}

public sealed record FormatterOperand(FormatterOperandKind Kind, string? Literal);

public sealed record FormatterContext(CultureInfo Culture);

public sealed record FormatterValidationResult(IReadOnlyList<string> Errors)
{
    public static FormatterValidationResult Valid { get; } = new([]);
}