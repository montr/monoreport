
using System;
using System.Collections.Generic;

namespace MonoReport.Core.Formatting;

internal static class PatternValidation
{
    internal static FormatterValidationResult Validate(IReadOnlyDictionary<string, FormatterOperand> parameters,
        string name, Func<string, string> format)
    {
        if (!parameters.TryGetValue("pattern", out var operand) || operand.Kind != FormatterOperandKind.Literal)
            return FormatterValidationResult.Valid;
        if (operand.Literal!.Length == 0)
            return new FormatterValidationResult([$"Formatter '{name}' requires a non-empty pattern."]);
        try
        {
            _ = format(operand.Literal);
            return FormatterValidationResult.Valid;
        }
        catch (FormatException)
        {
            return new FormatterValidationResult([$"Formatter '{name}' has an invalid .NET format pattern."]);
        }
    }
}