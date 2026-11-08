using System;
using System.Collections.Generic;
using System.Linq;

namespace MonoReport.Core.Formatting;

public sealed class FormatterParameterBinder
{
    public FormatterValidationResult Bind(string name, IFormatter formatter,
        IEnumerable<KeyValuePair<string, FormatterOperand>> supplied,
        out IReadOnlyDictionary<string, FormatterOperand> parameters)
    {
        var metadata = formatter.Metadata;
        var known = metadata.Parameters.ToDictionary(parameter => parameter.Name, StringComparer.Ordinal);
        var bound = new Dictionary<string, FormatterOperand>(StringComparer.Ordinal);
        parameters = bound;
        foreach (var (parameter, value) in supplied)
        {
            if (!known.ContainsKey(parameter))
                return new FormatterValidationResult([$"Formatter '{name}' does not declare parameter '{parameter}'."]);
            if (!bound.TryAdd(parameter, value))
                return new FormatterValidationResult([$"Formatter '{name}' parameter '{parameter}' was supplied more than once."]);
        }

        var missing = metadata.Parameters.FirstOrDefault(parameter =>
            parameter.Required && !bound.ContainsKey(parameter.Name));
        if (missing is not null)
            return new FormatterValidationResult([$"Formatter '{name}' requires parameter '{missing.Name}'."]);

        return formatter.Validate(bound);
    }
}
