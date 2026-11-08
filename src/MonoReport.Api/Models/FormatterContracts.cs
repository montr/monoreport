using System.Collections.Generic;
using System.Text.Json;

namespace MonoReport.Api.Models;

public sealed record FormatterDescription(
    string Name,
    IReadOnlyList<FormatterParameterDescription> Parameters,
    string? PositionalParameter);

public sealed record FormatterParameterDescription(string Name, string Type, bool Required);

public sealed record FormatterExecutionRequest(
    JsonElement Value,
    Dictionary<string, string>? Parameters = null,
    string? Culture = null);

public sealed record FormatterExecutionResult(string Result);
