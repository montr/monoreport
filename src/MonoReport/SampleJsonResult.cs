using System;
using System.Collections.Generic;
using System.Text.Json;
using MonoReport.Core.Diagnostics;

namespace MonoReport;

public sealed class SampleJsonResult
{
    internal SampleJsonResult(JsonElement data, IEnumerable<GenerationDiagnostic> warnings)
    {
        Data = data.Clone();
        Warnings = Array.AsReadOnly(new List<GenerationDiagnostic>(warnings).ToArray());
    }

    public JsonElement Data { get; }
    public IReadOnlyList<GenerationDiagnostic> Warnings { get; }
}
