using System.Collections.Generic;
using System.Text.Json;
using MonoReport.Core.Diagnostics;

namespace MonoReport.Api.Models;

public sealed record SampleJsonResponse(JsonElement Data, IReadOnlyList<GenerationDiagnostic> Warnings);
