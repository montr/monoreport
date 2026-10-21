using System.Text.Json;

namespace MonoReport.Models;

// Current value/path first, followed by logical ancestors; neither document format owns this input.
internal sealed record DataScope(JsonElement Root, IReadOnlyList<JsonElement> Values, IReadOnlyList<string> Paths);
