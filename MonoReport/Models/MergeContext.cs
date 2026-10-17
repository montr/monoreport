using System.Text.Json;

namespace MonoReport.Models
{
    public sealed class MergeContext(JsonElement root, MergeOptions options)
    {
        internal const int ErrorLimit = 100;

        public Stack<JsonElement> Values { get; } = new([root]);

        internal JsonElement Root { get; } = root;

        internal Stack<string> DataPaths { get; } = new(["$"]);

        public MergeOptions Options { get; } = options;

        public List<MergeDiagnostic> Warnings { get; } = new();

        internal List<GenerationDiagnostic> Errors { get; } = new();

        internal TemplateException? FirstError { get; private set; }

        internal bool LimitReached { get; private set; }

        internal HashSet<string> WarningKeys { get; } = new(StringComparer.Ordinal);

        internal bool Record(TemplateException error, string dataPath, string? location = null)
        {
            if (LimitReached) return false;
            FirstError ??= error;
            Errors.Add(error.ToDiagnostic(dataPath, location));
            if (Errors.Count == ErrorLimit) LimitReached = true;
            return !LimitReached;
        }
    }
}
