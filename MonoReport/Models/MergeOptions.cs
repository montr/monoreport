using System.Globalization;

namespace MonoReport.Models
{
    public sealed class MergeOptions
    {
        public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;

        public MissingValueBehavior MissingValues { get; init; } = MissingValueBehavior.EmptyString;
    }
}
