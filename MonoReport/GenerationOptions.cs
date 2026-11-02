using System.Globalization;

namespace MonoReport;

public sealed class GenerationOptions
{
    public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;

    public MissingValueBehavior MissingValues { get; init; } = MissingValueBehavior.EmptyString;
}