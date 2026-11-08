using System.Numerics;

namespace MonoReport;

public sealed record SampleJsonOptions
{
    public const int DefaultItemsPerCollection = 10;

    public BigInteger ItemsPerCollection { get; init; } = new(DefaultItemsPerCollection);
    public SampleJsonValues Values { get; init; } = SampleJsonValues.Filled;
}

public enum SampleJsonValues
{
    Filled,
    Empty
}
