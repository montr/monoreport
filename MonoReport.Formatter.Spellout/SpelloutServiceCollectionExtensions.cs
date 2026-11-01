using System.Text.RegularExpressions;
using Microsoft.Extensions.DependencyInjection.Extensions;
using MonoReport.Formatter.Spellout;
using MonoReport.Models;
using MonoReport.Services;

namespace Microsoft.Extensions.DependencyInjection;

public static partial class SpelloutServiceCollectionExtensions
{
    public static IServiceCollection AddMonoReportSpelloutFormatter(this IServiceCollection services, string key = "spellout")
    {
        ArgumentNullException.ThrowIfNull(services);
        ArgumentException.ThrowIfNullOrWhiteSpace(key);
        if (!FormatterKeyRegex().IsMatch(key))
            throw new ArgumentException("Formatter keys must start with a Unicode letter or underscore and contain only Unicode letters, digits, or underscores.", nameof(key));

        services.TryAddSingleton<InvariantDecimalParser>();
        services.TryAddSingleton<RubSpellOutFormatter>();
        services.AddKeyedSingleton<IFormatter, SpelloutFormatter>(key);
        return services;
    }

    [GeneratedRegex(@"^[\p{L}_][\p{L}\p{Nd}_]*$", RegexOptions.CultureInvariant)]
    private static partial Regex FormatterKeyRegex();
}
