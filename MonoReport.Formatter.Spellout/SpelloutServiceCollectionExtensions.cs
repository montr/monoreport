using System;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using MonoReport.Core.Formatting;
using MonoReport.Core.Templates;

namespace MonoReport.Formatter.Spellout;

public static class SpelloutServiceCollectionExtensions
{
    public static IServiceCollection AddMonoReportSpelloutFormatter(this IServiceCollection services, string key = "spellout")
    {
        ArgumentNullException.ThrowIfNull(services);
        ArgumentException.ThrowIfNullOrWhiteSpace(key);
        if (!TemplateIdentifier.IsValid(key))
            throw new ArgumentException(
                "Formatter keys must start with a Unicode letter or underscore and contain only Unicode letters, digits, or underscores.",
                nameof(key));

        services.TryAddSingleton<InvariantDecimalParser>();
        services.TryAddSingleton<SpelloutLanguageReader>();
        services.TryAddSingleton<SpelloutService>();
        services.AddKeyedSingleton<IFormatter, SpelloutFormatter>(key);
        return services;
    }
}
