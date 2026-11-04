using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using MonoReport.Core.Formatting;
using MonoReport.Core.Templates;

namespace MonoReport.Formatter.Spellout;

public static class SpelloutServiceCollectionExtensions
{
    public static IServiceCollection AddMonoReportSpelloutFormatter(this IServiceCollection services,
        string key = "spellout")
    {
        ArgumentNullException.ThrowIfNull(services);
        ArgumentException.ThrowIfNullOrWhiteSpace(key);
        if (!TemplateIdentifier.IsValid(key))
            throw new ArgumentException(
                "Formatter keys must start with a Unicode letter or underscore and contain only Unicode letters, digits, or underscores.",
                nameof(key));

        services.TryAddSingleton<InvariantDecimalParser>();
        services.TryAddSingleton<SpelloutLanguageReader>();
        services.TryAddSingleton<SpelloutService>(provider =>
        {
            var reader = provider.GetRequiredService<SpelloutLanguageReader>();
            var assembly = typeof(SpelloutServiceCollectionExtensions).Assembly;
            var languages = new Dictionary<string, SpelloutLanguage>(StringComparer.Ordinal);
            foreach (var code in new[] { "ru", "en" })
            {
                var suffix = $"spell-{code}.json";
                var resource = assembly.GetManifestResourceNames()
                    .SingleOrDefault(name => name.EndsWith(suffix, StringComparison.Ordinal));
                if (resource is null)
                    throw new InvalidOperationException($"Missing embedded spellout resource: {suffix}.");
                using var stream = assembly.GetManifestResourceStream(resource)
                                   ?? throw new InvalidOperationException($"Unable to load spellout resource: {suffix}.");
                var language = reader.Read(stream);
                if (language.Language != code)
                    throw new InvalidOperationException($"Unexpected language in {suffix}.");
                languages.Add(code, language);
            }

            return new SpelloutService(languages);
        });
        services.AddKeyedSingleton<IFormatter, SpelloutFormatter>(key);
        return services;
    }
}
