using Microsoft.Extensions.DependencyInjection.Extensions;
using MonoReport.Services;

namespace Microsoft.Extensions.DependencyInjection;

/// <summary>Registers the MonoReport template engine.</summary>
public static class MonoReportServiceCollectionExtensions
{
    public static IServiceCollection AddMonoReport(this IServiceCollection services)
    {
        ArgumentNullException.ThrowIfNull(services);

        services.TryAddTransient<MergeEngine>();
        services.TryAddTransient<TemplateAnalyzer>();
        services.TryAddTransient<DocxStoryReader>();
        services.TryAddTransient<TemplateLanguageParser>();
        services.TryAddTransient<DocxTemplatePlanner>();
        services.TryAddTransient<DocumentRenderer>();
        services.TryAddTransient<DataResolver>();
        services.TryAddTransient<ValueFormatter>();
        services.TryAddSingleton<RubSpellOutFormatter>();

        return services;
    }
}
