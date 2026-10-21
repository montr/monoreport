using Microsoft.Extensions.DependencyInjection.Extensions;
using MonoReport.Services;
using MonoReport.Services.Xlsx;

namespace Microsoft.Extensions.DependencyInjection;

/// <summary>Registers the MonoReport template engine.</summary>
public static class MonoReportServiceCollectionExtensions
{
    public static IServiceCollection AddMonoReport(this IServiceCollection services)
    {
        ArgumentNullException.ThrowIfNull(services);

        services.TryAddTransient(provider => new MergeEngine(provider.GetRequiredService<DocumentRenderer>(),
            provider.GetRequiredService<DocxStoryReader>(), provider.GetRequiredService<TemplateAnalyzer>(),
            provider.GetRequiredService<PackageFormatDetector>(), provider.GetRequiredService<XlsxTemplateHandler>()));
        services.TryAddTransient(provider => new TemplateAnalyzer(provider.GetRequiredService<DocxStoryReader>(),
            provider.GetRequiredService<DocxTemplatePlanner>(), provider.GetRequiredService<PackageFormatDetector>(),
            provider.GetRequiredService<XlsxTemplateHandler>()));
        services.TryAddTransient<PackageFormatDetector>();
        services.TryAddTransient<XlsxTemplateHandler>();
        services.TryAddTransient<XlsxPackageReader>();
        services.TryAddTransient<XlsxGeometryPlanner>();
        services.TryAddTransient<XlsxTemplatePlanner>();
        services.TryAddTransient<XlsxDocumentRenderer>();
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
