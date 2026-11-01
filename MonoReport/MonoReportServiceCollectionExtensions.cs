using Microsoft.Extensions.DependencyInjection.Extensions;
using MonoReport.Services;
using MonoReport.Services.Xlsx;
using MonoReport.Models;

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
        services.TryAddTransient<XlsxFormulaReferences>();
        services.TryAddTransient<XlsxFormulaPlanner>();
        services.TryAddTransient<DocxStoryReader>();
        services.TryAddTransient<TemplateLanguageParser>();
        services.TryAddTransient<DocxTemplatePlanner>();
        services.TryAddTransient<DocumentRenderer>();
        services.TryAddTransient<DataResolver>();
        services.TryAddTransient<ValueFormatter>();
        services.TryAddSingleton<InvariantDecimalParser>();
        services.TryAdd(ServiceDescriptor.KeyedSingleton<IFormatter, UpperFormatter>("upper"));
        services.TryAdd(ServiceDescriptor.KeyedSingleton<IFormatter, LowerFormatter>("lower"));
        services.TryAdd(ServiceDescriptor.KeyedSingleton<IFormatter, CapsFormatter>("caps"));
        services.TryAdd(ServiceDescriptor.KeyedSingleton<IFormatter, FirstCapFormatter>("firstcap"));
        services.TryAdd(ServiceDescriptor.KeyedSingleton<IFormatter, NumberFormatter>("number"));
        services.TryAdd(ServiceDescriptor.KeyedSingleton<IFormatter, DateFormatter>("date"));

        return services;
    }
}
