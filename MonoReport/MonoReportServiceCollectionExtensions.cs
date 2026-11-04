using System;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using MonoReport.Core.Formatting;
using MonoReport.Core.Templates;
using MonoReport.Docx;
using MonoReport.Docx.Planning;
using MonoReport.Xlsx;
using MonoReport.Xlsx.Planning;
using XlsxGeometryPlanner = MonoReport.Xlsx.Planning.XlsxGeometryPlanner;
using XlsxTemplatePlanner = MonoReport.Xlsx.Planning.XlsxTemplatePlanner;

namespace MonoReport;

/// <summary>Registers the MonoReport template engine.</summary>
public static class MonoReportServiceCollectionExtensions
{
    public static IServiceCollection AddMonoReport(this IServiceCollection services)
    {
        ArgumentNullException.ThrowIfNull(services);

        services.TryAddTransient<DocumentGenerator>();
        services.TryAddTransient<TemplateAnalyzer>();
        services.TryAddTransient<PackageFormatDetector>();
        services.TryAddTransient<DocxGenerator>();
        services.TryAddTransient<DocxTemplateAnalyzer>();
        services.TryAddTransient<XlsxGenerator>();
        services.TryAddTransient<XlsxTemplateAnalyzer>();
        services.TryAddTransient<XlsxPackageReader>();
        services.TryAddTransient<XlsxGeometryPlanner>();
        services.TryAddTransient<XlsxTemplatePlanner>();
        services.TryAddTransient<XlsxDocumentRenderer>();
        services.TryAddTransient<XlsxFormulaReferences>();
        services.TryAddTransient<XlsxFormulaPlanner>();
        services.TryAddTransient<DocxStoryReader>();
        services.TryAddTransient<TemplateLanguageParser>();
        services.TryAddTransient<DocxTemplatePlanner>();
        services.TryAddTransient<DocxRenderer>();
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