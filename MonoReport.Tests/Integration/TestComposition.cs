using Microsoft.Extensions.DependencyInjection;
using MonoReport.Core.Formatting;
using MonoReport.Core.Templates;
using MonoReport.Formatter.Spellout;

namespace MonoReport.Tests.Integration;

internal static class TestComposition
{
    private static readonly ServiceProvider Provider =
        new ServiceCollection().AddMonoReport().AddMonoReportSpelloutFormatter().BuildServiceProvider();

    internal static TemplateLanguageParser Parser => Provider.GetRequiredService<TemplateLanguageParser>();
    internal static ValueFormatter Formatter => Provider.GetRequiredService<ValueFormatter>();
    internal static TemplateAnalyzer Analyzer => Provider.GetRequiredService<TemplateAnalyzer>();
}