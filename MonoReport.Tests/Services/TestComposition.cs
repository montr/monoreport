using Microsoft.Extensions.DependencyInjection;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

internal static class TestComposition
{
    private static readonly ServiceProvider Provider = new ServiceCollection().AddMonoReport().AddMonoReportSpelloutFormatter().BuildServiceProvider();
    internal static TemplateLanguageParser Parser => Provider.GetRequiredService<TemplateLanguageParser>();
    internal static ValueFormatter Formatter => Provider.GetRequiredService<ValueFormatter>();
}
