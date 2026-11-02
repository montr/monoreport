using Microsoft.Extensions.DependencyInjection;
using MonoReport.Cli.Tests;
using MonoReport.Core.Templates;
using Xunit;
using System;
using System.IO;

namespace MonoReport.Tests.Core;

public sealed class SpelloutCoreOptInTests
{
    [Fact]
    public void Core_only_composition_leaves_spellout_unregistered()
    {
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        using var input = new MemoryStream(DocxFixture.Create(DocxFixture.P(DocxFixture.Run("{Value spellout}"))));
        var result = provider.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(result.Valid);
        var parse = provider.GetRequiredService<TemplateLanguageParser>().Parse("{Value spellout}", "test");
        Assert.Contains("Formatter 'spellout' is not registered", Assert.Single(parse.Diagnostics).Detail,
            StringComparison.Ordinal);
    }
}