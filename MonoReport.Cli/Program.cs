using Microsoft.Extensions.DependencyInjection;
using MonoReport.Formatter.Spellout;
using System;

namespace MonoReport.Cli;

internal static class Program
{
    private static int Main(string[] args)
    {
        var services = new ServiceCollection();
        services.AddMonoReport();
        services.AddMonoReportSpelloutFormatter();
        services.AddTransient<CliOptionsParser>();
        services.AddTransient<CliApplication>();
        using var provider = services.BuildServiceProvider();
        var application = provider.GetRequiredService<CliApplication>();
        return application.Run(args, Console.OpenStandardInput(), Console.OpenStandardOutput(), Console.Error);
    }
}