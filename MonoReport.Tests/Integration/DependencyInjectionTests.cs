using Microsoft.Extensions.DependencyInjection;
using Xunit;
using System.Linq;

namespace MonoReport.Tests.Integration;

public sealed class DependencyInjectionTests
{
    [Fact]
    public void AddMonoReport_resolves_every_registration_with_scope_and_build_validation_enabled()
    {
        var services = new ServiceCollection().AddMonoReport();
        var registrations = services.ToArray();
        using var provider = services.BuildServiceProvider(new ServiceProviderOptions
        {
            ValidateOnBuild = true,
            ValidateScopes = true
        });

        Assert.IsType<DocumentGenerator>(provider.GetRequiredService<DocumentGenerator>());
        Assert.IsType<TemplateAnalyzer>(provider.GetRequiredService<TemplateAnalyzer>());

        foreach (var registration in registrations)
        {
            if (registration.IsKeyedService)
                provider.GetRequiredKeyedService(registration.ServiceType, registration.ServiceKey);
            else
                provider.GetRequiredService(registration.ServiceType);
        }
    }
}
