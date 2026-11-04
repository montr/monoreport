using System;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Core;
using Xunit;

namespace MonoReport.Tests.Core;

public sealed class KeyedServiceProviderTests
{
    [Fact]
    public void Key_discovery_filters_service_and_key_types_and_deduplicates_replacements()
    {
        var services = new ServiceCollection().AddServiceKeyProvider();
        services.AddKeyedSingleton<ProbeService>("probe");
        services.AddKeyedSingleton<ProbeService>("probe");
        services.AddKeyedSingleton<ProbeService>(1);
        services.AddKeyedSingleton<object>("other");
        services.AddSingleton<ProbeService>();
        using var provider = services.BuildServiceProvider();

        var keys = provider.GetRequiredService<KeyedServiceKeyProvider<string, ProbeService>>();

        Assert.Equal(new[] {"probe"}, keys.Keys);
    }

    [Fact]
    public void Registered_provider_resolves_scoped_services_in_the_caller_scope_and_preserves_disposal()
    {
        var services = new ServiceCollection().AddServiceKeyProvider();
        services.AddKeyedScoped<ProbeService>("probe");
        using var provider = services.BuildServiceProvider(new ServiceProviderOptions
        {
            ValidateOnBuild = true,
            ValidateScopes = true
        });

        ProbeService first;
        using (var scope = provider.CreateScope())
        {
            var keyed = scope.ServiceProvider.GetRequiredService<KeyedServiceProvider<ProbeService>>();
            first = keyed.GetRequiredKeyedService("probe");
            Assert.Same(first, keyed.GetKeyedService("probe"));
            Assert.False(first.Disposed);
        }

        Assert.True(first.Disposed);
        using var secondScope = provider.CreateScope();
        var second = secondScope.ServiceProvider.GetRequiredService<KeyedServiceProvider<ProbeService>>()
            .GetRequiredKeyedService("probe");
        Assert.NotSame(first, second);
        Assert.False(second.Disposed);
    }

    private sealed class ProbeService : IDisposable
    {
        public bool Disposed { get; private set; }

        public void Dispose()
        {
            Disposed = true;
        }
    }
}
