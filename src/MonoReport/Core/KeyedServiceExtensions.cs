using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using Microsoft.Extensions.DependencyInjection;

namespace MonoReport.Core;

public static class KeyedServiceExtensions
{
	public static IServiceCollection AddServiceKeyProvider(this IServiceCollection services)
	{
		return services
			.AddSingleton(services)
			.AddTransient(typeof(KeyedServiceProvider<>))
			.AddSingleton(typeof(KeyedServiceKeyProvider<,>));
	}
}
	
public sealed class KeyedServiceProvider<TService>(IServiceProvider serviceProvider) 
	where TService : notnull
{
	public TService? GetKeyedService(object serviceKey)
	{
		return serviceProvider.GetKeyedService<TService>(serviceKey);
	}

	public TService GetRequiredKeyedService(object serviceKey)
	{
		return serviceProvider.GetRequiredKeyedService<TService>(serviceKey);
	}
}
	
public sealed class KeyedServiceKeyProvider<TKey, TService>(IServiceCollection services) 
	where TKey : notnull 
	where TService : notnull
{
	public IReadOnlyList<TKey> Keys => (
		from service in services
		where service.ServiceKey?.GetType() == typeof(TKey)
		where service.ServiceType == typeof(TService)
		select (TKey)service.ServiceKey!
	).Distinct().ToImmutableList();
}
