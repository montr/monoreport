using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.Configuration;
using Xunit;

namespace MonoReport.Api.Tests;

public sealed class StaticHostingTests : IDisposable
{
    private readonly string _webRoot = Path.Combine(Path.GetTempPath(), $"monoreport-static-{Guid.NewGuid():N}");

    public StaticHostingTests()
    {
        Directory.CreateDirectory(Path.Combine(_webRoot, "assets"));
        File.WriteAllText(Path.Combine(_webRoot, "index.html"), "<!doctype html><title>MonoReport test UI</title>");
        File.WriteAllText(Path.Combine(_webRoot, "assets", "app.abc123.js"), "globalThis.monoreport = true;");
    }

    public void Dispose()
    {
        Directory.Delete(_webRoot, true);
    }

    [Fact]
    public async Task Serves_ui_assets_and_frontend_routes_with_correct_content_types()
    {
        using var factory = Factory(true);
        using var client = factory.CreateClient(new WebApplicationFactoryClientOptions {AllowAutoRedirect = false});

        using var root = await client.GetAsync("/");
        Assert.Equal(HttpStatusCode.OK, root.StatusCode);
        Assert.Equal("text/html", root.Content.Headers.ContentType?.MediaType);
        Assert.Contains("MonoReport test UI", await root.Content.ReadAsStringAsync());

        using var asset = await client.GetAsync("/assets/app.abc123.js");
        Assert.Equal(HttpStatusCode.OK, asset.StatusCode);
        Assert.Equal("text/javascript", asset.Content.Headers.ContentType?.MediaType);
        Assert.Contains("monoreport", await asset.Content.ReadAsStringAsync());

        using var route = await client.GetAsync("/future/report/42");
        Assert.Equal(HttpStatusCode.OK, route.StatusCode);
        Assert.Equal("text/html", route.Content.Headers.ContentType?.MediaType);
    }

    [Theory]
    [InlineData("/api/not-real")]
    [InlineData("/docs/not-real")]
    [InlineData("/openapi/not-real")]
    [InlineData("/missing.js")]
    public async Task Does_not_fall_back_for_backend_prefixes_or_file_paths(string path)
    {
        using var factory = Factory(true);
        using var client = factory.CreateClient(new WebApplicationFactoryClientOptions {AllowAutoRedirect = false});

        using var response = await client.GetAsync(path);

        Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
        Assert.NotEqual("text/html", response.Content.Headers.ContentType?.MediaType);
        Assert.DoesNotContain("MonoReport test UI", await response.Content.ReadAsStringAsync());
    }

    [Theory]
    [InlineData("/docs")]
    [InlineData("/docs/")]
    [InlineData("/openapi/v1.json")]
    public async Task Disabled_documentation_does_not_become_a_spa_route(string path)
    {
        using var factory = Factory(false);
        using var client = factory.CreateClient(new WebApplicationFactoryClientOptions {AllowAutoRedirect = false});

        using var response = await client.GetAsync(path);

        Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
        Assert.DoesNotContain("MonoReport test UI", await response.Content.ReadAsStringAsync());
    }

    private WebApplicationFactory<Program> Factory(bool documentationEnabled)
    {
        return new WebApplicationFactory<Program>().WithWebHostBuilder(builder =>
        {
            builder.UseWebRoot(_webRoot);
            builder.ConfigureAppConfiguration((_, configuration) => configuration.AddInMemoryCollection(
                new Dictionary<string, string?>
                    {["MonoReport:DocumentationEnabled"] = documentationEnabled.ToString()}));
        });
    }
}
