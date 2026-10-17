using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text.Json;

namespace MonoReport.Tests.Services;

public sealed class ApiHostTests
{
    [Fact]
    [Trait("Category", "RealHost")]
    public async Task Real_kestrel_enforces_whole_request_bound_including_streamed_requests_and_recovers()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0); listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port; listener.Stop();
        var root = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Parent!.Parent!.Parent!.FullName;
        var configuration = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Name;
        var dll = Path.Combine(root, "MonoReport.Api", "bin", configuration, "net10.0", "MonoReport.Api.dll");
        var start = new ProcessStartInfo("dotnet") { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        start.ArgumentList.Add(dll);
        start.Environment["ASPNETCORE_URLS"] = $"http://127.0.0.1:{port}";
        start.Environment["ASPNETCORE_ENVIRONMENT"] = "Development";
        start.Environment["MonoReport__MaxRequestBytes"] = "8192";
        using var process = Process.Start(start)!;
        var output = process.StandardOutput.ReadToEndAsync(); var errors = process.StandardError.ReadToEndAsync();
        try
        {
            using var client = new HttpClient { BaseAddress = new Uri($"http://127.0.0.1:{port}"), Timeout = TimeSpan.FromSeconds(15) };
            var ready = false;
            for (var attempt = 0; attempt < 100; attempt++)
            {
                if (process.HasExited) throw new InvalidOperationException(await errors);
                try { using var response = await client.GetAsync("/openapi/v1.json"); if (response.IsSuccessStatusCode) { ready = true; break; } }
                catch (HttpRequestException) { }
                await Task.Delay(100);
            }
            Assert.True(ready, "Local API host did not become ready.");
            foreach (var endpoint in new[] { "analyze", "generate" })
            {
                foreach (var streamed in new[] { false, true })
                {
                    using var multipart = HttpApiTests.Form(new byte[5000], endpoint == "generate" ? new string(' ', 5000) : null);
                    if (endpoint == "analyze") multipart.Add(new ByteArrayContent(new byte[5000]), "extra", "extra.docx");
                    using HttpContent content = streamed ? new StreamingContent(multipart) : multipart;
                    using var response = await client.PostAsync($"/api/documents/{endpoint}", content);
                    Assert.True(response.StatusCode == HttpStatusCode.RequestEntityTooLarge,
                        $"{endpoint}, streamed={streamed}: {(int)response.StatusCode} {await response.Content.ReadAsStringAsync()}");
                    if (response.Content.Headers.ContentType?.MediaType == "application/problem+json")
                    {
                        using var problem = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
                        Assert.Equal("request.too_large", problem.RootElement.GetProperty("code").GetString());
                    }
                }
                using var small = HttpApiTests.Form(Docx.Create(Docx.P(Docx.Run("{Name}"))), endpoint == "generate" ? "{\"Name\":\"fine\"}" : null);
                using var recovered = await client.PostAsync($"/api/documents/{endpoint}", small);
                Assert.Equal(HttpStatusCode.OK, recovered.StatusCode);
                if (endpoint == "generate")
                    Assert.Equal("fine", Docx.Text(Docx.Read(await recovered.Content.ReadAsByteArrayAsync())));
                else
                {
                    using var analysis = JsonDocument.Parse(await recovered.Content.ReadAsStringAsync());
                    Assert.Equal("Name", analysis.RootElement.GetProperty("requirements")[0].GetProperty("name").GetString());
                }
            }
        }
        finally
        {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync(); await output; await errors;
        }
    }

    private sealed class StreamingContent : HttpContent
    {
        private readonly HttpContent _inner;
        public StreamingContent(HttpContent inner)
        {
            _inner = inner;
            foreach (var header in inner.Headers) Headers.TryAddWithoutValidation(header.Key, header.Value);
        }
        protected override bool TryComputeLength(out long length) { length = 0; return false; }
        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) => _inner.CopyToAsync(stream);
    }
}
