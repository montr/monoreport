using System.Diagnostics;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc.Testing;

namespace MonoReport.Tests.Services;

public sealed class ValidationContractParityTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Cli_and_http_return_identical_analysis(bool valid)
    {
        var bytes = valid ? Docx.Create(Docx.P(Docx.Run("{Name}"))) : Docx.Create(Docx.P(Docx.Run("{EndIf}")));
        var path = Path.Combine(Path.GetTempPath(), $"monoreport-parity-{Guid.NewGuid():N}.docx");
        await File.WriteAllBytesAsync(path, bytes);
        try
        {
            var root = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Parent!.Parent!.Parent!.FullName;
            var configuration = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Name;
            var start = new ProcessStartInfo("dotnet") { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
            start.ArgumentList.Add(Path.Combine(root, "MonoReport.Cli", "bin", configuration, "net10.0", "MonoReport.Cli.dll"));
            foreach (var argument in new[] { "--analyze", "--template", path }) start.ArgumentList.Add(argument);
            using var process = Process.Start(start)!;
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
            Assert.Empty(await stderr);
            Assert.Equal(valid ? 0 : 1, process.ExitCode);
            using var cli = JsonDocument.Parse(await stdout);

            using var factory = new WebApplicationFactory<Program>();
            using var client = factory.CreateClient();
            using var form = HttpApiTests.Form(bytes);
            using var response = await client.PostAsync("/api/documents/analyze", form);
            response.EnsureSuccessStatusCode();
            using var http = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
            Assert.True(JsonElement.DeepEquals(cli.RootElement, http.RootElement));
        }
        finally { File.Delete(path); }
    }
}
