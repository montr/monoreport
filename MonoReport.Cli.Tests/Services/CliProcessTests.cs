using System.Diagnostics;
using System.Text.Json;

namespace MonoReport.Tests.Services;

public sealed class CliProcessTests
{
    [Fact]
    public async Task Real_entry_point_analyzes_and_generates_unified_templates()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"monoreport-cli-process-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var template = Path.Combine(directory, "template.docx");
        var result = Path.Combine(directory, "result.docx");
        await File.WriteAllBytesAsync(template, Docx.Create(Docx.P(Docx.Run("Hello {Name}!"))));
        try
        {
            var analysis = await Run(["--analyze", "--template", template]);
            Assert.Equal(0, analysis.ExitCode);
            Assert.Empty(analysis.Error);
            using (var json = JsonDocument.Parse(analysis.Output))
                Assert.Equal("Name", json.RootElement.GetProperty("requirements")[0].GetProperty("name").GetString());

            var generation = await Run(["--template", template, "--json", "{\"Name\":\"Ada\"}", "--output", result]);
            Assert.Equal(0, generation.ExitCode);
            Assert.Empty(generation.Error);
            Assert.Equal("Hello Ada!", Docx.Text(Docx.Read(await File.ReadAllBytesAsync(result))));
        }
        finally { Directory.Delete(directory, true); }
    }

    private static async Task<(int ExitCode, byte[] Output, string Error)> Run(string[] arguments)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Parent!.Parent!.Parent!.FullName;
        var configuration = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Name;
        var start = new ProcessStartInfo("dotnet") { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        start.ArgumentList.Add(Path.Combine(root, "MonoReport.Cli", "bin", configuration, "net10.0", "MonoReport.Cli.dll"));
        foreach (var argument in arguments) start.ArgumentList.Add(argument);
        using var process = Process.Start(start)!;
        using var output = new MemoryStream();
        var copy = process.StandardOutput.BaseStream.CopyToAsync(output);
        var error = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
        await copy;
        return (process.ExitCode, output.ToArray(), await error);
    }
}
