using Microsoft.Extensions.DependencyInjection;
using MonoReport.Cli.Services;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature012CliTransportTests
{
    [Fact]
    public void Test_owned_xlsx_generates_through_public_cli()
    {
        var fixture = new XlsxTransportFixture();
        using var services = new XlsxFixture().Services();
        var app = new CliApplication(new CliOptionsParser(), services.GetRequiredService<MergeEngine>(), services.GetRequiredService<TemplateAnalyzer>());
        using var output = new DocumentOutputStream(); using var errors = new StringWriter();
        var path = Path.Combine(Path.GetTempPath(), "xlsx-transport-" + Guid.NewGuid().ToString("N"));
        File.WriteAllBytes(path, fixture.Source);
        try
        {
            Assert.Equal(0, app.Run(["--template", path, "--json", fixture.Data], Stream.Null, output, errors));
            Assert.Empty(errors.ToString());
            Assert.Equal(Assert.Single(output.Writes), output.ToArray());
            fixture.AssertGenerated(output.ToArray());
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false, 0)] [InlineData(false, 1)] [InlineData(false, 3)]
    [InlineData(true, 0)] [InlineData(true, 1)] [InlineData(true, 3)]
    public void Successful_warnings_match_core_in_order_and_keep_stdout_and_files_binary(bool spreadsheet, int count)
    {
        var fixture = new TransportFixture(); var source = fixture.Create(spreadsheet, count);
        var core = fixture.Merge(source); Assert.Equal(count, core.Warnings.Count);
        var path = Path.Combine(Path.GetTempPath(), "m5-" + Guid.NewGuid().ToString("N"));
        var destination = path + ".output"; File.WriteAllBytes(path, source);
        try
        {
            using var services = new XlsxFixture().Services();
            var app = new CliApplication(new CliOptionsParser(), services.GetRequiredService<MergeEngine>(), services.GetRequiredService<TemplateAnalyzer>());
            foreach (var toFile in new[] { false, true })
            {
                using var output = new DocumentOutputStream(); using var errors = new StringWriter();
                string[] args = ["--template", path, "--json", TransportFixture.Data];
                if (toFile) args = [.. args, "--output", destination];
                Assert.Equal(0, app.Run(args, Stream.Null, output, errors));
                var lines = errors.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
                Assert.Equal(count, lines.Length);
                for (var i = 0; i < count; i++)
                {
                    var warning = core.Warnings[i]; var line = lines[i];
                    Assert.StartsWith("Warning: ", line); Assert.Contains(warning.Message, line);
                    foreach (var pair in new[] { ("field", warning.Field), ("fieldCode", warning.FieldCode), ("part", warning.Part), ("location", warning.Location), ("dataPath", warning.DataPath) })
                        if (pair.Item2 is not null) Assert.Contains(pair.Item1 + "=" + pair.Item2, line);
                    if (warning.Code is not null) Assert.Contains(warning.Code + ": ", line);
                }
                if (toFile) Assert.Empty(output.ToArray());
                else Assert.Equal(Assert.Single(output.Writes), output.ToArray());
                var bytes = toFile ? File.ReadAllBytes(destination) : output.ToArray();
                Assert.Equal(core.Document.Length, bytes.Length);
                fixture.AssertPackage(bytes, spreadsheet);
            }
            File.WriteAllText(destination, "sentinel");
            using var failedOutput = new MemoryStream(); using var failedErrors = new StringWriter();
            Assert.Equal(1, app.Run(["--template", path, "--json", "{", "--output", destination], Stream.Null, failedOutput, failedErrors));
            Assert.Empty(failedOutput.ToArray()); Assert.Equal("sentinel", File.ReadAllText(destination));
            Assert.Equal(source, File.ReadAllBytes(path));
        }
        finally { File.Delete(path); File.Delete(destination); }
    }

    // Record the generated document supplied to stdout, without relying on ZIP timestamps from another merge.
    private sealed class DocumentOutputStream : MemoryStream
    {
        internal List<byte[]> Writes { get; } = [];

        public override void Write(byte[] buffer, int offset, int count)
        {
            Writes.Add(buffer.AsSpan(offset, count).ToArray());
            base.Write(buffer, offset, count);
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            var bytes = buffer.ToArray();
            Writes.Add(bytes);
            base.Write(bytes, 0, bytes.Length);
        }
    }
}
