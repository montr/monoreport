using MonoReport.Models;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature012TransportTests
{
    [Fact]
    public void Test_owned_xlsx_generates_through_public_core()
    {
        var fixture = new XlsxTransportFixture(); var source = fixture.Source;
        using var services = new XlsxFixture().Services();
        using var input = new MemoryStream(source); using var data = System.Text.Json.JsonDocument.Parse(fixture.Data);
        var engine = services.GetRequiredService<MergeEngine>();
        var result = engine.Merge(input, data); Assert.Empty(result.Warnings);
        fixture.AssertGenerated(result.Document); Assert.Equal(source, fixture.Source);
    }

    [Theory]
    [InlineData(false, 0)] [InlineData(false, 1)] [InlineData(false, 3)]
    [InlineData(true, 0)] [InlineData(true, 1)] [InlineData(true, 3)]
    public void Core_success_retains_original_order_and_context_across_containers(bool spreadsheet, int count)
    {
        var fixture = new TransportFixture(); var source = fixture.Create(spreadsheet, count);
        var result = fixture.Merge(source);
        Assert.Equal(count, result.Warnings.Count);
        Assert.Equal(spreadsheet ? MonoReport.Models.DocumentFormat.Xlsx : MonoReport.Models.DocumentFormat.Docx, result.Format);
        Assert.Equal(new[] { "First", "Second", "Third" }.Take(count), result.Warnings.Select(warning =>
            new[] { "First", "Second", "Third" }.Single(name => warning.Field.Contains(name, StringComparison.Ordinal))));
        Assert.All(result.Warnings, warning =>
        {
            Assert.NotEmpty(warning.Message); Assert.NotNull(warning.FieldCode); Assert.NotNull(warning.Location);
            Assert.NotNull(warning.Part); Assert.NotNull(warning.DataPath);
            Assert.Equal(spreadsheet ? DiagnosticCodes.XlsxMissingValue : null, warning.Code);
        });
        if (count == 3) Assert.NotEqual(result.Warnings[0].Part, result.Warnings[2].Part);
        fixture.AssertPackage(result.Document, spreadsheet);
    }
}
