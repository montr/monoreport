using System.Globalization;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Formatting;
using MonoReport.Cli.Tests;
using Xunit;

namespace MonoReport.Tests.Docx;

public sealed class DocxXmlSanitizationTests
{
    [Fact]
    public void Public_formatting_api_keeps_text_unchanged()
    {
        using var provider = new ServiceCollection().BuildServiceProvider();
        var formatter = new ValueFormatter(provider);
        using var data = JsonDocument.Parse("\"a\\u0001b\"");

        Assert.Equal("a\u0001b", formatter.Format(data.RootElement, null, null,
            new FormatterContext(CultureInfo.InvariantCulture)));
    }

    [Theory]
    [InlineData("{Name}", "acafé 😀z")]
    [InlineData("{Name upper}", "ACAFÉ 😀Z")]
    public void Replacements_remove_only_xml_10_characters_and_report_context(string expression, string expected)
    {
        var template = DocxFixture.Create(DocxFixture.P(DocxFixture.Run(expression)));
        var result = DocxFixture.Generate(template, """{"Name":"a\u0001café 😀\u000Bz\uFFFE"}""");

        Assert.Equal(expected, DocxFixture.Text(DocxFixture.Read(result.Document)));
        var warning = Assert.Single(result.Warnings);
        Assert.Equal(DiagnosticCodes.GenerationXmlCharactersRemoved, warning.Code);
        Assert.Equal(3, warning.RemovedCharacterCount);
        Assert.Equal(expression, warning.Field);
        Assert.Equal(expression, warning.FieldCode);
        Assert.Equal("$.Name", warning.DataPath);
        Assert.Equal("/word/document.xml, paragraph 1", warning.Location);
        Assert.Equal("Removed 3 prohibited XML 1.0 characters.", warning.Message);
    }
}
