using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Core.Templates;
using Xunit;
using System.IO;

namespace MonoReport.Formatter.Spellout.Tests;

public sealed class SpelloutBehaviorTests
{
    [Theory]
    [InlineData("0", "Ноль рублей 00 копеек")]
    [InlineData("-1.005", "Минус один рубль 01 копейка")]
    [InlineData("2.5", "Два рубля 50 копеек")]
    [InlineData("5.05", "Пять рублей 05 копеек")]
    [InlineData("11.11", "Одиннадцать рублей 11 копеек")]
    [InlineData("21.21", "Двадцать один рубль 21 копейка")]
    [InlineData("22.22", "Двадцать два рубля 22 копейки")]
    [InlineData("25.25", "Двадцать пять рублей 25 копеек")]
    [InlineData("1000", "Одна тысяча рублей 00 копеек")]
    [InlineData("2000", "Две тысячи рублей 00 копеек")]
    [InlineData("11000", "Одиннадцать тысяч рублей 00 копеек")]
    [InlineData("21000", "Двадцать одна тысяча рублей 00 копеек")]
    public void Formats_rubles_with_rounding_and_russian_grammatical_forms(string value, string expected)
    {
        Assert.Equal(expected, Render(value, "tr-TR"));
    }

    [Fact]
    public void Accepts_the_upper_boundary_and_rejects_the_first_rounded_out_of_range_value()
    {
        Assert.EndsWith("99 копеек", Render("999999999999999999.994"));
        var error = Assert.Throws<TemplateException>(() => Render("999999999999999999.995"));
        Assert.Equal("generation.invalid_format", error.Code);
    }

    [Fact]
    public void Unsupported_json_values_return_the_generation_diagnostic()
    {
        var error = Assert.Throws<TemplateException>(() => Render("true"));
        Assert.Equal("generation.invalid_format", error.Code);
        Assert.Equal("spellout", Assert.Single(error.Errors).Formatter);
    }

    private static string Render(string value, string? culture = null)
    {
        using var provider = new ServiceCollection().AddMonoReport().AddMonoReportSpelloutFormatter()
            .BuildServiceProvider();
        using var input = new MemoryStream(CreateDocx("{Value spellout}"));
        using var data = JsonDocument.Parse("{\"Value\":" + value + "}");
        var result = provider.GetRequiredService<DocumentGenerator>().Generate(input, data,
            new GenerationOptions {Culture = CultureInfo.GetCultureInfo(culture ?? "en-US")});
        using var documentStream = new MemoryStream(result.Document);
        using var document = WordprocessingDocument.Open(documentStream, false);
        return document.MainDocumentPart!.Document!.Body!.InnerText;
    }

    internal static byte[] CreateDocx(string expression)
    {
        using var output = new MemoryStream();
        using (var document = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text(expression)))));
        }

        return output.ToArray();
    }
}