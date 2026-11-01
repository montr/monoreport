using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class FormatterSpelloutRegistrationTests
{
    [Fact]
    public void Core_only_composition_leaves_spellout_unregistered()
    {
        using var provider = new ServiceCollection().AddMonoReport().BuildServiceProvider();
        Assert.Null(provider.GetKeyedService<IFormatter>("spellout"));
        using var input = new MemoryStream(Docx.Create(Docx.P(Docx.Run("{Value spellout}"))));
        var result = provider.GetRequiredService<TemplateAnalyzer>().Analyze(input);
        Assert.False(result.Valid);
        var parse = provider.GetRequiredService<TemplateLanguageParser>().Parse("{Value spellout}", "test");
        Assert.Contains("Formatter 'spellout' is not registered", Assert.Single(parse.Diagnostics).Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void Default_and_custom_keys_run_in_docx_and_xlsx_and_keep_invoked_key_diagnostics()
    {
        using var provider = new ServiceCollection().AddMonoReport()
            .AddMonoReportSpelloutFormatter()
            .AddMonoReportSpelloutFormatter("rub_words")
            .AddMonoReportSpelloutFormatter("amount_words")
            .BuildServiceProvider();

        foreach (var key in new[] { "spellout", "rub_words", "amount_words" })
        {
            var docx = Docx.Create(Docx.P(Docx.Run("{Value " + key + "}")));
            using var docxInput = new MemoryStream(docx);
            Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(docxInput).Valid);
            Assert.Equal("Один рубль 00 копеек", Docx.Text(Docx.Read(Merge(provider, docx, """{"Value":1}"""))));

            var fixture = new XlsxFixture();
            var xlsx = fixture.Create((_, worksheet) =>
            {
                var cell = worksheet.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                cell.InlineString = new DocumentFormat.OpenXml.Spreadsheet.InlineString(
                    new DocumentFormat.OpenXml.Spreadsheet.Text("{Value " + key + "}"));
            });
            using var xlsxInput = new MemoryStream(xlsx);
            Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(xlsxInput).Valid);
            var generated = Merge(provider, xlsx, """{"Value":1}""");
            using var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(new MemoryStream(generated), false);
            var cellOutput = Assert.Single(package.WorkbookPart!.WorksheetParts.Single().Worksheet!
                .Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>());
            Assert.Equal("Один рубль 00 копеек", cellOutput.InnerText);
        }

        var invalid = Docx.Create(Docx.P(Docx.Run("{Value rub_words}")));
        var error = Assert.Throws<TemplateException>(() => Merge(provider, invalid, """{"Value":true}"""));
        Assert.Equal("rub_words", Assert.Single(error.Errors).Formatter);
    }

    [Fact]
    public void Custom_only_registration_does_not_create_the_default_key()
    {
        using var provider = new ServiceCollection().AddMonoReport()
            .AddMonoReportSpelloutFormatter("only_here")
            .BuildServiceProvider();
        Assert.NotNull(provider.GetKeyedService<IFormatter>("only_here"));
        Assert.Null(provider.GetKeyedService<IFormatter>("spellout"));
        using var customInput = new MemoryStream(Docx.Create(Docx.P(Docx.Run("{Value only_here}"))));
        using var defaultInput = new MemoryStream(Docx.Create(Docx.P(Docx.Run("{Value spellout}"))));
        Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(customInput).Valid);
        Assert.False(provider.GetRequiredService<TemplateAnalyzer>().Analyze(defaultInput).Valid);
    }

    [Fact]
    public void Registration_accepts_a_unicode_template_identifier_key()
    {
        const string key = "слово_٢";
        using var provider = new ServiceCollection().AddMonoReport()
            .AddMonoReportSpelloutFormatter(key)
            .BuildServiceProvider();

        Assert.NotNull(provider.GetKeyedService<IFormatter>(key));
        var parsed = provider.GetRequiredService<TemplateLanguageParser>().Parse("{Value " + key + "}", "test");
        Assert.Empty(parsed.Diagnostics);
    }

    [Fact]
    public void Repeated_same_key_registrations_are_normal_keyed_services_and_distinct_keys_do_not_share_instances()
    {
        var services = new ServiceCollection().AddMonoReport()
            .AddMonoReportSpelloutFormatter()
            .AddMonoReportSpelloutFormatter()
            .AddMonoReportSpelloutFormatter("custom");
        using var provider = services.BuildServiceProvider();
        var defaults = provider.GetKeyedServices<IFormatter>("spellout").ToArray();
        Assert.Equal(2, defaults.Length);
        Assert.NotSame(defaults[0], defaults[1]);
        Assert.NotSame(defaults[1], provider.GetRequiredKeyedService<IFormatter>("custom"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("bad-key")]
    [InlineData("9spellout")]
    public void Registration_rejects_keys_the_template_grammar_cannot_represent(string key)
    {
        var services = new ServiceCollection();
        Assert.Throws<ArgumentException>(() => services.AddMonoReportSpelloutFormatter(key));
    }

    private static byte[] Merge(IServiceProvider provider, byte[] template, string json)
    {
        using var input = new MemoryStream(template);
        using var data = JsonDocument.Parse(json);
        return provider.GetRequiredService<MergeEngine>().Merge(input, data).Document;
    }
}
