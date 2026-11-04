using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Core.Formatting;
using MonoReport.Core.Templates;
using Xunit;
using System;
using System.IO;
using System.Linq;

namespace MonoReport.Formatter.Spellout.Tests;

public sealed class SpelloutRegistrationTests
{
    [Fact]
    public void Default_and_custom_keys_analyze_and_generate_in_docx_and_xlsx()
    {
        using var provider = new ServiceCollection().AddMonoReport()
            .AddMonoReportSpelloutFormatter()
            .AddMonoReportSpelloutFormatter("rub_words")
            .AddMonoReportSpelloutFormatter("amount_words")
            .BuildServiceProvider();

        foreach (var key in new[] {"spellout", "rub_words", "amount_words"})
        {
            var expression = "{Value " + key + " mode=\"money\" currency=\"RUB\"}";
            var docx = SpelloutBehaviorTests.CreateDocx(expression);
            using var docxInput = new MemoryStream(docx);
            Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(docxInput).Valid);
            Assert.Equal("One ruble 00 copecks", ReadDocx(Generate(provider, docx, """{"Value":1}""")));

            var xlsx = CreateXlsx(expression);
            using var xlsxInput = new MemoryStream(xlsx);
            Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(xlsxInput).Valid);
            Assert.Equal("One ruble 00 copecks", ReadXlsx(Generate(provider, xlsx, """{"Value":1}""")));
        }

        var invalid = SpelloutBehaviorTests.CreateDocx("{Value rub_words}");
        var error = Assert.Throws<TemplateException>(() => Generate(provider, invalid, """{"Value":true}"""));
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
        using var customInput = new MemoryStream(SpelloutBehaviorTests.CreateDocx("{Value only_here}"));
        using var defaultInput = new MemoryStream(SpelloutBehaviorTests.CreateDocx("{Value spellout}"));
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
        using var input = new MemoryStream(SpelloutBehaviorTests.CreateDocx("{Value " + key + "}"));
        Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(input).Valid);
    }

    [Fact]
    public void Repeated_same_key_registrations_are_keyed_services_and_distinct_keys_do_not_share_instances()
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
        Assert.Throws<ArgumentException>(() => new ServiceCollection().AddMonoReportSpelloutFormatter(key));
    }

    private static byte[] Generate(IServiceProvider provider, byte[] template, string json)
    {
        using var input = new MemoryStream(template);
        using var data = JsonDocument.Parse(json);
        return provider.GetRequiredService<DocumentGenerator>().Generate(input, data).Document;
    }

    private static string ReadDocx(byte[] source)
    {
        using var stream = new MemoryStream(source);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.Document!.Body!.InnerText;
    }

    private static string ReadXlsx(byte[] source)
    {
        using var stream = new MemoryStream(source);
        using var document = SpreadsheetDocument.Open(stream, false);
        return Assert.Single(document.WorkbookPart!.WorksheetParts.Single().Worksheet!
            .Descendants<Cell>()).InnerText;
    }

    private static byte[] CreateXlsx(string expression)
    {
        using var output = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook))
        {
            var workbook = document.AddWorkbookPart();
            workbook.Workbook = new Workbook();
            var worksheet = workbook.AddNewPart<WorksheetPart>();
            worksheet.Worksheet = new Worksheet(new SheetData(new Row(new Cell
            {
                CellReference = "B3",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(expression) {Space = SpaceProcessingModeValues.Preserve})
            }) {RowIndex = 3}));
            workbook.Workbook.Append(new Sheets(new Sheet
            {
                Name = "Report",
                SheetId = 1,
                Id = workbook.GetIdOfPart(worksheet)
            }));
            workbook.Workbook.Append(new DefinedNames(new DefinedName("'Report'!$B$3:$D$5")
            {
                Name = "_xlnm.Print_Area",
                LocalSheetId = 0
            }));
        }

        return output.ToArray();
    }
}
