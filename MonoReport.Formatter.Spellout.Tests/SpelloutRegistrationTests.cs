using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Core.Formatting;
using MonoReport.Core.Templates;
using Xunit;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace MonoReport.Formatter.Spellout.Tests;

public sealed class SpelloutRegistrationTests
{
    [Fact]
    public void Default_formatter_renders_numbers_and_custom_key_renders_money_in_documents()
    {
        using var provider = new ServiceCollection().AddMonoReport()
            .AddMonoReportSpelloutFormatter()
            .AddMonoReportSpelloutFormatter("euro_words")
            .BuildServiceProvider();

        var docx = CreateDocx("{Value spellout}");
        using var docxInput = new MemoryStream(docx);
        Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(docxInput).Valid);
        Assert.Equal("One", ReadDocx(Generate(provider, docx, """{"Value":1}""")));

        var xlsx = CreateXlsx("{Value euro_words mode=\"money\" currency=\"EUR\"}");
        using var xlsxInput = new MemoryStream(xlsx);
        Assert.True(provider.GetRequiredService<TemplateAnalyzer>().Analyze(xlsxInput).Valid);
        Assert.Equal("Two euros 01 cent", ReadXlsx(Generate(provider, xlsx, """{"Value":2.01}""")));

        var error = Assert.Throws<TemplateException>(() => Generate(provider, docx, """{"Value":true}"""));
        Assert.Equal("spellout", Assert.Single(error.Errors).Formatter);
    }

    [Fact]
    public void Custom_only_registration_does_not_create_the_default_key()
    {
        using var provider = new ServiceCollection().AddMonoReport()
            .AddMonoReportSpelloutFormatter("only_here")
            .BuildServiceProvider();

        Assert.NotNull(provider.GetKeyedService<IFormatter>("only_here"));
        Assert.Null(provider.GetKeyedService<IFormatter>("spellout"));
        using var customInput = new MemoryStream(CreateDocx("{Value only_here}"));
        using var defaultInput = new MemoryStream(CreateDocx("{Value spellout}"));
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
        using var input = new MemoryStream(CreateDocx("{Value " + key + "}"));
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
        Assert.Same(provider.GetRequiredService<SpelloutService>(), provider.GetRequiredService<SpelloutService>());
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

    private static byte[] CreateDocx(string expression)
    {
        using var output = new MemoryStream();
        using (var document = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(
                new Body(
                    new Paragraph(
                        new Run(
                            new Text(expression)))));
        }

        return output.ToArray();
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
                InlineString = new InlineString(new global::DocumentFormat.OpenXml.Spreadsheet.Text(expression)
                    {Space = SpaceProcessingModeValues.Preserve})
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