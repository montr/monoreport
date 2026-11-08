using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Docx;
using MonoReport.Formatter.Spellout;
using Xunit;

namespace MonoReport.Cli.Tests;

internal static class DocxFixture
{
    private static readonly ServiceProvider Services =
        new ServiceCollection().AddMonoReport().AddMonoReportSpelloutFormatter().BuildServiceProvider();

    internal static IReadOnlyDictionary<string, XDocument> ReadStories(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        var parts = new List<OpenXmlPart> {main};
        parts.AddRange(main.HeaderParts);
        parts.AddRange(main.FooterParts);
        if (main.FootnotesPart is not null) parts.Add(main.FootnotesPart);
        if (main.EndnotesPart is not null) parts.Add(main.EndnotesPart);
        return parts.ToDictionary(p => p.Uri.ToString(), p =>
        {
            using var input = p.GetStream();
            return XDocument.Load(input);
        });
    }

    internal static XElement Character(string type)
    {
        return new XElement(WordXml.FieldCharacterElementName,
            new XAttribute(WordXml.FieldCharacterTypeAttributeName, type));
    }

    internal static XElement Run(string text)
    {
        return new XElement(WordXml.RunElementName, new XElement(WordXml.TextElementName, text));
    }

    internal static XElement P(params object[] contents)
    {
        return new XElement(WordXml.ParagraphElementName, contents);
    }

    internal static XElement Field(string code, string cached = "old", XElement? properties = null)
    {
        return new XElement(WordXml.SimpleFieldElementName, new XAttribute(WordXml.InstructionAttributeName, code),
            new XElement(WordXml.RunElementName, properties, new XElement(WordXml.TextElementName, cached)));
    }

    internal static XElement Row(params XElement[] contents)
    {
        return new XElement(WordXml.TableRowElementName,
            contents.Select(x => new XElement(WordXml.TableCellElementName, P(x))));
    }

    internal static XElement Table(params XElement[] rows)
    {
        return new XElement(WordXml.TableElementName, new XElement(WordXml.TablePropertiesElementName),
            new XElement(WordXml.TableGridElementName,
                Enumerable.Range(0, 3).Select(_ =>
                    new XElement(WordXml.GridColumnElementName, new XAttribute(WordXml.WidthAttributeName, "2000")))),
            rows);
    }

    internal static XElement Properties()
    {
        return new XElement(WordXml.RunPropertiesElementName,
            new XElement(WordXml.RunFontsElementName, new XAttribute(WordXml.AsciiFontAttributeName, "Arial"),
                new XAttribute(WordXml.HighAnsiFontAttributeName, "Arial")), new XElement(WordXml.BoldElementName),
            new XElement(WordXml.ItalicElementName),
            new XElement(WordXml.ColorElementName, new XAttribute(WordXml.ValueAttributeName, "FF0000")),
            new XElement(WordXml.FontSizeElementName, new XAttribute(WordXml.ValueAttributeName, "28")),
            new XElement(WordXml.HighlightElementName, new XAttribute(WordXml.ValueAttributeName, "yellow")),
            new XElement(WordXml.UnderlineElementName,
                new XAttribute(WordXml.ValueAttributeName, WordXml.SingleUnderlineValue)),
            new XElement(WordXml.LanguageElementName, new XAttribute(WordXml.ValueAttributeName, "ru-RU")));
    }

    internal static byte[] Create(params XElement[] content)
    {
        using (var stream = new MemoryStream())
        {
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                using (var target = main.GetStream(FileMode.Create, FileAccess.Write))
                {
                    new XDocument(new XElement(WordXml.DocumentElementName,
                        new XAttribute(WordXml.WordprocessingNamespaceDeclarationName, WordXml.WordprocessingNamespace),
                        new XElement(WordXml.BodyElementName, content))).Save(target);
                }
            }

            return stream.ToArray();
        }
    }

    internal static DocumentGenerator CreateGenerator()
    {
        return Services.GetRequiredService<DocumentGenerator>();
    }

    internal static GenerationResult Generate(byte[] template, string json, GenerationOptions? options = null,
        FileFormatVersions validationVersion = FileFormatVersions.Office2007)
    {
        using (var stream = new MemoryStream(template))
        {
            using (var data = JsonDocument.Parse(json))
            {
                var result = CreateGenerator().Generate(stream, data, options);
                Assert.True(stream.CanRead);
                Assert.Equal(template, stream.ToArray());
                AssertValid(result.Document, validationVersion);
                return result;
            }
        }
    }

    internal static XDocument Read(byte[] bytes)
    {
        using (var stream = new MemoryStream(bytes))
        {
            using (var doc = WordprocessingDocument.Open(stream, false))
            {
                using (var part = doc.MainDocumentPart!.GetStream())
                {
                    return XDocument.Load(part);
                }
            }
        }
    }

    internal static string Text(XContainer xml)
    {
        return string.Concat(xml.Descendants(WordXml.TextElementName).Select(x => x.Value));
    }

    internal static void AssertValid(byte[] bytes, FileFormatVersions version = FileFormatVersions.Office2007)
    {
        using (var stream = new MemoryStream(bytes))
        {
            using (var doc = WordprocessingDocument.Open(stream, false))
            {
                var errors = new OpenXmlValidator(version).Validate(doc)
                    .Select(x => x.Description + " " + x.Path?.XPath).ToArray();
                Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
            }
        }
    }
}