using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Cli.Tests;
using MonoReport.Core.Diagnostics;
using MonoReport.Core.Templates;
using MonoReport.Docx;
using Xunit;

namespace MonoReport.Tests.Docx;

public sealed class DocxPackageDiagnosticsTests
{
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    [Theory]
    [InlineData("malformed-relationships")]
    [InlineData("missing-target")]
    public void Package_relationship_failures_are_structured_in_analyze_and_generate(string scenario)
    {
        var source = WithHeader();
        source = RewriteEntry(source, "word/_rels/document.xml.rels", xml =>
        {
            if (scenario == "malformed-relationships") return xml[..^"</Relationships>".Length];

            var relationships = XDocument.Parse(xml);
            var header = relationships.Root!.Elements(Relationships + "Relationship").Single(relationship =>
                ((string?) relationship.Attribute("Type"))?.EndsWith("/header", StringComparison.Ordinal) == true);
            header.SetAttributeValue("Target", "missing-header.xml");
            return relationships.ToString(SaveOptions.DisableFormatting);
        });

        using var input = new MemoryStream(source);
        var analysis = DocxIdentityTests.Analyzer().Analyze(input);

        Assert.False(analysis.Valid);
        var issue = Assert.Single(analysis.Errors);
        Assert.Equal(DiagnosticCodes.TemplateInvalidPackage, issue.Code);

        input.Position = 0;
        using var data = JsonDocument.Parse("{}");
        var error = Assert.Throws<TemplateException>(() => DocxFixture.CreateGenerator().Generate(input, data));
        Assert.Equal(issue.Code, error.Code);
        Assert.Equal(issue.Message, error.Errors[0].Message);
        Assert.Equal(source, input.ToArray());
    }

    private static byte[] WithHeader()
    {
        var source = DocxFixture.Create(DocxFixture.P(DocxFixture.Run("Report")));
        using var stream = new MemoryStream();
        stream.Write(source);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var header = document.MainDocumentPart!.AddNewPart<HeaderPart>();
            using var target = header.GetStream(FileMode.Create, FileAccess.Write);
            new XDocument(new XElement(WordXml.HeaderElementName,
                DocxFixture.P(DocxFixture.Run("Header")))).Save(target);
        }

        return stream.ToArray();
    }

    private static byte[] RewriteEntry(byte[] source, string path, Func<string, string> rewrite)
    {
        using var buffer = new MemoryStream();
        buffer.Write(source);
        buffer.Position = 0;
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, true))
        {
            var entry = archive.GetEntry(path) ?? throw new InvalidOperationException("Fixture part not found: " + path);
            string contents;
            using (var reader = new StreamReader(entry.Open())) contents = reader.ReadToEnd();
            entry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry(path).Open());
            writer.Write(rewrite(contents));
        }

        return buffer.ToArray();
    }
}
