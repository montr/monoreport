using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Models;

namespace MonoReport.Services;

public sealed class DocxStoryReader
{
    public WordprocessingDocument Open(Stream stream, bool editable)
    {
        var document = OpenForValidation(stream, editable, out var diagnostic);
        if (diagnostic is not null) throw diagnostic.ToException();
        return document!;
    }

    internal WordprocessingDocument? OpenForValidation(Stream stream, bool editable,
        out StructuralDiagnostic? diagnostic)
    {
        try
        {
            diagnostic = null;
            return WordprocessingDocument.Open(stream, editable);
        }
        catch (Exception error) when (error is OpenXmlPackageException or InvalidDataException or FileFormatException)
        {
            diagnostic = new("template.invalid_package", "DOCX package is unreadable.");
            return null;
        }
    }

    public IReadOnlyList<(OpenXmlPart Part, XDocument Xml)> Read(WordprocessingDocument document)
    {
        var inspection = Inspect(document);
        if (inspection.Diagnostics.Count != 0) throw inspection.Diagnostics[0].ToException();
        return inspection.Stories;
    }

    internal StoryInspection Inspect(WordprocessingDocument document)
    {
        var main = document.MainDocumentPart;
        if (main is null)
            return new([], [new("template.unsupported_structure", "DOCX has no main document part.")],
                new Dictionary<string, int>(StringComparer.Ordinal), []);
        var stories = new List<(OpenXmlPart, XDocument)>();
        var diagnostics = new List<StructuralDiagnostic>();
        var partOrder = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var part in Parts(main))
        {
            partOrder.Add(part.Uri.ToString(), partOrder.Count);
            try
            {
                using var input = part.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
                var xml = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
                var expected = part switch
                {
                    MainDocumentPart => "document", HeaderPart => "hdr", FooterPart => "ftr",
                    FootnotesPart => "footnotes", EndnotesPart => "endnotes", _ => ""
                };
                if (xml.Root?.Name != WordXml.DocumentElementName.Namespace + expected ||
                    (part is MainDocumentPart && xml.Root.Element(WordXml.BodyElementName) is null))
                {
                    diagnostics.Add(new("template.unsupported_structure", "Unsupported DOCX story structure.",
                        Location: part.Uri.ToString()));
                    continue;
                }
                stories.Add((part, xml));
            }
            catch (XmlException)
            {
                diagnostics.Add(new("template.invalid_xml", "Invalid DOCX story XML.", Location: part.Uri.ToString()));
            }
        }
        var packageXml = new List<(OpenXmlPart Part, XDocument Xml)>(stories);
        foreach (var part in PackageParts(document).OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal))
        {
            if (partOrder.ContainsKey(part.Uri.ToString())) continue;
            partOrder.Add(part.Uri.ToString(), partOrder.Count);
            if (!part.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase) &&
                !part.ContentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase)) continue;
            try
            {
                using var input = part.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
                packageXml.Add((part, XDocument.Load(reader, LoadOptions.PreserveWhitespace)));
            }
            catch (XmlException)
            {
                diagnostics.Add(new("template.invalid_xml", "Invalid DOCX auxiliary XML.", Location: part.Uri.ToString()));
            }
        }
        return new(stories, diagnostics, partOrder, packageXml);
    }

    private static IEnumerable<OpenXmlPart> PackageParts(WordprocessingDocument document)
    {
        var pending = new Queue<OpenXmlPart>(document.Parts.Select(item => item.OpenXmlPart));
        var seen = new HashSet<OpenXmlPart>();
        while (pending.TryDequeue(out var part))
        {
            if (!seen.Add(part)) continue;
            foreach (var child in part.Parts) pending.Enqueue(child.OpenXmlPart);
            yield return part;
        }
    }

    private static IEnumerable<OpenXmlPart> Parts(MainDocumentPart main)
    {
        yield return main;
        foreach (var part in main.HeaderParts.OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal)) yield return part;
        foreach (var part in main.FooterParts.OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal)) yield return part;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
    }
}
