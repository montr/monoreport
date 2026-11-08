using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Core.Diagnostics;
using MonoReport.Docx.Planning;

namespace MonoReport.Docx;

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
        catch (Exception error) when (IsPackageReadFailure(error) || error is XmlException)
        {
            diagnostic = PackageDiagnostic();
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
        MainDocumentPart? main;
        try
        {
            main = document.MainDocumentPart;
        }
        catch (Exception error) when (IsPackageReadFailure(error) || error is XmlException)
        {
            return new StoryInspection([], [PackageDiagnostic()],
                new Dictionary<string, int>(StringComparer.Ordinal), []);
        }

        if (main is null)
            return new StoryInspection([],
                [
                    new StructuralDiagnostic(DiagnosticCodes.TemplateUnsupportedStructure,
                        "DOCX has no main document part.")
                ],
                new Dictionary<string, int>(StringComparer.Ordinal), []);
        var stories = new List<(OpenXmlPart, XDocument)>();
        var diagnostics = new List<StructuralDiagnostic>();
        var partOrder = new Dictionary<string, int>(StringComparer.Ordinal);
        OpenXmlPart[] storyParts;
        try
        {
            storyParts = Parts(main).ToArray();
        }
        catch (Exception error) when (IsPackageReadFailure(error) || error is XmlException)
        {
            return new StoryInspection([], [PackageDiagnostic()], partOrder, []);
        }

        foreach (var part in storyParts)
        {
            partOrder.Add(part.Uri.ToString(), partOrder.Count);
            try
            {
                using var input = part.GetStream(FileMode.Open, FileAccess.Read);
                using var reader =
                    XmlReader.Create(input, new XmlReaderSettings {DtdProcessing = DtdProcessing.Prohibit});
                var xml = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
                var expected = part switch
                {
                    MainDocumentPart => "document", HeaderPart => "hdr", FooterPart => "ftr",
                    FootnotesPart => "footnotes", EndnotesPart => "endnotes", _ => ""
                };
                if (xml.Root?.Name != WordXml.DocumentElementName.Namespace + expected ||
                    (part is MainDocumentPart && xml.Root.Element(WordXml.BodyElementName) is null))
                {
                    diagnostics.Add(new StructuralDiagnostic(DiagnosticCodes.TemplateUnsupportedStructure,
                        "Unsupported DOCX story structure.",
                        Location: part.Uri.ToString()));
                    continue;
                }

                stories.Add((part, xml));
            }
            catch (XmlException)
            {
                diagnostics.Add(new StructuralDiagnostic(DiagnosticCodes.TemplateInvalidXml, "Invalid DOCX story XML.",
                    Location: part.Uri.ToString()));
            }
            catch (Exception error) when (IsPackageReadFailure(error))
            {
                diagnostics.Add(PackageDiagnostic(part.Uri.ToString()));
            }
        }

        var packageXml = new List<(OpenXmlPart Part, XDocument Xml)>(stories);
        OpenXmlPart[] packageParts;
        try
        {
            packageParts = PackageParts(document).OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal)
                .ToArray();
        }
        catch (Exception error) when (IsPackageReadFailure(error) || error is XmlException)
        {
            diagnostics.Add(PackageDiagnostic());
            return new StoryInspection(stories, diagnostics, partOrder, packageXml);
        }

        foreach (var part in packageParts)
        {
            if (partOrder.ContainsKey(part.Uri.ToString())) continue;
            partOrder.Add(part.Uri.ToString(), partOrder.Count);
            try
            {
                if (!part.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase) &&
                    !part.ContentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase)) continue;
                using var input = part.GetStream(FileMode.Open, FileAccess.Read);
                using var reader =
                    XmlReader.Create(input, new XmlReaderSettings {DtdProcessing = DtdProcessing.Prohibit});
                packageXml.Add((part, XDocument.Load(reader, LoadOptions.PreserveWhitespace)));
            }
            catch (XmlException)
            {
                diagnostics.Add(new StructuralDiagnostic(DiagnosticCodes.TemplateInvalidXml,
                    "Invalid DOCX auxiliary XML.", Location: part.Uri.ToString()));
            }
            catch (Exception error) when (IsPackageReadFailure(error))
            {
                diagnostics.Add(PackageDiagnostic(part.Uri.ToString()));
            }
        }

        return new StoryInspection(stories, diagnostics, partOrder, packageXml);
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
        foreach (var part in main.HeaderParts.OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal))
            yield return part;
        foreach (var part in main.FooterParts.OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal))
            yield return part;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
    }

    private static bool IsPackageReadFailure(Exception error)
    {
        return error is OpenXmlPackageException or InvalidDataException or FileFormatException or
            InvalidOperationException or KeyNotFoundException or ArgumentException or UriFormatException;
    }

    private static StructuralDiagnostic PackageDiagnostic(string? location = null)
    {
        return new StructuralDiagnostic(DiagnosticCodes.TemplateInvalidPackage, "DOCX package is unreadable.",
            Location: location);
    }
}
