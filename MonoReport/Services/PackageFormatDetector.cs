using System.IO.Packaging;
using System.Xml;
using MonoReport.Models;
using DocumentKind = MonoReport.Models.DocumentFormat;

namespace MonoReport.Services;

internal sealed class PackageFormatDetector
{
    internal DocumentKind Detect(Stream buffer)
    {
        var position = buffer.Position;
        try
        {
            using var package = Package.Open(buffer, FileMode.Open, FileAccess.Read);
            var roots = package.GetRelationships().Where(item => item.RelationshipType is
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" or
                "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument").ToArray();
            if (roots.Length != 1 || roots[0].TargetMode != TargetMode.Internal) throw Invalid();
            var uri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), roots[0].TargetUri);
            if (!package.PartExists(uri)) throw Invalid();
            return package.GetPart(uri).ContentType switch
            {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" or
                "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml" or
                "application/vnd.ms-word.document.macroEnabled.main+xml" or
                "application/vnd.ms-word.template.macroEnabledTemplate.main+xml" => DocumentKind.Docx,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" => DocumentKind.Xlsx,
                _ => throw Invalid()
            };
        }
        catch (Exception error) when (error is InvalidDataException or FileFormatException or XmlException or ArgumentException)
        {
            throw Invalid();
        }
        finally { buffer.Position = position; }
    }

    private static TemplateException Invalid() => new("Use a readable Word or ordinary XLSX OOXML package.",
        code: DiagnosticCodes.TemplateInvalidPackage);
}
