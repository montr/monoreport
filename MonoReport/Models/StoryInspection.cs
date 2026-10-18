using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace MonoReport.Models;

internal sealed record StoryInspection(
    IReadOnlyList<(OpenXmlPart Part, XDocument Xml)> Stories,
    IReadOnlyList<StructuralDiagnostic> Diagnostics,
    IReadOnlyDictionary<string, int> PartOrder,
    IReadOnlyList<(OpenXmlPart Part, XDocument Xml)> PackageXml);
