using System.Collections.Generic;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using MonoReport.Core.Diagnostics;

namespace MonoReport.Docx.Planning;

internal sealed record StoryInspection(
    IReadOnlyList<(OpenXmlPart Part, XDocument Xml)> Stories,
    IReadOnlyList<StructuralDiagnostic> Diagnostics,
    IReadOnlyDictionary<string, int> PartOrder,
    IReadOnlyList<(OpenXmlPart Part, XDocument Xml)> PackageXml);