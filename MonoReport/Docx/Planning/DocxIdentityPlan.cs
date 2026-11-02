using System.Xml.Linq;
using MonoReport.Core.Diagnostics;
using System.Collections.Generic;

namespace MonoReport.Docx.Planning;

internal sealed record StructuralOwner(BlockPlan Block, string Branch);

internal sealed record IdentityLocation(
    XElement Element,
    string Part,
    string Location,
    long SourceOrder,
    XElement Root,
    IReadOnlyList<StructuralOwner> Owners);

internal sealed record SourceRange(
    string Kind,
    string Id,
    IReadOnlyList<IdentityLocation> Members,
    XElement? CommentDefinition = null)
{
    internal IdentityLocation First => Members[0];

    internal string? Name => (string?) First.Element.Attribute(
        XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main") + "name");
}

internal sealed record DocxIdentityPlan(
    IReadOnlyDictionary<XElement, SourceRange> Ranges,
    IReadOnlyList<StructuralDiagnostic> Diagnostics,
    IReadOnlyList<XElement>? DeletedCommentDefinitions = null);