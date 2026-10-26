using System.Collections.Immutable;
using System.Globalization;
using System.Xml.Linq;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxPackageReader
{
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace Ac = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac";
    private static readonly XNamespace Revision = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    private const string Excel15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

    private ImmutableArray<int> DisposableWorkbookNodes(XElement workbook) => workbook.Elements()
        .Select((node, index) => (node, index)).Where(item => IsSavePath(item.node) || IsRevisionPointer(item.node))
        .Select(item => item.index).ToImmutableArray();

    private bool IsSavePath(XElement node)
    {
        if (node.Name != Mc + "AlternateContent" || !OnlyAttributes(node) || !OnlyElementContent(node)) return false;
        var branches = node.Elements().ToArray();
        if (branches.Length == 0 || branches[0].Name != Mc + "Choice" ||
            branches.Count(branch => branch.Name == Mc + "Fallback") > 1 ||
            branches.Take(branches.Length - 1).Any(branch => branch.Name == Mc + "Fallback")) return false;
        foreach (var branch in branches)
        {
            if (!OnlyElementContent(branch)) return false;
            if (branch.Name == Mc + "Choice")
            {
                if (!OnlyAttributes(branch, "Requires")) return false;
                var prefixes = ((string?)branch.Attribute("Requires"))?.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries) ?? [];
                if (prefixes.Length != 1 || branch.GetNamespaceOfPrefix(prefixes[0])?.NamespaceName != Excel15) return false;
                var content = branch.Elements().ToArray();
                if (content.Length != 1 || content[0].Name != Ac + "absPath" || !OnlyAttributes(content[0], "url") ||
                    content[0].Attribute("url") is null || content[0].Nodes().Any()) return false;
            }
            else if (branch.Name != Mc + "Fallback" || !OnlyAttributes(branch) || branch.HasElements) return false;
        }
        return true;
    }

    private bool IsRevisionPointer(XElement node)
    {
        XNamespace revision6 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6";
        XNamespace revision10 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";
        if (node.Name != Revision + "revisionPtr" || node.Nodes().Any()) return false;
        foreach (var attribute in node.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration))
        {
            if (attribute.Name == "revIDLastSave" || attribute.Name == revision6 + "coauthVersionLast" || attribute.Name == revision6 + "coauthVersionMax")
            {
                if (!ulong.TryParse(attribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out _)) return false;
            }
            else if (attribute.Name == "documentId")
            {
                // Excel also writes composite IDs such as 13_ncr:1_{GUID}; this SDK attribute is a string.
                if (string.IsNullOrWhiteSpace(attribute.Value)) return false;
            }
            else if (attribute.Name == revision10 + "uidLastSave")
            {
                if (!Guid.TryParse(attribute.Value, out _)) return false;
            }
            else return false;
        }
        return node.Attribute("revIDLastSave") is not null;
    }

    private bool OnlyAttributes(XElement node, params XName[] names) =>
        node.Attributes().All(attribute => attribute.IsNamespaceDeclaration || names.Contains(attribute.Name));
    private bool OnlyElementContent(XElement node) =>
        node.Nodes().All(child => child is XElement || child is XText text && string.IsNullOrWhiteSpace(text.Value));
}
