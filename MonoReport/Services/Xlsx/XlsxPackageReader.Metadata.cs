using System.Collections.Immutable;
using System.Globalization;
using System.Xml.Linq;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxPackageReader
{
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace Ac = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac";
    private static readonly XNamespace Revision = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    private static readonly XNamespace CalcFeatures = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";
    private const string CalcFeaturesExtension = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}";
    private const string Excel15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

    private ImmutableArray<int> DisposableWorkbookNodes(XElement workbook) => workbook.Elements()
        .Select((node, index) => (node, index)).Where(item => IsSavePath(item.node) || IsRevisionPointer(item.node) || IsCalculationFeatures(item.node))
        .Select(item => item.index).ToImmutableArray();

    private bool IsCalculationFeatures(XElement node) =>
        node.Name == S + "extLst" && OnlyAttributes(node) && OnlyElementContent(node) && node.HasElements &&
        node.Elements().All(IsCalculationFeaturesExtension);

    private bool IsCalculationFeaturesExtension(XElement extension)
    {
        if (extension.Name != S + "ext" || !OnlyAttributes(extension, "uri") || !OnlyElementContent(extension) ||
            !string.Equals((string?)extension.Attribute("uri"), CalcFeaturesExtension, StringComparison.OrdinalIgnoreCase)) return false;
        var content = extension.Elements().ToArray();
        if (content.Length != 1 || content[0].Name != CalcFeatures + "calcFeatures" ||
            !OnlyAttributes(content[0]) || !OnlyElementContent(content[0]) || !content[0].HasElements) return false;
        // These are calculation-engine capabilities, not functions used by the report. Rendering forces full recalculation.
        return content[0].Elements().All(feature => feature.Name == CalcFeatures + "feature" &&
            OnlyAttributes(feature, "name") && feature.Attribute("name") is not null && !feature.Nodes().Any());
    }

    private string DescribeWorkbookFeature(XElement node)
    {
        var description = "element " + WorkbookIdentifier(node.Name.ToString());
        if (node.Name != S + "extLst") return description;
        var extension = node.Elements().FirstOrDefault(element => !IsCalculationFeaturesExtension(element));
        if (extension is null) return description;
        description += "/" + WorkbookIdentifier(extension.Name.ToString());
        // Extension URIs are arbitrary strings. Only disclose a GUID, never an arbitrary URI/attribute value.
        description += Guid.TryParse((string?)extension.Attribute("uri"), out var id)
            ? " (extension URI " + id.ToString("B").ToUpperInvariant() + ")" : " (unrecognized extension URI)";
        if (extension.Elements().FirstOrDefault() is { } payload)
            description += ", payload " + WorkbookIdentifier(payload.Name.ToString());
        return description;
    }

    private string WorkbookIdentifier(string value) => value.Length <= 240 &&
        value.All(character => char.IsAsciiLetterOrDigit(character) || "{}:/._-#".Contains(character))
        ? value : "[unrecognized identifier]";

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
