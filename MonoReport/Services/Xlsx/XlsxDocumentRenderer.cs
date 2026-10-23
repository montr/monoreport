using System.Xml.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxDocumentRenderer
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal void Render(SpreadsheetDocument document, XlsxWorkbookPlan plan, XlsxOperationContext context)
    {
        context.ThrowIfErrors();
        // Validate the entire workbook before writing any worksheet; source follower XML is never a fallback.
        foreach (var sheet in plan.Sheets)
            foreach (var cell in sheet.Rows.SelectMany(row => row.Cells))
            {
                if (context.LimitReached) break;
                if (!cell.Suppressed && cell.Source.HasFormula && string.IsNullOrWhiteSpace(cell.Formula))
                    context.Report(new StructuralDiagnostic(DiagnosticCodes.XlsxUnsafeFormula,
                        "Formula placement has no finalized text. Finalize every surviving scalar A1 formula before rendering.",
                        Location: sheet.Geometry.Source.Part + ", sheet '" + sheet.Geometry.Source.Name + "', " + cell.Source.Id.Address).ToException());
            }
        context.ThrowIfErrors();
        var parts = document.WorkbookPart!.WorksheetParts.ToDictionary(part => part.Uri.ToString(), StringComparer.Ordinal);
        foreach (var sheet in plan.Sheets)
        {
            var root = new XElement(sheet.Settings);
            var data = root.Element(S + "sheetData")!;
            foreach (var row in sheet.Rows)
            {
                var element = new XElement(row.Properties);
                element.SetAttributeValue("r", row.OutputRow);
                foreach (var placement in row.Cells)
                {
                    var cell = new XElement(placement.Source.Xml);
                    cell.SetAttributeValue("r", placement.Output);
                    if (placement.Suppressed)
                    {
                        cell.Elements(S + "f").Remove(); WriteValue(cell, new(null, null));
                    }
                    else if (placement.Formula is { } formula)
                    {
                        cell.Elements(S + "f").Remove(); cell.Elements(S + "v").Remove(); cell.Elements(S + "is").Remove();
                        cell.Attribute("t")?.Remove(); cell.Add(new XElement(S + "f", formula));
                    }
                    else if (placement.Value is { } value) WriteValue(cell, value);
                    else if (placement.Source.SharedString is { } text)
                    {
                        cell.Elements(S + "v").Remove();
                        cell.SetAttributeValue("t", "inlineStr");
                        cell.Add(new XElement(S + "is", text.Elements().Select(child => new XElement(child))));
                    }
                    element.Add(cell);
                }
                data.Add(element);
            }
            using var output = parts[sheet.Geometry.Source.Part].GetStream(FileMode.Create, FileAccess.Write);
            using var writer = XmlWriter.Create(output, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false), NewLineHandling = NewLineHandling.Entitize
            });
            new XDocument(root).Save(writer);
        }
        var workbook = document.WorkbookPart!;
        var definedNames = workbook.Workbook!.DefinedNames;
        if (definedNames is not null)
            foreach (var (name, index) in definedNames.Elements<DefinedName>().Select((name, index) => (name, index))) name.Text = plan.Names[index].Value;
        if (workbook.SharedStringTablePart is { } shared) workbook.DeletePart(shared);
        if (workbook.CalculationChainPart is { } chain) workbook.DeletePart(chain);
        var calculation = workbook.Workbook.CalculationProperties;
        if (calculation is null) { calculation = new CalculationProperties(); workbook.Workbook.CalculationProperties = calculation; }
        calculation.CalculationMode = CalculateModeValues.Auto;
        calculation.FullCalculationOnLoad = true;
        calculation.ForceFullCalculation = true;
    }

    private static void WriteValue(XElement cell, XlsxPreparedValue value)
    {
        cell.Elements(S + "v").Remove();
        cell.Elements(S + "is").Remove();
        cell.SetAttributeValue("t", value.Type);
        if (value.Value is null || value.Value.Length == 0)
        {
            cell.Attribute("t")?.Remove();
            return;
        }
        if (value.Type == "inlineStr") cell.Add(new XElement(S + "is", new XElement(S + "t",
            new XAttribute(XNamespace.Xml + "space", "preserve"), EncodeString(value.Value))));
        else cell.Add(new XElement(S + "v", value.Value));
    }

    private static string EncodeString(string value) => EscapedCharacter().Replace(EscapePrefix().Replace(value, "_x005F_"),
        match => "_x" + ((int)match.Value[0]).ToString("X4", System.Globalization.CultureInfo.InvariantCulture) + "_");

    // Protect only source underscores; never re-encode an escape emitted by this operation.
    [GeneratedRegex(@"_(?=x[0-9A-Fa-f]{4}_)", RegexOptions.CultureInvariant)]
    private static partial Regex EscapePrefix();

    // Authored Xstrings may contain escaped XML controls; JSON controls still reject during binding.
    [GeneratedRegex(@"[\x00-\x08\x0B-\x0D\x0E-\x1F\uFFFE\uFFFF]", RegexOptions.CultureInvariant)]
    private static partial Regex EscapedCharacter();
}
