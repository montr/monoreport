using System.Collections.Immutable;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxGeometryPlanner
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    internal XlsxWorkbookPlan Plan(XlsxTemplatePlan template, ImmutableDictionary<SourceCellId, XlsxPreparedValue> values,
        XlsxOperationContext context)
    {
        var sheets = ImmutableArray.CreateBuilder<XlsxSheetOutput>();
        var names = template.Source.Names.Select(name => new XElement(name)).ToArray();
        if (names.Length != template.Source.Sheets.Select(sheet => sheet.PrintAreaNameIndex).Distinct().Count())
            throw new StructuralDiagnostic(DiagnosticCodes.XlsxUnsupportedFeature, "Unsupported or unresolved workbook names.",
                Location: "/xl/workbook.xml", Explanation: "Remove defined names/print titles or invalid PrintArea scopes until their reference transformation is implemented.").ToException();
        foreach (var sheet in template.Sheets)
        {
            if (context.LimitReached) break;
            try
            {
                ValidateFeatures(template.Source, sheet.Source);
                var area = sheet.Source.PrintArea;
                var dimensions = Dimensions(area.Height, area.Width, 1, 1);
                var geometry = new XlsxSheetGeometry(sheet.Source, new(new(1, 1), new(dimensions.Rows, dimensions.Columns)),
                    1 - area.First.Row, 1 - area.First.Column);
                var settings = new XElement(sheet.Source.Xml.Root!);
                settings.Element(S + "sheetData")!.RemoveNodes();
                CropColumns(settings, geometry);
                CropMerges(settings, geometry);
                CropViews(settings, geometry);
                CropBreaks(settings, geometry, rows: true);
                CropBreaks(settings, geometry, rows: false);
                settings.Element(S + "dimension")?.Remove();
                var dimension = new XElement(S + "dimension", new XAttribute("ref", geometry.OutputArea));
                if (settings.Element(S + "sheetPr") is { } properties) properties.AddAfterSelf(dimension);
                else settings.AddFirst(dimension);
                var byRow = sheet.Cells.GroupBy(cell => cell.Source.Id.Address.Row).ToDictionary(group => group.Key, group => group.ToArray());
                var rows = ImmutableArray.CreateBuilder<XlsxRowPlan>();
                foreach (var row in sheet.Source.Xml.Root!.Element(S + "sheetData")!.Elements(S + "row"))
                {
                    var sourceRow = (int)row.Attribute("r")!;
                    if (sourceRow < area.First.Row || sourceRow > area.Last.Row) continue;
                    var propertiesCopy = new XElement(row); propertiesCopy.RemoveNodes(); propertiesCopy.Attribute("spans")?.Remove();
                    var cells = ImmutableArray.CreateBuilder<XlsxCellPlacement>();
                    foreach (var cell in byRow.GetValueOrDefault(sourceRow) ?? [])
                    {
                        var output = Move(cell.Source.Id.Address, geometry);
                        cells.Add(new(new(cell.Source.Id, null, null), output, cell.Source, values.GetValueOrDefault(cell.Source.Id)));
                    }
                    rows.Add(new(propertiesCopy, checked(sourceRow + geometry.RowOffset), cells.ToImmutable()));
                }
                names[sheet.Source.PrintAreaNameIndex].Value = SheetPrefix(sheet.Source.Name) + Absolute(geometry.OutputArea);
                sheets.Add(new(geometry, settings, rows.ToImmutable()));
            }
            catch (TemplateException error) { context.Report(error); }
        }
        context.ThrowIfErrors();
        context.OutputCells = sheets.Sum(sheet => sheet.Rows.Sum(row => (long)row.Cells.Length));
        return new(sheets.ToImmutable(), names.ToImmutableArray());
    }

    internal XlsxRangeImage ProjectRange(XlsxRectangle range, XlsxSheetGeometry geometry)
    {
        var area = geometry.Source.PrintArea;
        var first = new XlsxAddress(Math.Max(range.First.Row, area.First.Row), Math.Max(range.First.Column, area.First.Column));
        var last = new XlsxAddress(Math.Min(range.Last.Row, area.Last.Row), Math.Min(range.Last.Column, area.Last.Column));
        if (first.Row > last.Row || first.Column > last.Column) return new(XlsxRangeImageKind.Empty, null, false);
        var output = new XlsxRectangle(Move(first, geometry), Move(last, geometry));
        return new(first == last ? XlsxRangeImageKind.Cell : XlsxRangeImageKind.Rectangle, output,
            area.Contains(range.First) && area.Contains(range.Last));
    }

    private static void ValidateFeatures(XlsxSourceInventory inventory, XlsxSheetSource sheet)
    {
        var safeWorkbook = new HashSet<string>(new[] { "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "definedNames", "calcPr" }
            .Select(name => (S + name).ToString()), StringComparer.Ordinal);
        foreach (var feature in inventory.WorkbookFeatures)
            if (!safeWorkbook.Contains(feature) && feature is not (Relationships + "worksheet" or Relationships + "styles" or
                Relationships + "sharedStrings" or Relationships + "theme" or Relationships + "calcChain"))
                throw Unsupported(sheet, "Remove the unsupported workbook feature or relationship before generation: " + feature.Split('/')[^1]);
        foreach (var name in inventory.Names)
            if ((string?)name.Attribute("name") != "_xlnm.Print_Area")
                throw Unsupported(sheet, "Remove defined names/print titles until their reference transformation is implemented.");
        var safeSheet = new HashSet<XName>(new[] { "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection",
            "mergeCells", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks" }.Select(name => S + name));
        foreach (var element in sheet.Xml.Root!.Elements())
            if (!safeSheet.Contains(element.Name))
                throw Unsupported(sheet, "Remove or re-author the unsupported worksheet structure: " + element.Name.LocalName);
        if (sheet.Xml.Descendants().Any(element => element.Name.Namespace != S || element.Name == S + "extLst"))
            throw Unsupported(sheet, "Remove unsupported nested extension/alternate-content metadata before cropping.");
        foreach (var feature in sheet.Features.Where(feature => feature.Contains("://", StringComparison.Ordinal)))
            if (feature != Relationships + "printerSettings")
                throw Unsupported(sheet, "Remove or re-author the unsupported worksheet relationship: " + feature.Split('/')[^1]);
        foreach (var row in sheet.Xml.Descendants(S + "row"))
            if (row.Elements().Any(element => element.Name != S + "c"))
                throw Unsupported(sheet, "Remove unsupported row extension metadata.");
        foreach (var cell in sheet.Cells.Where(cell => sheet.PrintArea.Contains(cell.Id.Address)))
        {
            if (cell.HasFormula) throw Unsupported(sheet, "Formula transformation is not implemented in M1. Use a static value until M2/M3.", cell.Id.Address.ToString());
            if (cell.Xml.Elements().Any(element => element.Name != S + "v" && element.Name != S + "is"))
                throw Unsupported(sheet, "Remove unsupported cell metadata.", cell.Id.Address.ToString());
            if (cell.NativeType is not (null or "n" or "b" or "s" or "inlineStr" or "str" or "e" or "d"))
                throw Unsupported(sheet, "Repair the invalid cell storage type.", cell.Id.Address.ToString());
            if (cell.Xml.Elements(S + "v").Skip(1).Any() || cell.Xml.Elements(S + "is").Skip(1).Any() ||
                cell.NativeType == "inlineStr" && (cell.Xml.Element(S + "is") is null || cell.Xml.Element(S + "v") is not null) ||
                cell.NativeType != "inlineStr" && cell.Xml.Element(S + "is") is not null)
                throw Unsupported(sheet, "Repair the inconsistent cell value/string storage.", cell.Id.Address.ToString());
            if (cell.NativeType is null or "n" && cell.Xml.Element(S + "v") is { Value.Length: > 0 } numeric &&
                (!double.TryParse(numeric.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number)))
                throw Unsupported(sheet, "Repair the invalid authored numeric cell value.", cell.Id.Address.ToString());
        }
    }

    private static void CropColumns(XElement settings, XlsxSheetGeometry geometry)
    {
        var columns = settings.Element(S + "cols");
        if (columns is null) return;
        foreach (var column in columns.Elements().ToArray())
        {
            var min = Integer(column.Attribute("min"), geometry, 1, 16384);
            var max = Integer(column.Attribute("max"), geometry, min, 16384);
            min = Math.Max(min, geometry.Source.PrintArea.First.Column);
            max = Math.Min(max, geometry.Source.PrintArea.Last.Column);
            if (min > max) { column.Remove(); continue; }
            column.SetAttributeValue("min", checked(min + geometry.ColumnOffset));
            column.SetAttributeValue("max", checked(max + geometry.ColumnOffset));
        }
        if (!columns.HasElements) columns.Remove();
    }

    private void CropMerges(XElement settings, XlsxSheetGeometry geometry)
    {
        var merges = settings.Element(S + "mergeCells");
        if (merges is null) return;
        var ranges = new List<XlsxRectangle>();
        foreach (var merge in merges.Elements().ToArray())
        {
            var range = ParseRange((string?)merge.Attribute("ref"), geometry);
            var image = ProjectRange(range, geometry);
            if (image.Kind == XlsxRangeImageKind.Empty) { merge.Remove(); continue; }
            if (!image.Complete) throw Unsupported(geometry.Source, "Move the merge wholly inside or outside PrintArea; a partial merge cannot survive cropping.", range.ToString());
            if (ranges.Any(prior => prior.First.Row <= range.Last.Row && range.First.Row <= prior.Last.Row &&
                prior.First.Column <= range.Last.Column && range.First.Column <= prior.Last.Column))
                throw Unsupported(geometry.Source, "Remove overlapping merged ranges.", range.ToString());
            ranges.Add(range);
            if (image.Kind == XlsxRangeImageKind.Cell) merge.Remove();
            else merge.SetAttributeValue("ref", image.Rectangle);
        }
        if (!merges.HasElements) merges.Remove();
        else merges.SetAttributeValue("count", merges.Elements().Count());
    }

    private void CropViews(XElement settings, XlsxSheetGeometry geometry)
    {
        if (settings.Element(S + "sheetPr")?.Attribute("syncRef") is { } synchronization)
        {
            try
            {
                if (geometry.RowOffset != 0 || geometry.ColumnOffset != 0)
                    throw Unsupported(geometry.Source, "Synchronization would move through crop.");
                MappedCell(synchronization.Value, geometry);
            }
            catch (TemplateException)
            {
                throw Unsupported(geometry.Source, "Remove the sheetPr/@syncRef synchronization setting or re-author it wholly inside an A1-based PrintArea; its relocation is not supported in M1.");
            }
        }
        foreach (var view in settings.Element(S + "sheetViews")?.Elements(S + "sheetView") ?? [])
        {
            if (view.Element(S + "pane") is { } pane && (geometry.RowOffset != 0 || geometry.ColumnOffset != 0))
                throw Unsupported(geometry.Source, "Remove the frozen/split pane or move PrintArea to A1 until pane relocation is implemented.");
            if (view.Attribute("topLeftCell") is { } topLeft) topLeft.Value = MappedCell(topLeft.Value, geometry);
            if (view.Element(S + "pane")?.Attribute("topLeftCell") is { } paneTop) paneTop.Value = MappedCell(paneTop.Value, geometry);
            foreach (var selection in view.Elements(S + "selection"))
            {
                if (selection.Attribute("pane") is not null && view.Element(S + "pane") is null)
                    throw Unsupported(geometry.Source, "Repair selection pane settings; the referenced pane is missing.");
                if (selection.Attribute("activeCell") is { } active) active.Value = MappedCell(active.Value, geometry);
                if (selection.Attribute("sqref") is { } ranges)
                    ranges.Value = string.Join(' ', ranges.Value.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(value =>
                    {
                        var range = ParseRange(value, geometry); var image = ProjectRange(range, geometry);
                        if (!image.Complete) throw Unsupported(geometry.Source, "Select cells wholly inside PrintArea or remove the selection setting.");
                        return image.Kind == XlsxRangeImageKind.Cell ? image.Rectangle!.Value.First.ToString() : image.Rectangle.ToString()!;
                    }));
            }
        }
    }

    private string MappedCell(string text, XlsxSheetGeometry geometry)
    {
        var range = ParseRange(text, geometry);
        var image = ProjectRange(range, geometry);
        if (range.First != range.Last || !image.Complete)
            throw Unsupported(geometry.Source, "Move the view/selection cell inside PrintArea or remove that setting.");
        return image.Rectangle!.Value.First.ToString();
    }

    private static void CropBreaks(XElement settings, XlsxSheetGeometry geometry, bool rows)
    {
        var container = settings.Element(S + (rows ? "rowBreaks" : "colBreaks"));
        if (container is null) return;
        var area = geometry.Source.PrintArea;
        foreach (var point in container.Elements().ToArray())
        {
            var id = Integer(point.Attribute("id"), geometry, 0, rows ? 1048576 : 16384);
            var min = Integer(point.Attribute("min"), geometry, 0, rows ? 16383 : 1048575);
            var max = Integer(point.Attribute("max"), geometry, min, rows ? 16383 : 1048575);
            var first = rows ? area.First.Row : area.First.Column;
            var last = rows ? area.Last.Row : area.Last.Column;
            min = Math.Max(min, (rows ? area.First.Column : area.First.Row) - 1);
            max = Math.Min(max, (rows ? area.Last.Column : area.Last.Row) - 1);
            if (id < first || id >= last || min > max) { point.Remove(); continue; }
            point.SetAttributeValue("id", checked(id + (rows ? geometry.RowOffset : geometry.ColumnOffset)));
            point.SetAttributeValue("min", checked(min + (rows ? geometry.ColumnOffset : geometry.RowOffset)));
            point.SetAttributeValue("max", checked(max + (rows ? geometry.ColumnOffset : geometry.RowOffset)));
        }
        if (!container.HasElements) container.Remove();
        else
        {
            container.SetAttributeValue("count", container.Elements().Count());
            container.SetAttributeValue("manualBreakCount", container.Elements().Count(element => (bool?)element.Attribute("man") == true));
        }
    }

    private static int Integer(XAttribute? value, XlsxSheetGeometry geometry, int min, int max) =>
        int.TryParse(value?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var number) && number >= min && number <= max
            ? number : throw Unsupported(geometry.Source, "Repair invalid row/column/print-setting coordinates.");

    private static XlsxRectangle ParseRange(string? text, XlsxSheetGeometry geometry)
    {
        var parts = text?.Split(':') ?? [];
        if (parts.Length is < 1 or > 2) throw Unsupported(geometry.Source, "Use a finite A1 cell or rectangular range.");
        var first = ParseAddress(parts[0], geometry);
        var last = parts.Length == 2 ? ParseAddress(parts[1], geometry) : first;
        if (first.Row > last.Row || first.Column > last.Column) throw Unsupported(geometry.Source, "Use increasing rectangular range coordinates.");
        return new(first, last);
    }

    private static XlsxAddress ParseAddress(string text, XlsxSheetGeometry geometry)
    {
        var match = CellAddress().Match(text);
        if (!match.Success) throw Unsupported(geometry.Source, "Use finite A1 coordinates for worksheet settings.");
        var column = 0;
        foreach (var character in match.Groups["column"].Value) column = checked(column * 26 + character - 'A' + 1);
        if (!int.TryParse(match.Groups["row"].Value, out var row) || row is < 1 or > 1048576 || column is < 1 or > 16384)
            throw Unsupported(geometry.Source, "Repair worksheet coordinates outside Excel dimensions.");
        return new(row, column);
    }

    private static XlsxAddress Move(XlsxAddress address, XlsxSheetGeometry geometry) =>
        new(checked(address.Row + geometry.RowOffset), checked(address.Column + geometry.ColumnOffset));
    private static string SheetPrefix(string name) => "'" + name.Replace("'", "''", StringComparison.Ordinal) + "'!";
    private static string Absolute(XlsxRectangle area) => AbsoluteCell(area.First) + ":" + AbsoluteCell(area.Last);
    private static string AbsoluteCell(XlsxAddress address)
    {
        var text = address.ToString();
        var offset = 0;
        while (offset < text.Length && char.IsLetter(text[offset])) offset++;
        return "$" + text[..offset] + "$" + text[offset..];
    }
    private static TemplateException Unsupported(XlsxSheetSource sheet, string explanation, string? address = null) =>
        new StructuralDiagnostic(DiagnosticCodes.XlsxUnsupportedFeature, explanation,
            Location: sheet.Part + ", sheet '" + sheet.Name + "'" + (address is null ? string.Empty : ", " + address), Explanation: explanation).ToException();

    [GeneratedRegex(@"^\$?(?<column>[A-Z]{1,3})\$?(?<row>[1-9][0-9]*)$", RegexOptions.CultureInvariant)]
    private static partial Regex CellAddress();
}
