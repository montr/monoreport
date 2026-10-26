using System.Collections.Immutable;
using System.Xml.Linq;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxGeometryPlanner
{
    internal XlsxSheetOutput EmitCells(XlsxSheetTemplate sheet, XlsxSheetGeometry geometry)
    {
        var settings = new XElement(sheet.Source.Xml.Root!);
        settings.Element(S + "sheetData")!.RemoveNodes();
        if (sheet.Source.PrintArea is not null)
        {
            settings.Element(S + "dimension")?.Remove();
            var dimension = new XElement(S + "dimension", new XAttribute("ref", geometry.OutputArea));
            if (settings.Element(S + "sheetPr") is { } properties) properties.AddAfterSelf(dimension); else settings.AddFirst(dimension);
        }
        var columns = settings.Element(S + "cols");
        if (columns is not null)
        {
            var source = columns.Elements().Select(element => new XElement(element)).ToArray(); columns.RemoveNodes();
            foreach (var span in geometry.ColumnSpans)
                foreach (var column in source)
                {
                    var first = Math.Max(Integer(column.Attribute("min"), geometry, 1, 16384), span.First);
                    var last = Math.Min(Integer(column.Attribute("max"), geometry, 1, 16384), span.Last);
                    if (first > last) continue;
                    var copy = new XElement(column);
                    copy.SetAttributeValue("min", span.OutputFirst + first - span.First);
                    copy.SetAttributeValue("max", span.OutputFirst + last - span.First); columns.Add(copy);
                }
            if (!columns.HasElements) columns.Remove();
        }
        var byRow = sheet.Cells.GroupBy(cell => cell.Source.Id.Address.Row).ToDictionary(group => group.Key, group => group.ToArray());
        var sourceRows = sheet.Source.Xml.Root!.Element(S + "sheetData")!.Elements(S + "row").ToArray();
        var rows = ImmutableArray.CreateBuilder<XlsxRowPlan>();
        foreach (var span in geometry.RowSpans)
            foreach (var row in sourceRows)
            {
                var sourceRow = (int)row.Attribute("r")!;
                if (sourceRow < span.First || sourceRow > span.Last) continue;
                var copy = new XElement(row); copy.RemoveNodes(); copy.Attribute("spans")?.Remove();
                var outputRow = span.OutputFirst + sourceRow - span.First;
                var cells = ImmutableArray.CreateBuilder<XlsxCellPlacement>();
                foreach (var column in geometry.ColumnSpans)
                    foreach (var cell in byRow.GetValueOrDefault(sourceRow) ?? [])
                    {
                        var sourceColumn = cell.Source.Id.Address.Column;
                        if (sourceColumn < column.First || sourceColumn > column.Last) continue;
                        cells.Add(new(new(cell.Source.Id, span.Vertical, column.Horizontal),
                            new(outputRow, column.OutputFirst + sourceColumn - column.First), cell.Source, null));
                    }
                rows.Add(new(copy, outputRow, cells.ToImmutable()));
            }
        return new(geometry, settings, rows.ToImmutable());
    }
}
