using System.Collections.Immutable;
using System.Globalization;
using System.Xml.Linq;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxFormulaReferences
{
    internal IReadOnlyDictionary<SourceCellId, XlsxParsedFormula> ReadSharedCopies(XlsxSourceInventory source,
        ImmutableArray<XlsxSharedFormulaGroup> groups, IReadOnlyDictionary<(uint, string), XlsxDefinedName> names,
        IReadOnlyDictionary<string, uint> sheetIds, XlsxOperationContext context)
    {
        var result = new Dictionary<SourceCellId, XlsxParsedFormula>();
        var sheets = source.Sheets.ToDictionary(sheet => sheet.Id);
        foreach (var group in groups)
        {
            if (context.LimitReached) break;
            var sheet = sheets[group.Master.SheetId];
            var retained = group.Members.Where(member => sheet.RetentionArea.Contains(member.Address)).ToArray();
            if (retained.Length == 0) continue;
            string Location(SourceCellId cell) => sheet.Part + ", sheet '" + sheet.Name + "', " + cell.Address;
            XlsxParsedFormula Parse(string text, string location) => ResolveSheets(Read(text, location,
                (prefix, name) => ResolveName(names, sheetIds, sheet.Id, prefix, name, location)), sheet.Id, sheetIds, location);
            XlsxParsedFormula master;
            try { master = Parse(group.Formula, Location(group.Master)); }
            catch (TemplateException error) { context.Report(error); continue; }
            foreach (var member in retained)
            {
                if (context.LimitReached) break;
                try
                {
                    // Copy in source coordinates only. Reparse changed spans before later owner relocation.
                    result.Add(member, member == group.Master ? master : Parse(Copy(master,
                        member.Address.Row - group.Master.Address.Row, member.Address.Column - group.Master.Address.Column,
                        Location(member)), Location(member)));
                }
                catch (TemplateException error) { context.Report(error); }
            }
        }
        context.ThrowIfErrors();
        return result;
    }

    internal ImmutableArray<XlsxSharedFormulaGroup> ReadSharedFormulas(XlsxSourceInventory source, XlsxOperationContext context)
    {
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var result = ImmutableArray.CreateBuilder<XlsxSharedFormulaGroup>();
        foreach (var sheet in source.Sheets)
        {
            // Inventory the authored sheet before crop; a retained follower can have a disposable master.
            var members = new List<(XlsxCellSource Cell, XElement Formula, uint Index)>();
            string Location(XlsxCellSource cell) => sheet.Part + ", sheet '" + sheet.Name + "', " + cell.Id.Address;
            foreach (var cell in sheet.Cells)
            {
                if (context.LimitReached) break;
                var formulas = cell.Xml.Elements(ns + "f").ToArray();
                if (!formulas.Any(formula => (string?)formula.Attribute("t") == "shared" || formula.Attribute("si") is not null)) continue;
                try
                {
                    if (formulas.Length != 1)
                        throw Unsafe(Location(cell), "A shared formula cell must contain exactly one formula element. Re-save it as an ordinary scalar A1 formula.");
                    var formula = formulas[0];
                    if ((string?)formula.Attribute("t") != "shared" ||
                        !uint.TryParse((string?)formula.Attribute("si"), NumberStyles.None, CultureInfo.InvariantCulture, out var index))
                        throw Unsafe(Location(cell), "A shared formula requires t=shared and a valid unsigned si group index on the same worksheet. Repair the group or re-save it as ordinary scalar A1 formulas.");
                    if (formula.HasElements || formula.Attributes().Any(attribute =>
                        attribute.Name != "t" && attribute.Name != "si" && attribute.Name != "ref" && attribute.Name != "ca" && attribute.Name != XNamespace.Xml + "space") ||
                        formula.Attribute("ca") is { Value: not ("0" or "1" or "true" or "false") } ||
                        formula.Attribute(XNamespace.Xml + "space") is { Value: not ("default" or "preserve") })
                        throw Unsafe(Location(cell), "Remove unsupported shared-formula metadata or special formula structures; re-save the group as ordinary scalar A1 formulas.");
                    members.Add((cell, formula, index));
                }
                catch (TemplateException error) { context.Report(error); }
            }
            context.ThrowIfErrors();
            var sheetGroups = new List<XlsxSharedFormulaGroup>();
            foreach (var group in members.GroupBy(member => member.Index))
            {
                if (context.LimitReached) break;
                try
                {
                    // ref identifies the master. Followers may redundantly store text, which never overrides it.
                    var masters = group.Where(member => member.Formula.Attribute("ref") is not null).ToArray();
                    if (masters.Length == 0)
                        throw Unsafe(Location(group.First().Cell), "Shared formula group " + group.Key + " has no master with a ref range. Supply one master on this worksheet or re-save the group as ordinary scalar A1 formulas.");
                    if (masters.Length > 1)
                        throw Unsafe(Location(masters[1].Cell), "Shared formula group " + group.Key + " has conflicting masters. Keep exactly one master with a ref range or re-save the group as ordinary scalar A1 formulas.");
                    var master = masters[0];
                    if (string.IsNullOrWhiteSpace(master.Formula.Value))
                        throw Unsafe(Location(master.Cell), "The shared formula master has no scalar formula text. Supply its formula or re-save the group as ordinary scalar A1 formulas.");
                    var range = SharedRange(master.Formula.Attribute("ref")!.Value, Location(master.Cell));
                    foreach (var member in group)
                        if (!range.Contains(member.Cell.Id.Address))
                            throw Unsafe(Location(member.Cell), "Shared formula group " + group.Key + " member lies outside its master ref " + range + ". Correct its group membership/range or re-save it as an ordinary scalar A1 formula.");
                    if (sheetGroups.Any(prior => prior.Range.First.Row <= range.Last.Row && range.First.Row <= prior.Range.Last.Row &&
                        prior.Range.First.Column <= range.Last.Column && range.First.Column <= prior.Range.Last.Column))
                        throw Unsafe(Location(master.Cell), "Shared formula master ranges overlap on this worksheet. Use nonoverlapping groups or ordinary scalar A1 formulas.");
                    // A rectangle is a membership bound, not a dense cell grid: ordinary overrides and holes stay independent.
                    sheetGroups.Add(new(group.Key, master.Cell.Id, range, master.Formula.Value,
                        group.Select(member => member.Cell.Id).ToImmutableArray()));
                }
                catch (TemplateException error) { context.Report(error); }
            }
            context.ThrowIfErrors();
            result.AddRange(sheetGroups);
        }
        return result.ToImmutable();
    }

    private static XlsxRectangle SharedRange(string text, string location)
    {
        if (!TryPoint(text, 0, out var first, out var after))
            throw Unsafe(location, "Use a finite worksheet-local A1 cell/rectangle for the shared master ref range.");
        var last = first;
        if (after < text.Length && text[after] == ':')
        {
            if (!TryPoint(text, after + 1, out last, out after))
                throw Unsafe(location, "Use a finite worksheet-local A1 endpoint for the shared master ref range.");
        }
        if (after != text.Length || last!.Address.Row < first!.Address.Row || last.Address.Column < first.Address.Column)
            throw Unsafe(location, "Use one increasing finite worksheet-local A1 cell/rectangle for the shared master ref range.");
        return new(first!.Address, last!.Address);
    }
}
