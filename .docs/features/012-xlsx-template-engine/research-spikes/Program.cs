using System.Text.Json;
using ClosedXML.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

// Research only: reference mapping is supplied by a fixture, not a template engine.
if (args[0] == "validate")
{
    foreach (var path in args.Skip(1))
    {
        using var workbook = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2010).Validate(workbook).ToArray();
        Console.WriteLine($"{Path.GetFileName(path)}: {errors.Length} schema errors");
        foreach (var error in errors.Take(8)) Console.WriteLine(error.Description + " " + error.Path?.XPath);
        if (errors.Length != 0) Environment.ExitCode = 1;
    }
    return;
}
if (args[0] == "transform")
{
    var requests = JsonSerializer.Deserialize<List<Request>>(File.ReadAllText(args[1]))!;
    var results = requests.Select(request => FormulaConverter.ModifyA1(request.Formula, "Report", 1, 1,
        new MapVisitor(request.Rows, request.Cols))).ToArray();
    File.WriteAllText(args[2], JsonSerializer.Serialize(results));
    Console.WriteLine($"Transformed {results.Length} fixture formulas through parsed reference nodes.");
    return;
}

var checks = new List<object>();
foreach (var formula in new[] { "SUM(D8:D9)", "COUNT(D8:D9)", "SUBTOTAL(9,D8:D9)",
    "'Other Sheet'!$B$2+D8", "SUM(DetailValues)", "SUM(Table1[Amount])", "SUM(D:D)",
    "SUM(8:9)", "SUM(Report:Other!D8)", "'[1]Other'!D8", "INDIRECT(\"D8\")", "OFFSET(D8,1,0)",
    "IF(D8>0,\"D8 literal\",D9)", "_xlfn.XLOOKUP(D8,D8:D9,E8:E9)", "D8#", "@D8",
    "MYFUNCTION(D8)", "SUM((D8:D9,D11:D12))" })
{
    var visitor = new ProbeVisitor();
    try
    {
        var result = FormulaConverter.ModifyA1(formula, "Report", 1, 1, visitor);
        checks.Add(new { formula, parsed = true, result, nodes = visitor.Nodes });
    }
    catch (Exception error) { checks.Add(new { formula, parsed = false, error = error.GetType().Name }); }
}
var copied = FormulaConverter.ToA1(FormulaConverter.ToR1C1("D8+$D$8+D$8+$D8", 8, 5), 12, 7);
if (copied != "F12+$D$8+F$8+$D12") throw new Exception("Mixed-reference copy failed: " + copied);
var changed = FormulaConverter.ModifyA1("IF(D8>0,\"D8 literal\",D9)", "Report", 1, 1,
    new MapVisitor(new() { [8] = [12], [9] = [13] }, new() { [4] = [6] }));
if (changed != "IF(F12>0,\"D8 literal\",F13)") throw new Exception(changed);
// Absolute references move for structural relocation, even though they do not move on copy.
var moved = FormulaConverter.ModifyA1("'Other Sheet'!$B$2+$D$8", "Report", 1, 1,
    new MapVisitor(new() { [8] = [12] }, new() { [4] = [6] }));
if (moved != "'Other Sheet'!$B$2+$F$12") throw new Exception(moved);
File.WriteAllText(args[1], JsonSerializer.Serialize(new { copied, changed, moved, checks }, new JsonSerializerOptions { WriteIndented = true }));
try
{
    FormulaConverter.ModifyA1("SUM(D8:D9)", "Report", 1, 1,
        new MapVisitor(new() { [8] = [8, 10, 14], [9] = [9, 11, 15] }, new()));
    throw new Exception("Unsafe bounding rectangle was accepted.");
}
catch (InvalidOperationException) { Console.WriteLine("Noncontiguous mapped detail range rejected."); }
try
{
    FormulaConverter.ModifyA1("D8", "Report", 1, 1,
        new MapVisitor(new() { [8] = [8, 9] }, new()));
    throw new Exception("Ambiguous scalar expansion was accepted.");
}
catch (InvalidOperationException) { Console.WriteLine("Ambiguous single-cell expansion rejected."); }
Console.WriteLine($"Parser probe: {checks.Count} forms recorded; copy, literal isolation, structural absolute relocation asserted.");

record Request(string Formula, Dictionary<int, int[]> Rows, Dictionary<int, int[]> Cols);

sealed class MapVisitor(Dictionary<int, int[]> rows, Dictionary<int, int[]> cols) : CopyVisitor
{
    public override TransformedSymbol Reference(ModContext ctx, SymbolRange range, ReferenceArea reference)
    {
        var first = reference.First;
        var second = reference.Second;
        int[] Positions(Dictionary<int, int[]> map, int start, int end) =>
            Enumerable.Range(Math.Min(start, end), Math.Abs(end - start) + 1)
                .SelectMany(x => map.TryGetValue(x, out var positions) ? positions : [x]).ToArray();
        var rr = Positions(rows, first.RowValue, second.RowValue);
        var cc = Positions(cols, first.ColumnValue, second.ColumnValue);
        if (!ctx.Formula.AsSpan(range.Start, range.Length).Contains(':') &&
            (rr.Distinct().Count() != 1 || cc.Distinct().Count() != 1))
            throw new InvalidOperationException("Ambiguous single-cell target after repetition.");
        if (rr.Distinct().Count() != rr.Max() - rr.Min() + 1 ||
            cc.Distinct().Count() != cc.Max() - cc.Min() + 1)
            throw new InvalidOperationException("Noncontiguous image: reference requires a product policy.");
        var mapped = new ReferenceArea(
            new RowCol(first.RowType, rr.Min(), first.ColumnType, cc.Min(), ReferenceStyle.A1),
            new RowCol(second.RowType, rr.Max(), second.ColumnType, cc.Max(), ReferenceStyle.A1));
        return base.Reference(ctx, range, mapped);
    }
    // Cross-sheet references deliberately remain unchanged in this local-sheet mapping probe.
}

sealed class ProbeVisitor : CopyVisitor
{
    public List<string> Nodes { get; } = [];
    public override TransformedSymbol Reference(ModContext ctx, SymbolRange range, ReferenceArea reference)
    { Nodes.Add("local:" + reference.GetDisplayStringA1()); return base.Reference(ctx, range, reference); }
    public override TransformedSymbol SheetReference(ModContext ctx, SymbolRange range, string sheet, ReferenceArea reference)
    { Nodes.Add("sheet:" + sheet); return base.SheetReference(ctx, range, sheet, reference); }
    public override TransformedSymbol Name(ModContext ctx, SymbolRange range, string name)
    { Nodes.Add("name:" + name); return base.Name(ctx, range, name); }
    public override TransformedSymbol StructureReference(ModContext ctx, SymbolRange range, string table,
        StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    { Nodes.Add("table:" + table); return base.StructureReference(ctx, range, table, area, firstColumn, lastColumn); }
    public override TransformedSymbol Function(ModContext ctx, SymbolRange range, ReadOnlySpan<char> name,
        IReadOnlyList<TransformedSymbol> arguments)
    { Nodes.Add("function:" + name.ToString()); return base.Function(ctx, range, name, arguments); }
}
