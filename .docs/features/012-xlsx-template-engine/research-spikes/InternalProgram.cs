using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var scanner = new InternalReferences();
if (args[0] == "validate")
{
    foreach (var path in args.Skip(1))
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2010).Validate(doc).ToArray();
        Console.WriteLine(Path.GetFileName(path) + ": " + errors.Length + " schema errors");
        foreach (var error in errors.Take(5)) Console.WriteLine(error.Description);
        if (errors.Length != 0) Environment.ExitCode = 1;
    }
    return;
}
if (args[0] == "transform")
{
    var requests = JsonSerializer.Deserialize<List<Request>>(File.ReadAllText(args[1]))!;
    var result = requests.Select(request => scanner.Rewrite(request.Formula, token => Map(token, request))).ToArray();
    File.WriteAllText(args[2], JsonSerializer.Serialize(result));
    Console.WriteLine("Internal transformer: " + result.Length + " supplied-plan formulas transformed.");
    return;
}
var old = JsonDocument.Parse(File.ReadAllText(args[1]));
var cases = old.RootElement.GetProperty("checks").EnumerateArray().Select(c => c.GetProperty("formula").GetString()!).Concat(new[] {
    "LOG10(D8)+1E10+D9", "IF(D8>0,\"A1 \"\"D9\"\"\",D9)", "'O''Brien'!$D8+SUM(D8:E9)",
    "SUM($D$8:$E9)+PI()", "D8+UnresolvedName", "SUM(D8", "\"unclosed", "XFE1+D8", "[1]Other!D8"
}).ToArray();
var results = new List<object>();
foreach (var formula in cases)
{
    try
    {
        var rewritten = scanner.Rewrite(formula, token =>
        {
            CellPoint Shift(CellPoint p) => p with { Row = p.Row + (p.RowAbsolute ? 0 : 4), Col = p.Col + (p.ColAbsolute ? 0 : 2) };
            return token.Prefix + Shift(token.First) + (token.IsRange ? ":" + Shift(token.Last) : "");
        }, new HashSet<string>(["DetailValues"], StringComparer.OrdinalIgnoreCase));
        results.Add(new { formula, supported = true, rewritten });
    }
    catch (Exception error) when (error is NotSupportedException or FormatException)
    { results.Add(new { formula, supported = false, reason = error.Message }); }
}
var copy = scanner.Rewrite("D8+$D$8+D$8+$D8", t =>
{
    var p = t.First with { Row = t.First.Row + (t.First.RowAbsolute ? 0 : 4), Col = t.First.Col + (t.First.ColAbsolute ? 0 : 2) };
    return t.Prefix + p;
});
if (copy != "F12+$D$8+F$8+$D12") throw new Exception("Copy flags failed.");
var lexical = scanner.Rewrite("LOG10(D8)+1E10+\"D8\"", t => (t.First with { Row = 12, Col = 6 }).ToString());
if (lexical != "LOG10(F12)+1E10+\"D8\"") throw new Exception("Lexical isolation failed.");
File.WriteAllText(args[2], JsonSerializer.Serialize(new { copy, lexical, cases = results }, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine("Internal probe: " + cases.Length + " supported/rejected cases recorded; lexical isolation and mixed-reference assertions passed.");

static string Map(ReferenceToken token, Request request)
{
    if (token.Prefix.Length != 0) return request.Formula.Substring(token.Start, token.Length); // local-only fixture map
    int[] Positions(Dictionary<int, int[]> map, int a, int b) => Enumerable.Range(Math.Min(a, b), Math.Abs(a - b) + 1).SelectMany(i => map.TryGetValue(i, out var outputs) ? outputs : [i]).Distinct().Order().ToArray();
    var rows = Positions(request.Rows, token.First.Row, token.Last.Row);
    var cols = Positions(request.Cols, token.First.Col, token.Last.Col);
    if (rows.Length == 0 || cols.Length == 0) throw new InvalidOperationException("Empty reference image.");
    if (!token.IsRange && (rows.Length != 1 || cols.Length != 1)) throw new InvalidOperationException("Ambiguous single-cell reference image.");
    if (rows.Length != rows[^1] - rows[0] + 1 || cols.Length != cols[^1] - cols[0] + 1) throw new InvalidOperationException("Disjoint reference image.");
    return (token.First with { Row = rows[0], Col = cols[0] }).ToString() + (token.IsRange ? ":" + (token.Last with { Row = rows[^1], Col = cols[^1] }) : "");
}
record Request(string Formula, Dictionary<int, int[]> Rows, Dictionary<int, int[]> Cols);
