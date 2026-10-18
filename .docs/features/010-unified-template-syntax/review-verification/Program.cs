using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MonoReport.Models;
using MonoReport.Services;

if (args.Length != 1) throw new ArgumentException("Supply an output directory outside the repository.");
var outputDirectory = Path.GetFullPath(args[0]);
var repository = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../../../../"));
if (outputDirectory == repository || outputDirectory.StartsWith(repository + Path.DirectorySeparatorChar, StringComparison.Ordinal))
    throw new ArgumentException("Generated evidence must be outside the repository.");
Directory.CreateDirectory(outputDirectory);
XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
var reader = new DocxStoryReader();
var analyzer = new TemplateAnalyzer(reader, new DocxTemplatePlanner(new TemplateLanguageParser()));
var engine = new MergeEngine(new DocumentRenderer(new DataResolver(), new ValueFormatter(new RubSpellOutFormatter())), reader, analyzer);
var failures = new List<string>();
var observations = new List<object>();
var originalRejections = new Dictionary<string, string>
{
    ["crossing-bookmark"] = "template.invalid_structure",
    ["external-bookmark-link"] = "template.invalid_structure",
    ["nested-bookmark-link"] = "template.invalid_structure",
    ["tracked-cell-deletion"] = "template.tracked_revision",
    ["tracked-move-range"] = "template.tracked_revision",
    ["expression-crosses-content-control-exit"] = "template.unmatched_brace"
};
foreach (var (name, code) in originalRejections)
    Run("originals", name, expectedCode: code);
Run("originals", "comment-with-image");
foreach (var name in new[] { "zero-padded-comment-id", "external-relationship-with-anchor", "cloned-external-relationship-with-anchor", "field-begins-in-control", "field-ends-in-control" })
    Run("", name);
foreach (var (suffix, json) in new[] { ("missing", "{}"), ("zero", "{\"Items\":[]}"), ("one", "{\"Items\":[\"A\"]}") })
    Run("", "zero-padded-comment-id", suffix: suffix, jsonOverride: json);

// The same services must remain usable after rejected and corrupting operations.
Run("originals", "comment-with-image", suffix: "after-probes");
var concurrent = await Task.WhenAll(Enumerable.Range(0, 8).Select(_ => Task.Run(() =>
{
    var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "originals/comment-with-image.template.docx"));
    using var input = new MemoryStream(bytes); using var data = JsonDocument.Parse("{\"Items\":[\"A\",\"B\"]}");
    return Validate(engine.Merge(input, data).Document).Length == 0 && bytes.SequenceEqual(input.ToArray());
})));
Check(concurrent.All(value => value), "operation-isolation", "Eight concurrent comment-image merges must validate and preserve input");
File.WriteAllText(Path.Combine(outputDirectory, "results.json"), JsonSerializer.Serialize(new { observations, failures, concurrentOperations = concurrent.Length }, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Independent contract failures: {failures.Count}; concurrent operations: {concurrent.Length}");
return failures.Count == 0 ? 0 : 1;

void Run(string folder, string name, string? expectedCode = null, string? suffix = null, string? jsonOverride = null)
{
    var label = name + (suffix is null ? "" : "-" + suffix);
    var path = Path.Combine(AppContext.BaseDirectory, folder, name);
    var bytes = File.ReadAllBytes(path + ".template.docx");
    var sourceErrors = Validate(bytes);
    if (sourceErrors.Length != 0) throw new InvalidDataException(label + ": invalid source: " + string.Join("; ", sourceErrors));
    using var input = new MemoryStream(bytes);
    var analysis = analyzer.Analyze(input);
    File.WriteAllText(Path.Combine(outputDirectory, label + ".analysis.json"), JsonSerializer.Serialize(analysis));
    input.Position = 0;
    Check(JsonSerializer.Serialize(analysis) == JsonSerializer.Serialize(analyzer.Analyze(input)), label, "Analysis must be deterministic");
    Check(analysis.Valid == (expectedCode is null), label, "Analyze validity must follow the approved contract");
    if (expectedCode is not null) Check(analysis.Errors.Any(e => e.Code == expectedCode), label, "Expected structural diagnostic");
    if (name.StartsWith("field-", StringComparison.Ordinal))
        Check(analysis.Occurrences.Select(e => e.Expression).SequenceEqual(new[] { "{Visible}" }), label, "Only the ordinary expression after the opaque field is a dependency");
    input.Position = 0;
    var json = jsonOverride ?? File.ReadAllText(path + ".json");
    using var data = JsonDocument.Parse(json);
    try
    {
        var result = engine.Merge(input, data);
        Check(expectedCode is null, label, "Rejected templates must not return any document");
        File.WriteAllBytes(Path.Combine(outputDirectory, label + ".output.docx"), result.Document);
        var outputErrors = Validate(result.Document);
        Check(outputErrors.Length == 0, label, "Generated package must pass Office 2010 validation");
        using var document = WordprocessingDocument.Open(new MemoryStream(result.Document), false);
        using var mainStream = document.MainDocumentPart!.GetStream();
        var main = XDocument.Load(mainStream);
        if (name.Contains("comment", StringComparison.Ordinal))
        {
            var package = Parts(document).Where(p => p.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase))
                .Select(p => { using var s = p.GetStream(); return XDocument.Load(s); }).ToArray();
            var refs = main.Descendants(w + "commentReference").Select(e => (string)e.Attribute(w + "id")!).ToArray();
            var definitions = package.SelectMany(x => x.Descendants(w + "comment")).ToArray();
            Check(refs.Length == refs.Distinct().Count(), label, "Every surviving attachment must have a distinct ID");
            Check(refs.All(id => definitions.Count(e => (string?)e.Attribute(w + "id") == id) == 1), label, "Every attachment must resolve to exactly one definition");
            var drawingIds = package.SelectMany(x => x.Descendants(wp + "docPr")).Select(e => (string)e.Attribute("id")!).ToArray();
            Check(drawingIds.Length == drawingIds.Distinct().Count(), label, "Drawing IDs must be unique across all physical definitions/parts");
            if (name == "comment-with-image")
            {
                Check(refs.Length == 2 && drawingIds.Length == 3, label, "Two comment copies and the original drawing must survive");
                using var original = WordprocessingDocument.Open(new MemoryStream(bytes), false);
                var originalPart = original.MainDocumentPart!.WordprocessingCommentsPart!;
                var generatedPart = document.MainDocumentPart.WordprocessingCommentsPart!;
                using var originalXml = originalPart.GetStream(); var definitionXml = XDocument.Load(originalXml);
                using var generatedXml = generatedPart.GetStream(); var generatedDefinitions = XDocument.Load(generatedXml);
                var sourceText = Text(definitionXml);
                Check(definitions.All(d => Text(d) == sourceText), label, "Retained and cloned comments must preserve text");
                foreach (var blip in generatedDefinitions.Descendants(a + "blip"))
                {
                    var relation = (string)blip.Attribute(r + "embed")!;
                    Check(Read(originalPart.GetPartById(relation)).SequenceEqual(Read(generatedPart.GetPartById(relation))), label, "Image relationship must resolve to unchanged bytes");
                }
            }
        }
        if (name.Contains("external-relationship", StringComparison.Ordinal))
        {
            var links = main.Descendants(w + "hyperlink").ToArray();
            Check(links.All(link => (string?)link.Attribute(w + "anchor") == "Target"), label, "External bookmark fragment must remain Target");
            Check(links.All(link => document.MainDocumentPart.HyperlinkRelationships.Any(rel => rel.Id == (string?)link.Attribute(r + "id") && rel.IsExternal && rel.Uri.ToString() == "https://example.org/report.docx")), label, "External relationship must remain intact");
        }
        if (name.StartsWith("field-", StringComparison.Ordinal))
        {
            Check(Text(main) == (name == "field-begins-in-control" ? "{Name}OK" : "cachedOK"), label, "Field cache must remain literal and the following ordinary expression must render");
            using var original = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            using var s = original.MainDocumentPart!.GetStream(); var source = XDocument.Load(s);
            Check(XNode.DeepEquals(source.Descendants(w + "sdt").Single(), main.Descendants(w + "sdt").Single()), label, "Complete content control must remain unchanged");
        }
        observations.Add(new { label, sourceErrors = 0, analysis.Valid, outcome = "generated", outputErrors });
        Console.WriteLine($"{label}: source valid; Analyze={analysis.Valid}; generated; validator errors={outputErrors.Length}");
    }
    catch (TemplateException error)
    {
        Check(expectedCode is not null, label, "Supported template must generate successfully");
        Check(analysis.Errors.Take(100).Select(e => (e.Code, e.Part, e.Location, e.Message)).SequenceEqual(error.Errors.Select(e => (e.Code, e.Part, e.Location, e.Message))), label, "Analyze/Generate diagnostics must match in order");
        observations.Add(new { label, sourceErrors = 0, analysis.Valid, outcome = "rejected", diagnostics = error.Errors });
        Console.WriteLine($"{label}: source valid; Analyze={analysis.Valid}; rejected; errors={error.Errors.Count}");
    }
    Check(input.CanRead && bytes.SequenceEqual(input.ToArray()), label, "Caller input must remain open and unchanged");
}

void Check(bool condition, string label, string requirement)
{
    if (condition) return;
    var message = label + ": " + requirement;
    failures.Add(message); Console.WriteLine("FAIL " + message);
}
string Text(XContainer xml) => string.Concat(xml.Descendants(w + "t").Select(e => e.Value));
string[] Validate(byte[] bytes)
{
    using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
    return new OpenXmlValidator(FileFormatVersions.Office2010).Validate(document)
        .Select(e => e.Description + " " + e.Part?.Uri + " " + e.Path?.XPath).ToArray();
}
byte[] Read(OpenXmlPart part) { using var s = part.GetStream(); using var b = new MemoryStream(); s.CopyTo(b); return b.ToArray(); }
IEnumerable<OpenXmlPart> Parts(WordprocessingDocument document)
{
    var pending = new Queue<OpenXmlPart>(document.Parts.Select(p => p.OpenXmlPart)); var seen = new HashSet<OpenXmlPart>();
    while (pending.TryDequeue(out var part))
    {
        if (!seen.Add(part)) continue;
        foreach (var child in part.Parts) pending.Enqueue(child.OpenXmlPart);
        yield return part;
    }
}
