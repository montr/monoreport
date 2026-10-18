using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MonoReport.Services;

if (args.Length != 1)
    throw new ArgumentException("Supply an output directory outside the repository.");

var outputDirectory = Path.GetFullPath(args[0]);
Directory.CreateDirectory(outputDirectory);
var reader = new DocxStoryReader();
var analyzer = new TemplateAnalyzer(reader, new DocxTemplatePlanner(new TemplateLanguageParser()));
var engine = new MergeEngine(
    new DocumentRenderer(new DataResolver(), new ValueFormatter(new RubSpellOutFormatter())), reader, analyzer);
XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
var cases = new[]
{
    "crossing-bookmark", "comment-with-image", "external-bookmark-link", "nested-bookmark-link",
    "tracked-cell-deletion", "tracked-move-range", "expression-crosses-content-control-exit"
};
var defectObserved = false;
foreach (var name in cases)
{
    var template = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, name + ".template.docx"));
    var sourceErrors = Validate(template);
    if (sourceErrors.Length != 0)
        throw new InvalidDataException(name + ": source package is invalid: " + string.Join("; ", sourceErrors));
    using var input = new MemoryStream(template);
    var analysis = analyzer.Analyze(input);
    File.WriteAllText(Path.Combine(outputDirectory, name + ".analysis.json"), JsonSerializer.Serialize(analysis));
    Console.WriteLine($"{name}: source valid; Analyze.Valid={analysis.Valid}");
    input.Position = 0;
    using var data = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, name + ".json")));
    var result = engine.Merge(input, data);
    File.WriteAllBytes(Path.Combine(outputDirectory, name + ".output.docx"), result.Document);
    var outputErrors = Validate(result.Document);
    foreach (var error in outputErrors) Console.WriteLine("  Open XML: " + error);
    using var output = new MemoryStream(result.Document);
    using var document = WordprocessingDocument.Open(output, false);
    using var mainStream = document.MainDocumentPart!.GetStream();
    var xml = XDocument.Load(mainStream);
    var names = xml.Descendants(w + "bookmarkStart").Select(item => (string?)item.Attribute(w + "name")).ToHashSet();
    var dangling = xml.Descendants(w + "hyperlink").Select(item => (string?)item.Attribute(w + "anchor"))
        .Where(anchor => anchor is not null && !names.Contains(anchor)).ToArray();
    var observed = false;
    switch (name)
    {
        case "crossing-bookmark":
            var ends = xml.Descendants(w + "bookmarkEnd").Select(item => (string?)item.Attribute(w + "id")).ToArray();
            observed = ends.Length != ends.Distinct().Count() && outputErrors.Length != 0;
            Console.WriteLine("  Bookmark end IDs: " + string.Join(", ", ends));
            break;
        case "comment-with-image":
            using (var commentsStream = document.MainDocumentPart.WordprocessingCommentsPart!.GetStream())
            {
                var comments = XDocument.Load(commentsStream);
                var ids = comments.Descendants(wp + "docPr").Select(item => (string?)item.Attribute("id")).ToArray();
                observed = ids.Length != ids.Distinct().Count() && outputErrors.Length != 0;
                Console.WriteLine("  Comment drawing IDs: " + string.Join(", ", ids));
            }
            break;
        case "external-bookmark-link":
        case "nested-bookmark-link":
            observed = dangling.Length != 0;
            Console.WriteLine("  Dangling anchors: " + string.Join(", ", dangling));
            Console.WriteLine("  Surviving bookmark names: " + string.Join(", ", names));
            break;
        case "tracked-cell-deletion":
        case "tracked-move-range":
            var revision = name == "tracked-cell-deletion" ? "cellDel" : "moveFromRangeStart";
            observed = analysis.Valid && xml.Descendants(w + revision).Any() &&
                xml.Descendants(w + "t").Any(item => item.Value == "CHANGED");
            Console.WriteLine($"  Revision {revision} retained; value rendered: {observed}");
            break;
        default:
            var crossed = analysis.Valid && analysis.Occurrences.Any(item => item.Expression == "{Name}") &&
                xml.Descendants(w + "t").Any(item => item.Value == "CROSSED");
            Console.WriteLine($"  Unconfirmed contract risk: expression crosses content-control exit: {crossed}");
            break;
    }
    defectObserved |= observed;
    if (name != "expression-crosses-content-control-exit")
        Console.WriteLine("  Confirmed defect signature reproduced: " + observed);
}
Console.WriteLine(defectObserved ? "FAIL: confirmed defects reproduced; Feature 010 remains not accepted."
    : "Recorded signatures absent; this observation does not establish acceptance or a review PASS.");
return defectObserved ? 1 : 0;

string[] Validate(byte[] bytes)
{
    using var stream = new MemoryStream(bytes);
    using var document = WordprocessingDocument.Open(stream, false);
    return new OpenXmlValidator(FileFormatVersions.Office2010).Validate(document)
        .Select(error => error.Description + " " + error.Path?.XPath).ToArray();
}
