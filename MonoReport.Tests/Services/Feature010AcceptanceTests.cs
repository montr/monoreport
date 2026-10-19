using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using System.IO.Compression;
using MonoReport.Models;
using MonoReport.Services;

namespace MonoReport.Tests.Services;

public sealed class Feature010AcceptanceTests
{
    [Theory]
    [InlineData("orders-report")]
    [InlineData("project-costs")]
    [InlineData("syntax-showcase")]
    [InlineData("construction-cost-estimate")]
    [InlineData("repeated-image")]
    public void Playground_template_analyzes_and_generates_valid_document(string id)
    {
        using var registry = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "examples.json")));
        var example = registry.RootElement.GetProperty("examples").EnumerateArray()
            .Single(item => item.GetProperty("id").GetString() == id);
        var template = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,
            Path.GetFileName(example.GetProperty("templatePath").GetString()!)));
        var json = File.ReadAllText(Path.Combine(AppContext.BaseDirectory,
            Path.GetFileName(example.GetProperty("dataPath").GetString()!)));
        Docx.AssertValid(template, FileFormatVersions.Office2010);
        using var input = new MemoryStream(template);
        var analysis = Analyzer().Analyze(input);
        Assert.True(analysis.Valid, string.Join(Environment.NewLine,
            analysis.Errors.Select(item => $"{item.Code}: {item.Field} — {item.Location}")));
        var result = Docx.Merge(template, json, new MergeOptions { Culture = CultureInfo.GetCultureInfo("en-GB") },
            FileFormatVersions.Office2010);
        Assert.Empty(result.Warnings);
        var xml = Docx.Read(result.Document);
        using var data = JsonDocument.Parse(json);
        var root = data.RootElement;
        var w = WordXml.WordprocessingNamespace;
        string[] Cells(XElement row) => row.Elements(w + "tc").Select(Docx.Text).ToArray();
        string Money(JsonElement value) => decimal.Parse(value.ToString(), CultureInfo.InvariantCulture)
            .ToString("N2", CultureInfo.GetCultureInfo("en-GB"));
        var text = Docx.Text(xml);
        Assert.DoesNotContain("{For:", text);
        Assert.DoesNotContain("{/For}", text);
        switch (id)
        {
            case "project-costs":
                var table = Assert.Single(xml.Descendants(w + "tbl"));
                Assert.Equal(5, table.Element(w + "tblGrid")!.Elements().Count());
                List<string[]> expected = [["Description", "Unit", "Quantity", "Unit Price, USD", "Total, USD"]];
                decimal total = 0;
                foreach (var project in root.GetProperty("Projects").EnumerateArray())
                {
                    expected.Add([project.GetProperty("Name").GetString()!, Money(project.GetProperty("Total"))]);
                    decimal projectTotal = 0;
                    foreach (var stage in project.GetProperty("Stages").EnumerateArray())
                    {
                        expected.Add(["    " + stage.GetProperty("Name").GetString(), Money(stage.GetProperty("Total"))]);
                        decimal stageTotal = 0;
                        foreach (var item in stage.GetProperty("Items").EnumerateArray())
                        {
                            expected.Add(["        " + item.GetProperty("Name").GetString(), item.GetProperty("Unit").GetString()!,
                                Money(item.GetProperty("Quantity")), Money(item.GetProperty("UnitPrice")), Money(item.GetProperty("Total"))]);
                            Assert.Equal(item.GetProperty("Quantity").GetDecimal() * item.GetProperty("UnitPrice").GetDecimal(), item.GetProperty("Total").GetDecimal());
                            stageTotal += item.GetProperty("Total").GetDecimal();
                        }
                        Assert.Equal(stage.GetProperty("Total").GetDecimal(), stageTotal);
                        projectTotal += stageTotal;
                    }
                    Assert.Equal(project.GetProperty("Total").GetDecimal(), projectTotal);
                    total += projectTotal;
                }
                Assert.Equal(root.GetProperty("GrandTotal").GetDecimal(), total);
                expected.Add(["GRAND TOTAL", Money(root.GetProperty("GrandTotal"))]);
                Assert.Equal(29, expected.Count); // Header + 3 projects + 8 stages + 16 items + total.
                Assert.Equal(expected, table.Elements(w + "tr").Select(Cells).ToList());
                Assert.Equal(12, table.Descendants(w + "gridSpan").Count());
                Assert.All(table.Descendants(w + "gridSpan"), span => Assert.Equal("4", (string?)span.Attribute(w + "val")));
                break;
            case "orders-report":
                var orders = root.GetProperty("Orders").EnumerateArray().ToArray();
                var tables = xml.Descendants(w + "tbl").ToArray();
                Assert.Equal(20, tables.Length);
                Assert.Equal(orders.Length, root.GetProperty("OrderCount").GetInt32());
                var subtotals = xml.Descendants(w + "p").Select(Docx.Text).Where(value => value.StartsWith("Order subtotal:")).ToArray();
                Assert.Equal(orders.Select(order => $"Order subtotal: {Money(order.GetProperty("Total"))} RUB"), subtotals);
                for (var index = 0; index < orders.Length; index++)
                {
                    var order = orders[index];
                    var items = order.GetProperty("Items").EnumerateArray().ToArray();
                    Assert.Equal(order.GetProperty("LineCount").GetInt32(), items.Length);
                    Assert.Contains("Order " + order.GetProperty("Number").GetString(), text);
                    Assert.Equal(items.Select(item => new[] { item.GetProperty("LineNumber").ToString(), item.GetProperty("Sku").GetString()!,
                        item.GetProperty("Name").GetString()!, item.GetProperty("Quantity").ToString(), item.GetProperty("Unit").GetString()!,
                        Money(item.GetProperty("UnitPrice")), Money(item.GetProperty("LineTotal")) }), tables[index].Elements(w + "tr").Skip(1).Select(Cells));
                    Assert.Equal(decimal.Parse(order.GetProperty("Total").GetString()!, CultureInfo.InvariantCulture),
                        items.Sum(item => decimal.Parse(item.GetProperty("LineTotal").GetString()!, CultureInfo.InvariantCulture)));
                }
                Assert.Contains("REPORT TOTAL: 60,206.90 RUB", text);
                Assert.Equal(decimal.Parse(root.GetProperty("GrandTotal").GetString()!, CultureInfo.InvariantCulture),
                    orders.Sum(order => decimal.Parse(order.GetProperty("Total").GetString()!, CultureInfo.InvariantCulture)));
                break;
            case "syntax-showcase":
                var paragraphs = xml.Descendants(w + "p").Select(Docx.Text).ToArray();
                Assert.DoesNotContain(analysis.Requirements, requirement => requirement.Name == "RawOnly");
                Assert.Equal(new[] { "• PRIORITY", "• EXPORT", "• REVIEW", "• READY" }, paragraphs.Where(value => value.StartsWith("• ")));
                Assert.Contains("APPROVED — the condition is true.", paragraphs);
                Assert.Contains("This record is active.", paragraphs);
                Assert.DoesNotContain("PENDING — the condition is false.", paragraphs);
                Assert.DoesNotContain("This record is archived.", paragraphs);
                var showcaseTables = xml.Descendants(w + "tbl").ToArray();
                Assert.Equal(new[] { "Northwind Demo Company", "HELLO WORLD FROM MONOREPORT", "hello world from monoreport",
                    "HELLO WORLD From MonoReport", "HELLO wORLD from monoReport", "12,345.68", "-42.51", "17 Oct 2026",
                    "21 Oct 2026 14:30 +03:00", "Одна тысяча двести пятьдесят рублей 75 копеек" },
                    showcaseTables[0].Elements(w + "tr").Skip(1).Select(row => Cells(row)[2]));
                Assert.Equal(root.GetProperty("Products").EnumerateArray().Select(product => new[] {
                    product.GetProperty("Sku").GetString()!, product.GetProperty("Name").GetString()!,
                    product.GetProperty("Quantity").ToString(), Money(product.GetProperty("Price")) }),
                    showcaseTables[1].Elements(w + "tr").Skip(1).Select(Cells));
                foreach (var department in root.GetProperty("Departments").EnumerateArray())
                    foreach (var member in department.GetProperty("Members").EnumerateArray())
                        Assert.Contains($"{member.GetString()} — {department.GetProperty("Name").GetString()}", paragraphs);
                foreach (var employee in root.GetProperty("Employees").EnumerateArray())
                {
                    Assert.Contains($"{employee.GetProperty("Name").GetString()} — {employee.GetProperty("Role").GetString()}", paragraphs);
                    Assert.Contains($"{employee.GetProperty("Name").GetString()} works at Northwind Demo Company.", paragraphs);
                }
                Assert.Contains("Literal braces: {Company.Name}", paragraphs);
                Assert.Contains("Braces around a value: {Northwind Demo Company}", paragraphs);
                Assert.Contains("Literal template text: {RawOnly.Value} {broken {{ {/If}", paragraphs);
                Assert.Equal(3, xml.Descendants(w + "pStyle").Count(style => (string?)style.Attribute(w + "val") == "ListNumber"));
                Assert.Equal(5, xml.Descendants(w + "pStyle").Count(style => (string?)style.Attribute(w + "val") == "ListBullet"));
                break;
            case "construction-cost-estimate":
                var sampleTable = Assert.Single(xml.Descendants(w + "tbl"));
                Assert.Equal(4, sampleTable.Elements(w + "tr").Count());
                Assert.Equal(root.GetProperty("Materials").EnumerateArray().Select(material => material.GetProperty("Reference").GetString()),
                    sampleTable.Elements(w + "tr").Skip(1).Select(row => Cells(row)[0]));
                Assert.Contains("Total estimated cost: " + Money(root.GetProperty("Totals").GetProperty("Amount")), text);
                Assert.True(XNode.DeepEquals(Docx.Read(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "sample.docx"))), xml));
                break;
            case "repeated-image":
                Assert.Equal(3, xml.Descendants(w + "drawing").Count());
                var imageIds = xml.Descendants(XName.Get("docPr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"))
                    .Select(item => (string?)item.Attribute("id")).ToArray();
                Assert.Equal(3, imageIds.Distinct().Count());
                Assert.Equal(root.GetProperty("Items").EnumerateArray().Select(item => $"Item: {item.GetProperty("Name").GetString()} — {item.GetProperty("Code").GetString()}"),
                    xml.Descendants(w + "p").Select(Docx.Text).Where(value => value.StartsWith("Item: ")));
                Assert.Equal(root.GetProperty("Items").EnumerateArray().Select(item => "Caption: " + item.GetProperty("Caption").GetString()),
                    xml.Descendants(w + "p").Select(Docx.Text).Where(value => value.StartsWith("Caption: ")));
                using (var originalPackage = new ZipArchive(new MemoryStream(template)))
                using (var generatedPackage = new ZipArchive(new MemoryStream(result.Document)))
                    foreach (var entry in originalPackage.Entries.Where(item => item.FullName.StartsWith("word/media/") || item.FullName == "word/_rels/document.xml.rels"))
                    {
                        using var before = entry.Open();
                        using var after = generatedPackage.GetEntry(entry.FullName)!.Open();
                        using var beforeBytes = new MemoryStream();
                        using var afterBytes = new MemoryStream();
                        before.CopyTo(beforeBytes); after.CopyTo(afterBytes);
                        Assert.Equal(beforeBytes.ToArray(), afterBytes.ToArray());
                    }
                break;
        }
    }

    [Theory]
    [InlineData("proofErr")]
    [InlineData("lastRenderedPageBreak")]
    [InlineData("bookmarkStart")]
    [InlineData("permStart")]
    public void Invisible_Word_metadata_does_not_split_scalar_or_control_expressions(string kind)
    {
        var w = WordXml.WordprocessingNamespace;
        XElement[] Metadata(string value)
        {
            var start = kind == "proofErr"
                ? new XElement(w + kind, new XAttribute(w + "type", "spellStart"))
                : kind == "lastRenderedPageBreak" ? new XElement(w + kind)
                : new XElement(w + kind, new XAttribute(w + "id", 1),
                    new XAttribute(w + (kind == "bookmarkStart" ? "name" : "ed"), value));
            var end = kind == "proofErr"
                ? new XElement(w + kind, new XAttribute(w + "type", "spellEnd"))
                : kind == "lastRenderedPageBreak" ? new XElement(w + kind)
                : new XElement(w + (kind == "bookmarkStart" ? "bookmarkEnd" : "permEnd"), new XAttribute(w + "id", 1));
            return [start, end];
        }
        XElement Split(string left, string middle, string right)
        {
            var metadata = Metadata("target");
            return kind == "lastRenderedPageBreak"
                ? Docx.P(new XElement(w + "r", new XElement(w + "t", left), metadata[0],
                    new XElement(w + "t", middle), metadata[1], new XElement(w + "t", right)))
                : Docx.P(Docx.Run(left), metadata[0], Docx.Run(middle), metadata[1], Docx.Run(right));
        }
        // Use separate fixtures so bookmark/permission IDs stay unique in each source package.
        var scalar = Docx.Create(Split("Cost: {", "Amount", " number:\"N2\"}"));
        Docx.AssertValid(scalar);
        using var input = new MemoryStream(scalar);
        Assert.True(Analyzer().Analyze(input).Valid);
        var generated = Docx.Merge(scalar, """{"Amount":1234.5}""",
            new MergeOptions { Culture = CultureInfo.GetCultureInfo("en-GB") }).Document;
        Assert.Equal("Cost: 1,234.50", Docx.Text(Docx.Read(generated)));
        Assert.Equal(2, Docx.Read(generated).Descendants().Count(item =>
            item.Name == w + kind || item.Name == w + "bookmarkEnd" || item.Name == w + "permEnd"));

        var loop = Docx.Create(Docx.P(Docx.Run("{For:Tags}")), Docx.P(Docx.Run("• {. upper}")),
            Split("{/", "For", "}"));
        Docx.AssertValid(loop);
        if (kind == "bookmarkStart")
        {
            Feature010IdentityTests.Reject(loop, """{"Tags":["one","two"]}""", "template.invalid_structure");
            return;
        }
        using var loopInput = new MemoryStream(loop);
        Assert.True(Analyzer().Analyze(loopInput).Valid);
        Assert.Equal("• ONE• TWO", Docx.Text(Docx.Read(Docx.Merge(loop, """{"Tags":["one","two"]}""").Document)));
    }

    [Fact]
    public void Three_nested_row_loops_keep_five_column_merged_markers_and_empty_inner_collections()
    {
        var w = WordXml.WordprocessingNamespace;
        XElement MergedRow(params object[] contents) => new(w + "tr", new XElement(w + "tc",
            new XElement(w + "tcPr", new XElement(w + "gridSpan", new XAttribute(w + "val", 5))), Docx.P(contents)));
        var template = Docx.Create(new XElement(w + "tbl", new XElement(w + "tblPr"),
            new XElement(w + "tblGrid", Enumerable.Range(0, 5).Select(_ => new XElement(w + "gridCol", new XAttribute(w + "w", 1000)))),
            MergedRow(Docx.Run("{For:Projects}")), MergedRow(Docx.Run("Project {Name}")),
            MergedRow(Docx.Run("{For:Stages}")), MergedRow(Docx.Run("Stage {Name}")),
            MergedRow(Docx.Run("{For:Items}")),
            new XElement(w + "tr", new[] { "{Name}", "{../../Name}", "{../Name}", "{$.Currency}", "{Total number:\"N2\"}" }
                .Select(text => new XElement(w + "tc", Docx.P(Docx.Run(text))))),
            MergedRow(Docx.Run("{/"), new XElement(w + "proofErr", new XAttribute(w + "type", "spellStart")),
                Docx.Run("For"), new XElement(w + "proofErr", new XAttribute(w + "type", "spellEnd")), Docx.Run("}")),
            MergedRow(Docx.Run("{/For}")), MergedRow(Docx.Run("{/For}"))));
        Docx.AssertValid(template);
        using var input = new MemoryStream(template);
        var analysis = Analyzer().Analyze(input);
        Assert.True(analysis.Valid, string.Join("; ", analysis.Errors));
        var result = Docx.Merge(template, """
            {"Currency":"USD","Projects":[{"Name":"P1","Stages":[{"Name":"S1","Items":[{"Name":"A","Total":2},{"Name":"B","Total":3}]},{"Name":"Empty","Items":[]}]},{"Name":"P2","Stages":[{"Name":"S2","Items":[{"Name":"C","Total":4}]}]}]}
            """, new MergeOptions { Culture = CultureInfo.GetCultureInfo("en-GB") });
        var table = Assert.Single(Docx.Read(result.Document).Descendants(w + "tbl"));
        Assert.Equal(new[] { "Project P1", "Stage S1", "AP1S1USD2.00", "BP1S1USD3.00", "Stage Empty", "Project P2", "Stage S2", "CP2S2USD4.00" },
            table.Elements(w + "tr").Select(Docx.Text));
        Assert.Equal(5, table.Descendants(w + "gridSpan").Count());
    }

    [Theory]
    [InlineData("{Broken", "{Broken")]
    [InlineData("Broken}", "}")]
    [InlineData("{For:Tags} extra", "{For:Tags}")]
    public void Diagnostics_include_actual_token_and_nearby_text_consistently(string text, string token)
    {
        var template = Docx.Create(Docx.P(Docx.Run("Invoice reference: " + text)));
        using var input = new MemoryStream(template);
        var issue = Assert.Single(Analyzer().Analyze(input).Errors);
        Assert.Equal(token, issue.Field);
        Assert.Contains("Invoice reference:", issue.Message);
        Assert.Contains(token, issue.Message);
        Assert.Contains("/word/document.xml, paragraph 1", issue.Location);
        using var generation = new MemoryStream(template);
        using var data = JsonDocument.Parse("{}");
        var error = Assert.Throws<TemplateException>(() => Docx.CreateEngine().Merge(generation, data));
        var diagnostic = Assert.Single(error.Errors);
        Assert.Equal(issue.Message, diagnostic.Message);
        Assert.Equal(issue.Field, diagnostic.Field);
        Assert.Equal(issue.Location, diagnostic.Location);
    }

    private static TemplateAnalyzer Analyzer() => new(new DocxStoryReader(), new DocxTemplatePlanner(new TemplateLanguageParser()));

    [Fact]
    public void Diagnostic_context_finds_a_late_fragment_without_dumping_the_entire_paragraph()
    {
        var text = new string('x', 5000) + " Invoice reference: {Missing" + new string('z', 5000);
        using var input = new MemoryStream(Docx.Create(Docx.P(Docx.Run(text))));
        var issue = Assert.Single(Analyzer().Analyze(input).Errors);
        Assert.StartsWith("{Missing", issue.Field);
        Assert.Contains("Invoice reference:", issue.Message);
        Assert.True(issue.Message.Length < 450);
        Assert.DoesNotContain("}", issue.Field);
    }
}
