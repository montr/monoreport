using System.Collections.Immutable;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MonoReport.Models;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxPackageReader
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal XlsxSourceInventory Read(SpreadsheetDocument document, XlsxOperationContext context)
    {
        var workbook = document.WorkbookPart ?? throw Invalid("The XLSX has no workbook part.");
        var xml = Load(workbook, context);
        if (xml.Root?.Name != S + "workbook") throw Invalid("The XLSX workbook XML is invalid.");
        var disposable = DisposableWorkbookNodes(xml.Root);
        var effectiveNodes = xml.Root.Elements().Where((_, index) => !disposable.Contains(index)).ToArray();
        var related = ReadRelatedParts(workbook);
        var names = xml.Root.Element(S + "definedNames")?.Elements().Select(name => new XElement(name)).ToImmutableArray() ?? [];
        var shared = workbook.SharedStringTablePart is { } strings ? ReadSharedStrings(strings, context) : [];
        var decodedShared = new Dictionary<int, XlsxRichString>();
        var result = ImmutableArray.CreateBuilder<XlsxSheetSource>();
        var sheetIds = new HashSet<uint>();
        var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sheetParts = new HashSet<Uri>();
        var relationships = related.ToDictionary(pair => pair.Id, pair => pair.Part, StringComparer.Ordinal);
        var sheets = xml.Root.Element(S + "sheets")?.Elements(S + "sheet").ToArray() ?? [];
        if (sheets.Length == 0) throw Invalid("The workbook must contain a worksheet.");
        for (var index = 0; index < sheets.Length && !context.LimitReached; index++)
        {
            var sheet = sheets[index];
            var name = (string?)sheet.Attribute("name") ?? string.Empty;
            uint.TryParse((string?)sheet.Attribute("sheetId"), NumberStyles.None, CultureInfo.InvariantCulture, out var id);
            var relationship = (string?)sheet.Attribute(XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships") + "id");
            if (id == 0 || !sheetIds.Add(id) || name.Length == 0 || !sheetNames.Add(name) || relationship is null ||
                !relationships.TryGetValue(relationship, out var target) || target is not WorksheetPart part || !sheetParts.Add(part.Uri))
                throw Invalid("Use uniquely named ordinary worksheets with valid IDs and distinct worksheet relationships.", workbook.Uri.ToString());
            var location = part.Uri + ", sheet '" + name + "'";
            try
            {
                var (area, nameIndex) = PrintArea(names, index, name, location);
                var sheetXml = Load(part, context);
                if (sheetXml.Root?.Name != S + "worksheet" || sheetXml.Root.Elements(S + "sheetData").Count() != 1)
                    throw Invalid("The worksheet must contain exactly one sheetData element.", location);
                var cells = ImmutableArray.CreateBuilder<XlsxCellSource>();
                var addresses = new HashSet<XlsxAddress>();
                var rows = new HashSet<int>();
                if (sheetXml.Root.Element(S + "sheetData")!.Elements().Any(element => element.Name != S + "row"))
                    throw Invalid("Remove unsupported sheetData child structures.", location);
                var previousRow = 0;
                foreach (var row in sheetXml.Root.Element(S + "sheetData")?.Elements(S + "row") ?? [])
                {
                    if (!int.TryParse((string?)row.Attribute("r"), NumberStyles.None, CultureInfo.InvariantCulture, out var rowIndex) ||
                        rowIndex is < 1 or > 1048576 || !rows.Add(rowIndex) || rowIndex <= previousRow)
                        throw Invalid("Use explicit unique row coordinates in increasing order.", location);
                    previousRow = rowIndex;
                    var previousColumn = 0;
                    foreach (var cell in row.Elements(S + "c"))
                    {
                        XlsxAddress address;
                        try { address = Address((string?)cell.Attribute("r")); }
                        catch (TemplateException) { throw Invalid("Use finite explicit cell coordinates within Excel dimensions.", location); }
                        if (address.Row != rowIndex || !addresses.Add(address) || address.Column <= previousColumn)
                            throw Invalid("Use unique increasing cell coordinates matching their row.", location);
                        previousColumn = address.Column;
                        cells.Add(ReadCell(id, address, cell, shared, decodedShared, location));
                    }
                }
                context.SourceCells = checked(context.SourceCells + cells.Count);
                context.SourceFormulas = checked(context.SourceFormulas + cells.Count(cell => cell.HasFormula));
                var features = sheetXml.Root.Elements().Where(element => element.Name != S + "sheetData")
                    .Select(element => element.Name.LocalName).Concat(ReadRelatedParts(part).Select(pair => pair.Part.RelationshipType))
                    .Concat(part.ExternalRelationships.Select(item => item.RelationshipType)).ToImmutableArray();
                result.Add(new(id, name, part.Uri.ToString(), sheetXml, area, area ?? new(new(1, 1), new(1048576, 16384)),
                    cells.OrderBy(cell => cell.Id.Address.Row).ThenBy(cell => cell.Id.Address.Column).ToImmutableArray(), features, nameIndex));
            }
            catch (TemplateException error) { context.Report(error); }
        }
        context.ThrowIfErrors();
        if (names.Select((name, index) => (name, index)).Any(item => (string?)item.name.Attribute("name") == "_xlnm.Print_Area" &&
            !result.Any(sheet => sheet.PrintAreaNameIndex == item.index)))
            throw AreaError(workbook.Uri.ToString());
        var date1904 = (string?)xml.Root.Element(S + "workbookPr")?.Attribute("date1904") switch
        {
            null or "false" or "0" => false,
            "true" or "1" => true,
            _ => throw Invalid("Repair the workbook date1904 Boolean setting.", workbook.Uri.ToString())
        };
        return new(result.ToImmutable(), date1904,
            names, effectiveNodes.Concat(effectiveNodes.SelectMany(element => element.Descendants())
                    .Where(element => element.Name.Namespace != S || element.Name == S + "extLst"))
                .Select(element => (element.Name.ToString(), DescribeWorkbookFeature(element)))
                .Concat(related.Select(pair => pair.Part.RelationshipType).Concat(workbook.ExternalRelationships.Select(item => item.RelationshipType))
                    .Select(type => (type, "relationship type " + WorkbookIdentifier(type)))).ToImmutableArray(),
            workbook.WorkbookStylesPart is { } styles ? Load(styles, context) : null)
        { DisposableWorkbookNodes = disposable, WorkbookPart = workbook.Uri.ToString() };
    }

    private XElement[] ReadSharedStrings(SharedStringTablePart part, XlsxOperationContext context)
    {
        var root = Load(part, context).Root;
        if (root?.Name != S + "sst" || !OnlyAttributes(root, "count", "uniqueCount") ||
            !OnlyElementContent(root) || root.Elements().Any(element => element.Name != S + "si"))
            throw Invalid("Remove unsupported shared-string table structures or extension metadata before conversion to inline strings.", part.Uri.ToString());
        if (ReadRelatedParts(part).Count != 0 || part.ExternalRelationships.Any() ||
            part.HyperlinkRelationships.Any() || part.DataPartReferenceRelationships.Any())
            throw Invalid("Remove unsupported shared-string table relationships before conversion to inline strings.", part.Uri.ToString());
        return root.Elements(S + "si").ToArray();
    }

    private static XlsxCellSource ReadCell(uint id, XlsxAddress address, XElement cell, XElement[] shared, Dictionary<int, XlsxRichString> decodedShared, string location)
    {
        var formula = cell.Element(S + "f") is not null;
        var type = (string?)cell.Attribute("t");
        string? text = null;
        XElement? sharedString = null;
        var rich = false;
        XlsxRichString? richString = null;
        if (!formula && type == "s")
        {
            if (!int.TryParse(cell.Element(S + "v")?.Value, out var index) || index < 0 || index >= shared.Length)
                throw Invalid("The cell has an invalid shared-string index.", location + ", " + address);
            if (!decodedShared.TryGetValue(index, out richString))
                decodedShared.Add(index, richString = ReadRichString(shared[index], location + ", " + address));
            text = richString.Text;
            sharedString = new XElement(shared[index]);
            rich = shared[index].Elements(S + "r").Any() || shared[index].Elements(S + "rPh").Any();
        }
        else if (!formula && type == "inlineStr")
        {
            var inline = cell.Element(S + "is");
            richString = inline is null ? null : ReadRichString(inline, location + ", " + address);
            text = richString?.Text ?? string.Empty;
            rich = inline?.Elements().Any(item => item.Name == S + "r" || item.Name == S + "rPh") == true;
        }
        else if (!formula && type == "str") text = DecodeString(cell.Element(S + "v")?.Value ?? string.Empty);
        return new(new(id, address), new XElement(cell), text, type, formula, rich, sharedString) { RichString = rich ? richString : null };
    }

    private static XlsxRichString ReadRichString(XElement container, string location)
    {
        if (container.Elements().Any(element => element.Name != S + "t" && element.Name != S + "r" &&
                element.Name != S + "rPh" && element.Name != S + "phoneticPr") ||
            container.Descendants().Any(element => element.Name.Namespace != S || element.Name == S + "extLst") ||
            container.Elements(S + "r").Any(run => run.Elements().Any(element => element.Name != S + "t" && element.Name != S + "rPr")))
            throw Invalid("Remove unsupported rich-string structures or extension metadata.", location);
        var spans = ImmutableArray.CreateBuilder<XlsxTextSpan>();
        var length = 0;
        foreach (var item in container.Elements())
        {
            var properties = item.Name == S + "r" ? item.Element(S + "rPr") : null;
            foreach (var text in item.Name == S + "t" ? new[] { item } : item.Name == S + "r" ? item.Elements(S + "t") : [])
            {
                var decoded = DecodeString(text.Value);
                spans.Add(new(length, decoded, properties is null ? null : new XElement(properties)));
                length = checked(length + decoded.Length);
            }
        }
        var phonetics = container.Elements(S + "rPh").Select(node => new XElement(node)).ToImmutableArray();
        foreach (var hint in phonetics)
            if (!int.TryParse((string?)hint.Attribute("sb"), NumberStyles.None, CultureInfo.InvariantCulture, out var start) ||
                !int.TryParse((string?)hint.Attribute("eb"), NumberStyles.None, CultureInfo.InvariantCulture, out var end) ||
                start < 0 || end <= start || end > length)
                throw Invalid("Repair phonetic annotation positions within the decoded cell text.", location);
        return new(string.Concat(spans.Select(span => span.Text)), spans.ToImmutable(), phonetics, container.Element(S + "phoneticPr") is { } phoneticProperties ? new XElement(phoneticProperties) : null);
    }

    // Each t element is one ST_Xstring. Decode once before scalar parsing; preserved XML stays untouched.
    private static string DecodeString(string value) => StringEscape().Replace(value,
        match => ((char)int.Parse(match.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString());

    [GeneratedRegex(@"_x([0-9A-Fa-f]{4})_", RegexOptions.CultureInvariant)]
    private static partial Regex StringEscape();

    private static IReadOnlyList<(string Id, OpenXmlPart Part)> ReadRelatedParts(OpenXmlPartContainer container)
    {
        try { return container.Parts.Select(pair => (pair.RelationshipId, pair.OpenXmlPart)).ToArray(); }
        catch (InvalidOperationException)
        {
            // The SDK reports a missing relationship target as InvalidOperationException while loading its graph.
            throw new TemplateException("The XLSX relationship graph is inconsistent.", code: DiagnosticCodes.XlsxInvalidPackage);
        }
    }

    private static (XlsxRectangle? Area, int NameIndex) PrintArea(ImmutableArray<XElement> names, int index, string sheet, string location)
    {
        var candidates = names.Where(name => (string?)name.Attribute("name") == "_xlnm.Print_Area" &&
            ((string?)name.Attribute("localSheetId") == index.ToString(CultureInfo.InvariantCulture) ||
             name.Attribute("localSheetId") is null && AreaRegex().Match(name.Value) is { Success: true } match && SheetName(match) == sheet)).ToArray();
        if (candidates.Length == 0) return (null, -1);
        if (candidates.Length != 1) throw AreaError(location);
        var parsed = AreaRegex().Match(candidates[0].Value);
        if (!parsed.Success || SheetName(parsed) != sheet) throw AreaError(location);
        try
        {
            var first = Address(parsed.Groups["first"].Value);
            var last = parsed.Groups["last"].Success ? Address(parsed.Groups["last"].Value) : first;
            if (last.Row < first.Row || last.Column < first.Column) throw AreaError(location);
            return (new(first, last), names.IndexOf(candidates[0]));
        }
        catch (TemplateException) { throw AreaError(location); }
    }

    private static string SheetName(Match match) => match.Groups["quoted"].Success
        ? match.Groups["quoted"].Value.Replace("''", "'", StringComparison.Ordinal) : match.Groups["plain"].Value;

    private static XlsxAddress Address(string? text)
    {
        var match = AddressRegex().Match(text ?? string.Empty);
        if (!match.Success) throw Invalid("Use finite A1 cell coordinates within Excel dimensions.");
        var column = 0;
        foreach (var letter in match.Groups["column"].Value) column = checked(column * 26 + letter - 'A' + 1);
        if (!int.TryParse(match.Groups["row"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var row) ||
            row is < 1 or > 1048576 || column is < 1 or > 16384) throw Invalid("Cell coordinates exceed Excel dimensions.");
        return new(row, column);
    }

    private static XDocument Load(OpenXmlPart part, XlsxOperationContext context)
    {
        try
        {
            using var stream = new CountedReadStream(part.GetStream(FileMode.Open, FileAccess.Read), context);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException)
        {
            throw new StructuralDiagnostic(DiagnosticCodes.TemplateInvalidXml, "Invalid XLSX part XML.", Location: part.Uri.ToString(),
                Explanation: "Repair or re-save the indicated workbook part in Excel.").ToException();
        }
    }

    private sealed class CountedReadStream(Stream inner, XlsxOperationContext context) : Stream
    {
        public override bool CanRead => inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = inner.Read(buffer, offset, count);
            context.PartBytesRead = checked(context.PartBytesRead + read);
            return read;
        }
        public override int Read(Span<byte> buffer)
        {
            var read = inner.Read(buffer);
            context.PartBytesRead = checked(context.PartBytesRead + read);
            return read;
        }
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) inner.Dispose(); base.Dispose(disposing); }
    }

    private static TemplateException Invalid(string explanation, string? location = null) =>
        new StructuralDiagnostic(DiagnosticCodes.XlsxUnsupportedFeature, explanation, Location: location, Explanation: explanation).ToException();
    private static TemplateException AreaError(string location) =>
        new("Set one finite rectangular PrintArea in Excel for this worksheet.", location: location, code: DiagnosticCodes.XlsxInvalidPrintArea);

    [GeneratedRegex(@"^(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[^'!,:\[\]]+))!(?<first>\$?[A-Z]{1,3}\$?[1-9][0-9]*)(?::(?<last>\$?[A-Z]{1,3}\$?[1-9][0-9]*))?$", RegexOptions.CultureInvariant)]
    private static partial Regex AreaRegex();
    [GeneratedRegex(@"^\$?(?<column>[A-Z]{1,3})\$?(?<row>[1-9][0-9]*)$", RegexOptions.CultureInvariant)]
    private static partial Regex AddressRegex();
}
