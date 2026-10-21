# MonoReport

MonoReport generates DOCX reports, contracts, acts, protocols, invoices, and similar documents from Microsoft Word templates and JSON data. It requires .NET SDK 10. Microsoft Word is useful for authoring and visual verification but is not required by the engine.

1. Feature 012 M2 also supports XLSX templates through Core, CLI and the existing HTTP document endpoints: static fields, single-level vertical runs, nested horizontal blocks, Cartesian tuple data and local scalar A1 formulas with physical PrintArea crop. See [XLSX compatibility and current milestone limits](docs/xlsx-template-compatibility.md). Nested vertical groups, sheet/name/shared-formula binding and Playground support are outside this milestone.

## Projects

1. `MonoReport` is the reusable library.
2. `MonoReport.Cli` is the command-line application and .NET tool.
3. `MonoReport.Api` is the ASP.NET Core HTTP application.
4. `MonoReport.UI` is the React and TypeScript playground served by the API in production.
5. The three test projects cover Core, CLI, HTTP, OpenAPI, and Open XML behavior.

## Build and verify

1. Run the backend checks from the repository root:

   ```bash
   dotnet build MonoReport.slnx -c Release
   dotnet test MonoReport.slnx -c Release
   ```

2. Run the frontend checks from `MonoReport.UI`:

   ```bash
   npm ci
   npm test -- --run
   npm run typecheck
   npm run build
   ```

3. Start the API and UI for local development in separate terminals:

   ```bash
   ASPNETCORE_URLS=http://127.0.0.1:5080 dotnet run --project MonoReport.Api
   cd MonoReport.UI && npm run dev
   ```

## Template syntax

| Purpose | Syntax |
| --- | --- |
| Scalar value | `{Customer.Name}` |
| Current property | `{./Name}` |
| Current array element | `{.}` |
| Parent loop context | `{../Number}` |
| Root context | `{$.Company.Name}` |
| Loop | `{For:Items}` … `{/For}` |
| Condition | `{If:IsActive}` … `{Else}` … `{/If}` |
| Literal DOCX region | `{Raw}` … `{/Raw}` |
| Literal braces | `{{` and `}}` |
| Text format | `{Name upper}`, `{Name lower}`, `{Name caps}`, `{Name firstcap}` |
| Decimal format | `{Amount number:"#,##0.00"}` |
| ISO date format | `{Date date:"dd MMM yyyy"}` |
| Russian RUB words | `{Amount spellout}` |

1. Paths are ordinal and case-sensitive. A segment starts with a Unicode letter or underscore and then contains Unicode letters, digits, or underscores.
2. `For` accepts JSON arrays, including mixed arrays of objects, strings, numbers, Booleans, and null. `{.}` renders the complete current primitive element and is valid only inside a loop.
3. `If` accepts a JSON Boolean. Missing or null structural values select zero iterations or the false branch; incompatible non-null structural values are errors.
4. A scalar expression supports at most one built-in formatter. `number` and `spellout` accept JSON numbers or invariant numeric strings. `date` accepts `YYYY-MM-DD` or an ISO date-time with `Z` or an explicit offset.
5. Explicit null and empty strings render empty. Missing scalar values follow `MergeOptions.MissingValues`.
6. Inserted JSON text and emitted literal braces are never reinterpreted as template syntax.
7. Raw markers occupy dedicated sibling paragraphs or dedicated sibling table rows under the same parent. Content between them is preserved literally without expression parsing, dependency extraction, brace unescaping, or data evaluation. Raw does not nest.

## DOCX compatibility

For the support, preservation, and rejection policy for Microsoft Word features, see [DOCX Template Compatibility](docs/docx-template-compatibility.md). Feature 011 implements this approved contract in the current source and was accepted by the Product Owner on 2026-10-19.

## DOCX structure

1. Scalar expressions work in ordinary paragraphs and tables in the main document, headers, footers, footnotes, and endnotes.
2. Control markers must occupy dedicated sibling paragraphs or dedicated sibling table rows. Paragraph blocks may contain complete tables. Row blocks operate only on sibling rows of one table.
3. Nested blocks and nested tables are supported when every block independently satisfies those boundaries.
4. Repeated and conditional list paragraphs preserve their styles, numbering references, levels, and indentation. Word controls visible continuation; MonoReport does not force a restart per iteration.
5. Horizontal merged cells are preserved. A row operation intersecting a vertical merge is rejected; a complete vertically merged table may be cloned by a paragraph block.
6. Hyperlinks and Word fields are preserved as opaque content. Expressions inside hyperlink text or Word field instructions/results are not processed.
7. Shapes, floating text boxes, SmartArt, charts, and other specialized content are preserved without expression processing.
8. Templates containing unaccepted tracked revisions in supported content are rejected, including cell changes, move/custom-XML revision ranges, deleted payloads, property changes, and Office 2010 conflicts. Accept or reject revisions in Word before generation.
9. Content controls (`w:sdt`) are fully opaque, including inline, block, row, and cell controls. Their XML, properties, binding, and relationships are preserved. Expressions inside are literal; expressions cannot cross entry or exit boundaries.
10. Unaffected comments and bookmarks, valid constructs within or spanning Raw, and complete ranges that enclose a `For` or `If` operation are preserved unchanged. Local comment or bookmark attachments inside `For` or `If`, split ranges, and crossing arrangements are rejected. Complete annotations wholly attached to a removed standalone Raw delimiter are removed safely with unneeded comment definitions.
11. Permission ranges retain local copying support with matching IDs. Drawing IDs remain unique across physical copies and package reservations, including images in repeated Raw content.
12. Explicit internal hyperlinks are preserved only outside `For` and `If` when their targets remain valid and unchanged, including stable Raw targets. Internal links inside structural operations or links to removed, repeated, ambiguous, or invalidated targets are rejected without remapping. External hyperlink relationships and fragments are preserved unchanged.
13. Word fields remain opaque: `REF`, `PAGEREF`, and field-code hyperlinks are not rewritten and their targets are not guaranteed to resolve after structural operations. Unrelated pre-existing unresolved anchors are preserved.

## Core API

1. Register services and inject `MergeEngine` and, when needed, `TemplateAnalyzer`:

   ```csharp
   services.AddMonoReport();

   using var template = File.OpenRead("template.docx");
   using var data = JsonDocument.Parse(json);
   MergeResult result = engine.Merge(template, data, new MergeOptions
   {
       Culture = CultureInfo.GetCultureInfo("en-GB"),
       MissingValues = MissingValueBehavior.Error
   });
   ```

2. `Merge` reads from the template stream's current position and leaves caller-owned inputs open. A successful result contains independent document bytes, warnings and format metadata. `AddMonoReport()` composes both Word and XLSX handlers; existing manually composed Word constructors remain supported. A failed merge throws `TemplateException`, aggregates up to 100 ordered diagnostics, and returns no partial document.
3. `MergeOptions.Culture` controls output casing, decimal formatting, and date formatting. It does not change invariant numeric input parsing, ISO date interpretation, or Russian RUB `spellout` behavior.
4. Custom formatter registration is not available in the unified language. Consumers can transform values before supplying JSON; formatter extensibility may be considered in a future Feature.

## Analyzer

1. `TemplateAnalyzer.Analyze(Stream)` uses the same parser and structural plan as generation without requiring JSON or producing a document.
   Diagnostics retain the offending token or fragment in `field` and the precise part and paragraph/table coordinates in `location`. Their `message` includes bounded nearby template text where available so it can be found in Word; an incomplete fragment is never presented as a complete expression. Analyzer and generation use the same diagnostic messages in Core, API, CLI, and Playground.
2. The result contains `valid`, `errors`, `warnings`, recursive `requirements`, and source-ordered `occurrences`.
3. Each requirement contains only `name`, `type`, and optional nonempty `fields`. Allowed types are `object`, `array`, `value`, and `boolean`. Array fields describe expressions evaluated against elements; they are dependencies, not a homogeneous JSON schema.
4. Each occurrence contains `kind`, the original `expression`, canonical `path`, document `part`, and source `location`. Duplicate occurrences are preserved. XLSX locations identify source worksheet names and A1 coordinates.
5. A loop using `{.}` can produce the requirement `{ "name": "Tags", "type": "array" }` and an occurrence at `$.Tags[]`; no artificial element property is inferred.

## CLI

1. Generate from a JSON file:

   ```bash
   dotnet run --project MonoReport.Cli -- \
     --template Samples/sample.template.docx \
     --json-file Samples/sample.json \
     --output result.docx
   ```

2. Analyze without JSON:

   ```bash
   dotnet run --project MonoReport.Cli -- \
     --analyze --template Samples/sample.template.docx
   ```

3. JSON may instead come from `--json` or `--stdin`. Omit `--output` or use `-` for binary stdout. Diagnostics go to stderr.

## HTTP API and playground

1. `POST /api/documents/generate` accepts multipart fields `template`, `data`, and optional `culture`; success returns raw DOCX or XLSX bytes with format-specific MIME and deterministic filename. Actual package content selects the format, independently of the upload name/MIME. XLSX requires JSON object data.
2. `POST /api/documents/analyze` accepts only multipart file field `template`; success returns the Analyzer contract even when `valid` is false.
3. Swagger UI is available at `/docs` and OpenAPI at `/openapi/v1.json` when documentation is enabled.
4. The browser playground loads the authoritative assets from `Samples`, can analyze a template to propose dependency-based JSON, and does not persist user data.

## Sample and migration

1. `Samples/sample.template.docx` and `Samples/sample.json` are the authoritative template and data pair. `Samples/sample.docx` is their generated output.
2. Feature 010 removed legacy Word `MERGEFIELD`, `TableStart:`, `TableEnd:`, Word switch formatting, and programmatic formatter bindings from MonoReport interpretation. Existing Word fields remain untouched as ordinary opaque document content.
3. Feature 011 changes control closes to `{/For}` and `{/If}`. The former spellings `EndFor` and `EndIf` are ordinary case-sensitive scalar property names, not aliases or migration controls.
4. Convert legacy scalar fields to `{Path}`, collection markers to `{For:Path}` and `{/For}`, conditional closes to `{/If}`, and Word formatting switches to the built-in formatter syntax shown above. Remove comments, bookmarks, and internal links from repeated or conditional bodies, or move complete ranges outside the operation. Use Raw only for literal, structurally bounded Word content.

## Production hosting

1. Publish the combined API and UI:

   ```bash
   dotnet publish MonoReport.Api/MonoReport.Api.csproj -c Release -o artifacts/publish
   ASPNETCORE_URLS=http://0.0.0.0:8080 dotnet artifacts/publish/MonoReport.Api.dll
   ```

2. Or build and run the container:

   ```bash
   docker build -t monoreport:local .
   docker run --rm -p 8080:8080 monoreport:local
   ```
