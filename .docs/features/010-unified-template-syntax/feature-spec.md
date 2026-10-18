# Feature Spec — Unified Template Syntax and Conditional Blocks

**Status:** Accepted — explicit Product Owner acceptance on 2026-10-18; see [review section 21](review.md#21-product-owner-acceptance-on-2026-10-18)

**Scope:** DOCX implementation; syntax and data semantics designed for future XLSX reuse

**Risk:** High-risk breaking Feature affecting public contracts and DOCX integrity

**Breaking change:** Replaces Word MERGEFIELD templates and their analyzer contract with MonoReport `{...}` expressions.

## 1. Goal

1. Replace the Word-specific MERGEFIELD model with one concise, format-neutral language for scalar fields, nested collection loops, parent/root access, and Boolean conditional blocks.
2. Simplify the DOCX engine around one shared parser and semantic model used by analysis and rendering.
3. Produce a normal DOCX with processed expressions removed. Generated documents are not reusable templates.
4. Keep DOCX manipulation isolated from the shared language model so future formats can reuse the language without implementing XLSX in this Feature.
5. The Product Owner will separately tag and release the final MERGEFIELD version. No Git or publishing operation is authorized by this Feature.
6. MonoReport is a document-generation engine for reports, contracts, acts, protocols, invoices, and similar documents. It does not aim to interpret every Microsoft Word feature.

## 2. Language grammar

1. A scalar field uses `{Customer.Name}` and resolves a dotted path against the current data context.
2. A loop uses `{For:Items}` followed by its body and `{EndFor}`. It renders once per JSON array element and pushes that entire element as the current context. Array elements may be objects, strings, numbers, Booleans, or null, and one array may mix these kinds. A string value is not a collection and is never iterated by character.
3. A condition uses `{If:IsActive}` followed by its true branch and `{EndIf}`. An optional `{Else}` selects an alternative branch. `If` does not push a data context.
4. `For` and `If` may be nested in any combination when their DOCX regions satisfy the structural rules in this specification. `Else` belongs to the nearest open `If`, is legal only once in that `If`, and is not legal in `For`.
5. `{.}` is a scalar expression for the entire current array element. It is valid only inside `For`; outside a loop it produces an Analyzer syntax/context error.
6. A bare property path and an explicit `./` property path resolve from the current context. `{./Name}` means the `Name` property of the current object and must remain lexically and semantically distinct from `{.}`.
7. Each `../` moves to one enclosing loop context and may be repeated. `$.` resolves from the root JSON value.
8. Property paths use dots as separators. Every segment starts with a Unicode letter or underscore and continues with Unicode letters, digits, or underscores. Matching is ordinal and case-sensitive.
9. Quoted property segments, bracket notation, arbitrary punctuation in property names, JSONPath evaluation, wildcard evaluation, calculations, functions, and general expressions are not supported.
10. Canonical dependency paths start with `$`, use dots for properties, and append `[]` to an array element scope. For example, a `Name` field inside nested `Orders` and `Items` loops has canonical path `$.Orders[].Items[].Name`; `{.}` directly inside `{For:Tags}` has canonical path `$.Tags[]`.
11. Canonical paths identify requirements and diagnostics; they are not an alternative input grammar or general-purpose query language.

## 3. Scalar formatting

1. A formatter follows the scalar expression after whitespace. Supported forms include `{Name upper}`, `{Amount number:"0.00"}`, `{Date date:"dd.MM.yyyy"}`, `{. upper}`, and `{. number:"0.00"}`.
2. Feature 010 supports exactly one formatter per scalar expression: `upper`, `lower`, `caps`, `firstcap`, `spellout`, `number:"pattern"`, or `date:"pattern"`.
3. Formatter chaining, pipelines, pipes, function syntax, arbitrary registered formatter names, and formatters on structural expressions are invalid. Chaining is explicitly deferred; consumers may perform complex transformations before supplying JSON.
4. `upper` and `lower` convert the ordinary textual representation of a string, number, or Boolean using `MergeOptions.Culture.TextInfo`. The ordinary textual representation of JSON numbers is their JSON text and of Booleans is lowercase `true` or `false`.
5. `caps` uppercases every Unicode letter at a word boundary using `MergeOptions.Culture.TextInfo`; it does not lowercase or otherwise change the remaining characters. `firstcap` uppercases the first Unicode letter using the same culture and leaves all preceding and following characters unchanged. Both accept the ordinary textual representation described in item 4.
6. `number:"pattern"` accepts a JSON number or an invariant-culture numeric string and converts it with decimal arithmetic. Numeric strings are parsed with standard .NET decimal parsing using `NumberStyles.Float` and `CultureInfo.InvariantCulture`; invariant decimal notation, exponent notation, and surrounding whitespace are accepted within the decimal range. Thousands or group separators, culture-specific decimal separators, non-finite values, invalid strings, and overflow are generation errors. The resulting decimal is formatted with the supplied .NET decimal format pattern and `MergeOptions.Culture`.
7. `date:"pattern"` accepts only these ISO 8601 strings: calendar dates in `YYYY-MM-DD` form, and date-times in `YYYY-MM-DDTHH:mm:ss` form with an optional fraction of one through seven digits followed by `Z` or an explicit `+HH:mm` or `-HH:mm` offset. A date-only value is formatted as midnight with offset zero. A date-time without an explicit offset, a space in place of `T`, an ambiguous culture-specific date, or another unsupported ISO 8601 representation is rejected rather than interpreted using the machine's local time zone.
8. Date parsing is invariant and deterministic. Preserve a date-time's supplied offset when formatting; treat `Z` as offset zero. `MergeOptions.Culture` controls only the formatted output produced by the supplied .NET date/time format pattern.
9. `spellout` accepts a JSON number or invariant-culture numeric string representing Russian ruble and kopeck amounts. It applies exactly the numeric-string parsing rule in item 6, including exponent and surrounding-whitespace support and rejection of group separators, culture-specific decimal separators, non-finite values, invalid strings, and overflow. It rounds to two fractional digits away from zero, spells the ruble amount with the existing Russian grammatical forms, renders kopecks as two numeric digits with the appropriate Russian form, supports negative values, and requires an absolute rounded ruble amount below 10^18. It is culture-independent and supports RUB only; no currency argument or additional currency is supported.
10. Explicit null and the empty string render empty without attempting formatter conversion. Missing values continue to follow `MergeOptions.MissingValues`.
11. Objects and arrays remain invalid scalar values for every formatter. Invalid numeric or date input, incompatible primitive types, overflow, and `spellout` range failures produce normal generation diagnostics under the existing error-aggregation and no-partial-document behavior. In particular, a Boolean passed to `spellout` must not cause an uncaught exception.
12. Analyzer errors cover unknown formatters, missing or invalid formatter arguments, multiple formatters, formatters on structural expressions, and statically invalid .NET number or date format patterns. These checks occur regardless of whether runtime data is null or empty. Data-dependent conversion and range failures are generation errors.
13. Reuse the existing casing and Russian RUB algorithms unchanged where they implement these semantics. Adapt numeric and date input parsing, value-kind checks, and exception handling to enforce the approved domains and diagnostics without retaining Word field-code switches.
14. Remove `MergeOptions.FieldFormatters`, `FormatterBinding`, custom formatter registration, arbitrary formatter option dictionaries, and related compatibility behavior. Do not introduce or design a replacement extension API in Feature 010, and do not preserve the old registry solely for compatibility.
15. The intentional capability loss includes consumer-defined transformations, per-merge formatter selection without editing a template, raw-JSON custom formatter access, arbitrary formatter options, and consumer-controlled formatter failures. Custom formatter extensibility remains a possible future Feature rather than a permanent prohibition.

## 4. Braces and tokenization

1. Every unescaped `{` begins a template expression and every unescaped `}` closes one. Malformed or unmatched braces are analyzer errors rather than ordinary text.
2. `{{` emits a literal `{` and `}}` emits a literal `}`. Escapes take precedence over expression delimiters during a left-to-right scan.
3. `{{Name}}` emits literal `{Name}`. `{{{Name}}}` emits a literal opening brace, the rendered `Name` value, and a literal closing brace.
4. Apply the same tokenization across adjacent expressions and across expressions or escapes split between multiple `w:r` and `w:t` nodes.
5. Unknown expressions, unsupported syntax, invalid escaping, and unmatched braces produce precise analyzer diagnostics. Do not silently retain them as text.
6. Inserted JSON values, formatted values, and emitted literal braces are output only and must never be interpreted as template syntax.

## 5. Data semantics and diagnostics

1. An ordinary scalar field or `{.}` accepts JSON strings, numbers, and Booleans. Objects and arrays cannot be inserted as scalar text. Explicit null and the empty string render empty, including when a formatter is present.
2. Missing scalar values continue to follow `MergeOptions.MissingValues`; this policy does not apply to structural dependencies.
3. `For` requires an array when a non-null value exists. An empty array, missing path, or explicit null produces zero iterations. Arrays may contain objects, strings, numbers, Booleans, null, or any mixture of those kinds; no homogeneous element type is required.
4. `If` requires a Boolean when a non-null value exists. Missing or explicit null selects the false or `Else` branch.
5. A missing `For` or `If` path produces a generation warning containing the original expression, canonical path, document part, and source location. Explicit null produces no warning.
6. When a property expression such as `{Name}` or `{./Name}` is evaluated with a string, number, Boolean, or null current element, resolve that property as missing. Apply `MergeOptions.MissingValues` when the expression is scalar, and apply the structural missing-path behavior in items 3 through 5 when the expression is used by `For` or `If`.
7. Do not coerce a scalar or null current element into an object and do not expose implicit properties. Property resolution stops with the explicit missing result.
8. Preserve the internal distinction between a missing property, an explicitly present null property, and an explicitly present non-null value of the wrong structural type. A present null scalar renders empty; a present null structural value uses the quiet empty/false behavior; a present non-null invalid structural value is an error.
9. Avoid duplicate missing-structural warnings for the same expression occurrence during repeated loop iterations where practical.
10. Restore the enclosing context after each loop iteration and nested loop. Root and parent references must remain stable through any supported nesting of `For` and `If`.
11. Evaluate each array element against the expressions actually used in the loop body. `{.}` applies ordinary scalar rules to that element, including built-in formatting; property, nested `For`, and `If` expressions retain their existing path, collection, condition, missing-value, and type rules.

## 6. DOCX structural boundaries

1. Process scalar expressions and supported structural blocks only in ordinary Word paragraphs and tables, including nested tables, within the supported story parts. Control markers are legal only in dedicated marker paragraphs or dedicated marker rows. A marker container contains exactly one control marker and no other visible content; whitespace is allowed.
2. A paragraph-delimited block has opening, optional `Else`, and closing markers in sibling `w:p` elements under the exact same block-level parent.
3. A paragraph-delimited branch may contain mixed sibling block elements, including paragraphs and complete tables. Rendering clones or removes only that sibling range and does not traverse to a common ancestor or restructure OOXML.
4. A table-row block has its opening, optional `Else`, and closing markers in sibling `w:tr` elements under the exact same `w:tbl`.
5. A row marker expression must be wholly contained in one paragraph in one cell. Every other visible item in the marker row must be empty. Marker rows are removed from generated output.
6. Table-row branches contain only sibling rows from that table. Cell-level or horizontal repetition is not supported.
7. Reject paragraph-to-row boundaries, markers with different structural parents, common-ancestor row promotion, mixed-content marker containers, multiple control markers in one container, crossing or overlapping regions, and mismatched or misordered markers.
8. Nested blocks are valid only when each block independently satisfies these rules and is wholly contained in one parent body or branch.
9. An outer row block may contain complete nested tables because the entire outer row is cloned. A block inside a nested table operates only on sibling rows of that nested table. Markers in an outer table and a nested table cannot form one block.
10. Paragraph-delimited blocks may clone or select an intact table, including a table with vertical merges, because the complete table structure moves together.
11. Preserve paragraphs, runs, styles, tables, rows, cells, images, drawings, relationships, and other unrelated XML. Preservation does not imply specialized interpretation. Marker-only containers and unselected branches must not remain in the output.
12. Inline control blocks within a paragraph or run are not supported. Scalar substitutions remain inline within ordinary text.

13. Match bookmark and permission endpoints, and comment start/end/reference attachments, by source identity within independent story roots. Their ordered ownership sequences must agree: root, exact loop body, exact conditional branch, or one deleted marker container. Preserve complete local ranges (including overlapping pairs with identical ownership), enclosing ranges, and reference-only point comments. Reject crossing, split-branch, deleted-marker-to-body, duplicate, or missing affected endpoints before rendering with `template.invalid_structure`, identifying both endpoints and the conflicting control. Each footnote/endnote is an independent root. Self-contained marker ranges may disappear; surviving links to them are unsupported.

## 7. Merged cells

1. Preserve horizontal merged cells represented by `w:gridSpan` in repeated or conditional rows.
2. Reject a row-level `For` or `If` operation whose affected rows contain or intersect any vertical merge represented by `w:vMerge`.
3. A vertical merge intersects a row operation when the operation contains a restart or continuation row, begins inside a merge started before the region, or ends before a merge continues outside the region.
4. Report analyzer errors identifying the affected table, rows, and relevant marker locations. Do not clone and attempt to repair such structures.
5. Do not reject vertical merges elsewhere in the document or inside a complete table cloned intact by a paragraph-delimited block.
6. Supporting complete vertical-merge row groups is deferred to a future feature.

## 8. Word lists

1. Numbered, bulleted, and multilevel list paragraphs are supported inside `For` and `If` regions, including nested lists and lists in mixed paragraph/table blocks.
2. Preserve complete paragraph properties, including `w:numPr`, `w:numId`, `w:ilvl`, indentation, styles, and formatting.
3. Preserve the existing numbering and abstract-numbering parts. Do not allocate, duplicate, or rewrite numbering definitions or numbering instances.
4. Word determines continuation according to the template's native numbering rules and the resulting paragraph order.
5. Feature 010 does not guarantee or implement a numbering restart for each loop iteration and introduces no numbering syntax.

## 9. Parse-once architecture

1. Parse every supported DOCX story once into an immutable syntax and render plan before rendering that story.
2. The plan contains source spans, scalar and escaped-brace operations, structural nodes and branches, canonical paths, and source locations sufficient to render clones without tokenizing them again.
3. Analyzer and Renderer use the same parser, grammar, path normalization, semantic model, and structural validation rules.
4. Rendering clones structural XML with a mapping from planned source nodes or relative locations to clone nodes, then recursively walks the existing child plan.
5. `For` renders its pre-parsed body for each element using an operation-scoped pushed context. `If` selects one pre-parsed branch without changing context.
6. Never re-tokenize cloned content, inserted data, formatted output, or emitted escaped braces.
7. Parsed plans are immutable. Data contexts, warnings, errors, and intermediate XML state belong to the individual generation operation and must not leak across concurrent or repeated use.
8. Do not retain the current behavior that reparses cloned collection fragments during each iteration.

## 10. Analyzer and public API

1. Replace the MERGEFIELD-specific analyzer `fields` and `formats` contract with `requirements` and `occurrences`; do not provide compatibility adapters.
2. `requirements` is a recursive, context-aware JSON dependency tree rather than a complete inferred JSON schema. Each node contains only `name`, `type`, and optional `fields`.
3. Requirement types are `object`, `array`, `value`, and `boolean`. For an `object`, `fields` contains dependencies on that object's properties. For an `array`, `fields` contains dependencies on properties of its elements, not properties of the array itself. Array fields describe expressions used by the template and do not require every element to be an object or to contain every listed property.
4. Resolve implicit current, explicit `./`, current-element `{.}`, repeated parent, and root references to their canonical locations before inserting or deduplicating requirements.
5. Deduplicate requirements by canonical path. Order top-level requirements by the first source occurrence of each canonical dependency and recursively apply the same first-occurrence rule to every `fields` array. A repeated reference contributes another occurrence but never moves an existing requirement node.
6. `occurrences` contains every dependency-bearing scalar field, `For`, and `If` expression, including duplicates, in source order. `{.}` is a scalar field occurrence rather than a new occurrence kind. Each occurrence contains its kind, original expression, canonical path, document part, and source location.
7. Do not include valid `Else`, `EndFor`, or `EndIf` markers in public occurrences because they introduce no data dependency. Retain their locations internally for validation and diagnostics.
8. Serialize `fields` only when a requirement has one or more child dependencies. Omit `fields` when there are none; never serialize an empty `fields` array.
9. An array requirement with no child fields is valid and makes no claim about its element type. For `{For:Tags}` containing `{.}`, serialize the requirement as `{"name":"Tags","type":"array"}` and record the scalar occurrence with canonical path `$.Tags[]`. Do not create a child such as `Value` or add a requirement type.
10. Mixed arrays may contain objects, primitives, and null even when an array requirement has `fields`. Requirements describe expressions used by the template rather than a homogeneous input schema. During rendering, property access against primitive or null elements follows the approved missing-property rules.
11. An array, object, value, or Boolean requirement that is structurally incompatible with another use of the same normalized property path remains an analyzer error. A Boolean dependency used by both `If` and ordinary scalar output is compatible.
12. Report every conflicting occurrence with its own part and location. Do not silently choose one type. Do not treat an array and fields of its elements as conflicting.
13. Preserve `valid`, `errors`, and `warnings` semantics while updating stable diagnostic codes and OpenAPI descriptions for the new language and structures.
14. Source locations must identify relevant DOCX parts, paragraphs, tables, rows, and cells sufficiently to correct invalid templates without exposing server filesystem paths.
15. Keep the contract minimal and suitable for later form generation, but do not implement form generation in this Feature. Do not add `items`, `elementType`, `required`, `nullable`, `format`, or other schema properties, and do not infer additional JSON-input restrictions from the dependency tree.

### 10.1. Normative requirements example

1. Given this template:

   ```css
   {Customer.Name}
   {For:Orders}
     {Number}
     {If:Paid}
       Оплачено
     {EndIf}
     {For:Items}
       {Product}
     {EndFor}
   {EndFor}
   ```

2. The Analyzer returns this `requirements` value:

   ```json
   [
     {
       "name": "Customer",
       "type": "object",
       "fields": [
         { "name": "Name", "type": "value" }
       ]
     },
     {
       "name": "Orders",
       "type": "array",
       "fields": [
         { "name": "Number", "type": "value" },
         { "name": "Paid", "type": "boolean" },
         {
           "name": "Items",
           "type": "array",
           "fields": [
             { "name": "Product", "type": "value" }
           ]
         }
       ]
     }
   ]
   ```

3. In this result, `Orders.fields` describes dependencies evaluated against order elements, and `Items.fields` describes dependencies evaluated against item elements. The canonical path of `Product` is `$.Orders[].Items[].Product`. These fields neither assert that all elements are objects nor make the properties mandatory on every element.

## 11. Supported stories and OOXML integrity

1. Preserve current coverage of the main document, headers, footers, footnotes, and endnotes.
2. Blocks cannot cross document parts or independent story roots.
3. Use one DOCX text-span adapter for supported ordinary paragraph and table content. Do not add special-purpose content parsers, a universal OOXML traversal framework, public content-scope options, or speculative extension points.
4. Preserve hyperlink content and relationships; explicit internal anchors follow the source/scope contract in section 11.12. Do not promise scalar substitution or structural markers inside hyperlink text, and never concatenate an expression token across a hyperlink boundary.
5. Do not interpret template expressions inside shapes, floating text boxes, SmartArt, charts, or other embedded or specialized content. Preserve these structures unchanged. Their presence, and the presence of images, drawings, hyperlinks, or other unsupported structures, is not itself an error.
6. Treat Word fields as opaque content. Preserve their instructions and cached results unchanged and never interpret either as MonoReport syntax. This includes MERGEFIELD instances remaining in a document; removing obsolete MERGEFIELD processing means deleting the legacy recognition and rendering mechanisms, not deleting field XML from generated documents.
7. Require tracked revisions to be accepted or rejected before processing. Detect recognized WordprocessingML revision markup during existing traversal of supported paragraph and table content whenever practical, including insertion, deletion, move, and property-change elements. If any is present, reject the template with a location-bearing Analyzer structural diagnostic; Analyzer and Generate must reject before producing output or performing irreversible document mutations. Do not attempt revision-aware parsing or rendering, and do not require a separate validation traversal when the existing traversal can detect revisions reliably.
8. Do not build a general-purpose OOXML expression scanner solely to find expressions in unsupported regions. When existing supported-content traversal can identify one reliably without substantial additional complexity, report an appropriate Analyzer diagnostic. Otherwise preserve it unchanged; Analyzer and Renderer make no claim that such an expression was discovered or substituted.
9. Do not flatten or rewrite unrelated XML while substituting fields, unescaping literal braces, cloning blocks, or removing branches.
10. Generated DOCX packages must pass applicable Open XML validation and open in Microsoft Word without repair warnings.
11. Remove all processed scalar expressions, control markers, and escape syntax from supported content. Expressions in unsupported content may remain unchanged as described in item 8.

12. Bind explicit internal `w:hyperlink/@w:anchor` references to source bookmarks before cloning. The target loop sequence must be a prefix of the reference loop sequence, and each target conditional branch must enclose the reference. Support local forward/backward and inner-to-ancestor references; reject enclosing-to-inner, sibling-loop, outside-to-repeated, potentially removed, and ambiguous affected targets in shared preflight, irrespective of runtime cardinality. Preserve unrelated unresolved anchors and external relationships. Preserve non-repeated bookmark names; allocate unique package-wide names of at most 40 characters for repeated bookmarks and cloned comment-definition bookmarks. Opaque Word fields, including `REF`, `PAGEREF`, and field-code hyperlinks, are not rewritten and are not guaranteed to resolve after structural operations.
13. All `w:sdt` content controls are opaque. Preserve complete subtrees, properties, data binding, and relationships; exclude their descendants from expression scanning and structural planning, with hard expression boundaries at entry and exit. Presence alone is not an error; visible outside fragments retain normal syntax diagnostics. Revision preflight remains independent of opacity.
14. Reject the documented WordprocessingML content, cell, move/custom-XML range, deleted-payload, and property-change revision families, plus Office 2010 conflict revisions, as specified in review section 16.6. Preserve source order and nearest paragraph/table/row/cell/property coordinates and shared Analyze/Generate rejection. Do not classify ordinary bookmarks, permissions, comments, proofing/session metadata, vertical merges, or the setting to track future edits as revisions.

## 12. Legacy removal

1. Remove MERGEFIELD recognition and rendering, Word field-code switch parsing, cached-result updating, legacy logic that preserved processed template field codes for re-merge, re-merge behavior, and technical `TableStart:` and `TableEnd:` interpretation. This removes processing mechanisms, not opaque Word field XML.
2. Remove obsolete field, switch, formatter-extension, cached-result, and region models or services after tracing their actual usages.
3. Retain reusable JSON resolution, package/story traversal, formatting algorithms, diagnostics, and structural XML handling where they fit the new architecture.
4. Do not implement two parsers, a MERGEFIELD fallback, an unused generic metadata framework, or an XLSX renderer.
5. Remove stale MERGEFIELD and custom-formatter extension guidance from samples and public documentation, except for a concise breaking-change note.

## 13. Tests and acceptance criteria

1. Scalar tests cover simple and nested paths, current and explicit-current property paths, root paths, repeated parent paths, missing values, null, unsupported scalar values, ordinal case sensitivity, ASCII identifiers, Cyrillic and other Unicode identifiers, and invalid identifiers.
2. Tokenization tests cover every expression form, every formatter, escaped braces, `{{Name}}`, `{{{Name}}}`, adjacent expressions, malformed syntax, unmatched braces, invalid escaping, and expressions or escapes split across multiple OOXML runs.
3. Parse-once tests prove that repeated and nested blocks do not re-tokenize clones and that JSON or formatted values resembling expressions or brace escapes remain inert.
4. Loop tests cover zero, one, and many elements; object, string, number, Boolean, null, and mixed elements; strings rejected as collections rather than iterated by character; nested arrays; empty inner collections; missing and null arrays; wrong collection types; parent restoration; root access; and warning deduplication.
5. Condition tests cover true and false, with and without `Else`, nested conditions, `If` in `For`, `For` in `If`, missing and null conditions, wrong types, and the rule that `If` does not change context.
6. Analyzer tests assert the exact `requirements` JSON in the normative example; object-field and array-element-field meaning; normalized canonical paths including `$.Orders[].Items[].Product`; first-source-occurrence ordering at the root and every nested `fields` level; repeated-reference deduplication without node repositioning; preservation of every occurrence, including duplicates, in source order with its original expression and source location; `{.}` normalized to the containing array-element path; exact `{"name":"Tags","type":"array"}` serialization without `fields`; mixed-array behavior; compatible Boolean `If` and scalar use; incompatible requirement shapes; every conflicting occurrence location; and the absence of closing markers from occurrences.
7. Paragraph-block tests cover single paragraphs, consecutive paragraphs, mixed paragraph/table sibling blocks, complete repeated sections, conditional complete tables, marker removal, and preservation of styles and runs.
8. Row-block tests cover zero, one, and many rows, horizontal merged cells, nested tables, nested loops inside tables, conditional rows, and precise rejection of cross-parent or paragraph-to-row boundaries.
9. Vertical-merge tests cover merges inside affected rows, merges crossing each boundary, unaffected merges elsewhere, intact tables cloned by paragraph blocks, and invalid nested row operations inside those tables.
10. List tests cover repeated numbered and bulleted items, multilevel and nested lists, conditional items, numbering references, indentation, formatting, and natural Word continuation without a restart guarantee.
11. Structural-error tests cover inline controls, mixed-content marker containers, multiple markers per container, duplicate or misplaced `Else`, unmatched and misordered markers, crossing regions, cell-level repetition, and common-ancestor row layouts.
12. Integration tests cover Core, CLI, HTTP generation and analysis, OpenAPI schemas, UI analyzer types and generated sample JSON, playground examples, all supported stories, deterministic diagnostics, concurrent/repeated use, and removal of obsolete mechanisms.
13. Representative real DOCX fixtures must pass Open XML validation. Generated fixtures and the authoritative sample must be opened in Microsoft Word or another Word-compatible viewer when available; Microsoft Word verification should cover numbering display, nested lists, intact merged tables, save/reopen behavior, and repair warnings.
14. All relevant existing tests must be updated rather than left as a parallel MERGEFIELD suite, and the final solution build and normal backend/frontend verification suites must pass.
15. Current-element tests cover `{.}` for strings, numbers, Booleans, and null; rejection of object and array scalar rendering; every applicable built-in formatter; `{.}` outside `For`; the distinction between `{.}` and `{./Name}`; nested loops; canonical paths; and JSON values containing template-like text under parse-once rendering.
16. Mixed-array property tests evaluate `{Name}` and `{./Name}` against string, number, Boolean, and null current elements under every supported `MergeOptions.MissingValues` policy. Assertions distinguish missing-property outcomes from explicitly present null properties and invalid non-null structural values.
17. Structural mixed-array tests verify that missing `For` and `If` properties on scalar or null current elements use structural warnings and empty/false rendering independently of `MergeOptions.MissingValues`, while present null values remain quiet and present non-null values of the wrong structural type remain errors.
18. Formatter grammar tests verify every built-in formatter on property values and compatible current elements; exactly-one-formatter enforcement; invalid, missing, and extra arguments; unknown formatters; formatters on structural expressions; statically invalid .NET number and date patterns; and the absence of custom registry or pipeline behavior.
19. Text-formatter tests cover strings, JSON numbers, and Booleans; lowercase Boolean text; culture-sensitive `upper` and `lower`; Unicode word-initial behavior for `caps`; first-Unicode-letter behavior for `firstcap`; preservation of every otherwise unchanged character; and object/array rejection.
20. Number-formatter tests cover JSON numbers and numeric strings parsed with `NumberStyles.Float` and `CultureInfo.InvariantCulture`; invariant decimal notation; positive and negative exponent notation; leading and trailing whitespace; decimal precision; culture-sensitive output; and valid standard and custom .NET patterns. They reject thousands or group separators, culture-specific decimal separators, non-finite values, invalid strings, Booleans, objects, arrays, and decimal overflow.
21. Date-formatter tests cover date-only input, `Z`, positive and negative explicit offsets, one-through-seven-digit fractional seconds, output-culture differences, preservation of the supplied offset, and deterministic date-only midnight at offset zero. They reject date-times without offsets, culture-specific or ambiguous dates, space-separated date-times, unsupported ISO forms, invalid calendar values, Booleans, numbers, objects, and arrays.
22. Empty-input tests prove that explicit null and empty strings render empty without conversion for every formatter while missing expressions still use every supported `MergeOptions.MissingValues` policy. Analyzer tests prove that invalid formatter syntax and statically invalid patterns are reported even when corresponding data would be null or empty.
23. `spellout` tests preserve existing positive, negative, grammatical, rounding-away-from-zero, and numeric-kopeck behavior; cover zero and the supported lower and upper boundaries; and apply the same numeric-string cases as `number`, including accepted invariant decimal, exponent, and surrounding-whitespace forms and rejected group separators, culture-specific decimal separators, non-finite values, invalid strings, and overflow. They also reject the first out-of-range rounded ruble amount, Booleans, objects, and arrays; prove culture-independent RUB-only output regardless of `MergeOptions.Culture`; and verify that incompatible inputs produce aggregated generation diagnostics rather than uncaught exceptions.
24. Supported-content fixtures verify scalar replacement and paragraph- and row-level `For`/`If` rendering in ordinary paragraphs, tables, and nested tables across every supported story, while preserving styles, runs, lists, numbering references, merged cells, images, drawings, relationships, and unrelated OOXML.
25. Hyperlink fixtures verify that explicit anchors follow section 11.12 while relationships and visible hyperlink text remain unchanged; expression-looking text inside a hyperlink remains unprocessed; and tokens split by a hyperlink boundary are not concatenated into an expression.
26. Unsupported-content fixtures verify that shapes, floating text boxes, SmartArt, charts, and other embedded or specialized structures are preserved unchanged and do not make an otherwise valid ordinary document fail. Tests distinguish any diagnostic reliably produced by supported traversal from the documented case where an unsupported expression remains undiscovered and unchanged.
27. Word-field fixtures verify that unrelated fields, their instructions, and cached results are preserved byte-for-byte where practical or semantically unchanged, that expression-looking cached text is inert, and that remaining MERGEFIELD instances are not processed or deleted.
28. Revision fixtures cover insertion, deletion, cell changes, both ends of move/custom-XML revision ranges, deleted payloads, property changes including table exceptions, and Office 2010 conflict revisions in supported paragraph and table content. Analyzer and Generate must reject the template before output or irreversible mutation with deterministic, location-bearing structural diagnostics, while revision-free documents continue normally. Behavioral tests must not prescribe or depend on a separate validation pass.

29. Range/reference fixtures verify the ownership and survival contracts in sections 6.13 and 11.12, including zero/one/many iteration independence, both conditional branches, local/ancestor/ambiguous links, package-wide drawing identity uniqueness in cloned comments, and semantic per-iteration anchor targets. Analyze/Generate diagnostic parity and no partial outputs apply through Core, CLI, and HTTP.
30. Content-control fixtures verify complete opacity/preservation and hard entry/exit boundaries for inline, block, row, cell, nested, adjacent, and empty controls, including when an enclosing operation moves a complete control. All seven persisted review inputs remain unchanged and source-valid; successful corrected outputs pass Office 2010 validation.

## 14. Samples and documentation

1. Update `Samples/sample.template.docx` and `Samples/sample.json` as the single authoritative sample sources.
2. Keep UI bundling through `PlaygroundExamplesPlugin`; do not create duplicate UI sample files.
3. Keep README guidance concise: a Template Syntax table, a minimal usage example, analyzer response summary, structural and numbering limitations, and a brief breaking-change note.
4. Ensure the playground, Analyzer, HTTP API documentation, OpenAPI contract, CLI behavior, sample files, and README describe the same supported syntax.

## 15. Explicitly deferred

1. Inline loops or conditions inside a paragraph or run.
2. Automatic separators between loop iterations.
3. `Else` within `For` for an empty collection.
4. Comparisons, Boolean operators, negation, calculations, functions, formatter chaining, and formatter pipelines.
5. Quoted or bracketed paths and arbitrary JSON property names.
6. Cell-level or horizontal repetition of arbitrary DOCX structures.
7. Row-region support for vertical merged cells, including otherwise complete merge groups.
8. Per-iteration list-numbering restarts.
9. A replacement custom formatter extension API or registered template formatters; this remains eligible for a separately approved future Feature.
10. XLSX rendering.
11. Form generation from Analyzer output.

## 16. Delivery workflow and definition of done

1. Obtain explicit Product Owner approval of this complete Feature Spec and its implementation plan before changing application code, tests, samples, or public documentation.
2. Implement incrementally with behavior-focused automated tests and representative real DOCX fixtures.
3. Before review, inspect the highest-risk parse-once and DOCX structural assertions, map every acceptance criterion to executed evidence or a justified manual check, and run the final solution and frontend verification suites.
4. Report changed public contracts, removed capabilities and code, verification results, manual-check limitations, and migration implications.
5. Do not stage, commit, tag, push, merge, publish, or release without separate Product Owner authorization.
6. The Feature is done only when Analyze and Generate use the new syntax and shared parse-once model; supported nested loops, conditions, mixed blocks, tables, and lists behave as specified; obsolete MERGEFIELD and formatter-extension mechanisms are removed; analyzer/API/sample/documentation contracts agree; generated DOCX files remain valid; and all required verification passes or disclosed limitations receive Product Owner acceptance.

## 17. Approved final correction decision on 2026-10-18

1. The Product Owner explicitly approved D1 (range ownership), D2 (static explicit-anchor scope/survival with opaque-field limitation), D3 (fully opaque content controls), Implementation Plan section 13, and implementation of F1–F4 and R1 in the current request. Sections 6.13 and 11.12–11.14 make those contracts normative; historical review sections 15–16 remain unchanged.
2. Acceptance additionally requires the regression matrix in Implementation Plan section 13.3: all seven unchanged persisted inputs, range ownership, scoped anchor semantics, package-wide clone identities including comments, revision families, content-control boundaries/preservation, diagnostic/transport parity and no partial outputs, Office 2010 validation, and all five example generations. Existing acceptance requirements remain in force.
3. Implementation authorization is not acceptance. Feature 010 remains FAIL / Pending Acceptance until independent verification and explicit Product Owner acceptance. No new independent review or Git publishing operation is authorized.
