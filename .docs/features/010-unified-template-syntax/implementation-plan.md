# Feature 010 — Implementation Plan

## 1. Summary and gates

1. Treat Feature 010 as a high-risk breaking Feature because it replaces template syntax and public analyzer contracts and changes DOCX structural processing.
2. Once approved, use the Feature Spec as the sole product-requirement source. Stop and obtain Product Owner approval before materially changing its grammar, contracts, structural rules, verification scope, or deferred work.
3. The Product Owner approved the complete Feature Spec and this plan before implementation began.
4. No independent Reviewer or Orchestrator is assumed. Add either only with Product Owner agreement.
5. Preserve user changes and follow the repository Git restrictions. Implementation approval does not authorize staging, commits, tags, pushes, merges, releases, or publication.

## 2. Shared language model and parse-once pipeline

1. Introduce one format-neutral parser and immutable semantic/render-plan model for fields, brace escapes, `For`, `If`, `Else`, and closing markers.
2. Implement the approved Unicode identifier grammar, current-element `{.}`, current-property `./`, parent/root addressing, canonical path normalization, strict brace precedence, and exactly one approved formatter per scalar expression.
3. Parse each DOCX story's logical text once while mapping expression spans across `w:r` and `w:t` nodes back to their source XML and diagnostic locations.
4. Build nested structural plan nodes only after validating marker pairing, `Else` ownership, dedicated marker containers, sibling boundaries, overlaps, and requirement compatibility.
5. Keep the parsed plan immutable. Store JSON context stacks, diagnostics, warning deduplication, cloned-node mappings, and output mutations in a separate per-generation context.
6. Share the parser, canonicalizer, semantic model, and structural validation between Analyzer and Renderer. Do not retain a legacy parser or construct a parallel analyzer model.
7. Render structural clones through precomputed node identities or relative source mappings. Never parse cloned XML or scan inserted JSON, formatted text, or emitted literal braces.
8. Render recursively: `For` clones and renders its body plan per JSON array element with the entire object, string, number, Boolean, or null element pushed as context; `If` selects and renders one pre-parsed branch without pushing context; child nodes inherit the correct root and ancestor contexts. Never treat a string as an enumerable collection.

## 3. Data resolution and formatting

1. Extend data resolution to distinguish missing, explicit null, wrong type, and present values while resolving property paths and `{.}` to canonical paths.
2. When property resolution encounters a string, number, Boolean, or null current element, return the explicit missing state without coercing the element or synthesizing properties.
3. Apply `MergeOptions.MissingValues` to scalar property expressions that resolve missing, including `{Name}` and `{./Name}` inside mixed arrays.
4. Apply structural missing-path behavior to missing `For` and `If` properties regardless of `MergeOptions.MissingValues`: record the approved warning and render zero iterations or the false/`Else` branch.
5. Keep explicit null distinct from missing: present null scalar properties render empty, and present null structural properties render empty/false without a missing warning. Keep present non-null values of the wrong structural type as recoverable generation errors under the existing error aggregation policy.
6. Deduplicate missing-structural warnings by dependency occurrence during repeated rendering where practical while retaining the expression, canonical path, part, and location.
7. Replace Word switch handling with the fixed built-in formatters `upper`, `lower`, `caps`, `firstcap`, `spellout`, `number`, and `date`.
8. Enforce exactly one formatter in the grammar and semantic model. Do not introduce a formatter pipeline; treat chaining as deferred and direct consumers needing complex transformations to preprocess JSON.
9. Reuse the existing culture-aware casing and Russian RUB spell-out algorithms where they match the Feature Spec. Preserve the exact `upper`, `lower`, `caps`, and `firstcap` character transformations for the ordinary textual representations of strings, JSON numbers, and Booleans.
10. Implement one shared numeric-input conversion for `number` and `spellout`, with explicit value-kind checks and decimal arithmetic. Parse numeric strings with `decimal.Parse` or equivalent decimal APIs using `NumberStyles.Float` and `CultureInfo.InvariantCulture`, accepting invariant decimal notation, exponent notation, and surrounding whitespace within the decimal range. Reject thousands or group separators, culture-specific decimal separators, non-finite values, invalid strings, and overflow through normal generation diagnostics. Format `number` output with the supplied .NET decimal pattern and `MergeOptions.Culture`.
11. Replace permissive date parsing with an exact invariant ISO 8601 parser for the approved date-only and offset-bearing date-time forms. Represent date-only values as midnight at offset zero, preserve explicit offsets, reject offsetless date-times rather than consulting local time, and apply `MergeOptions.Culture` only when formatting output with the supplied .NET date/time pattern.
12. Feed `spellout` through the same shared numeric-input conversion as `number`, then preserve the existing RUB-only algorithm, two-decimal away-from-zero rounding, Russian grammatical forms, numeric kopecks, negative handling, and absolute rounded-ruble limit below 10^18. Keep its output independent of `MergeOptions.Culture`. Add explicit kind, parsing, overflow, and range handling so Booleans and all other incompatible inputs become normal generation diagnostics. Do not add currency selection.
13. Short-circuit explicit null and empty strings to empty output before formatter conversion. Continue to route missing expressions through `MergeOptions.MissingValues`.
14. Validate formatter grammar and statically invalid number/date patterns during analysis independently of runtime data. Reserve generation diagnostics for actual value kinds, conversion, overflow, and range failures, and preserve error aggregation and no-partial-document behavior.
15. Remove `MergeOptions.FieldFormatters`, `FormatterBinding`, public custom formatter registration, registry lifecycle behavior, arbitrary option dictionaries, and associated DI registrations, diagnostics, tests, and documentation. Do not design a replacement API in this Feature; record extensibility as possible future work.
16. Preserve `MergeOptions.Culture`, `MergeOptions.MissingValues`, error aggregation, no-partial-document behavior, and relevant formatter failure diagnostics.
17. Render `{.}` through the ordinary scalar and formatter path. Accept string, number, and Boolean elements; render null empty; reject object and array elements as scalar values.
18. Allow mixed element kinds without inferring or enforcing array homogeneity. Evaluate each element against the field, `{.}`, nested `For`, and `If` expressions actually executed for that iteration.

## 4. DOCX adapter and structural rendering

1. Replace MERGEFIELD discovery and cached-result rewriting with one DOCX text-span adapter scoped to ordinary paragraph and table content. It must replace planned scalar and escape spans while retaining unrelated runs and run properties; do not introduce special-purpose parsers, universal OOXML traversal, public content-scope options, or speculative extension points.
2. Implement paragraph-delimited regions as exact sibling ranges under one block-level parent. Allow the branch range to contain mixed sibling paragraphs and complete tables.
3. Implement row-delimited regions as exact sibling `w:tr` ranges under one `w:tbl`. Require one marker expression in one paragraph in one cell of an otherwise marker-only row.
4. Remove marker containers and unselected branches. Clone selected or repeated block XML without flattening paragraphs, tables, cells, styles, lists, numbering references, merged cells, images, drawings, relationships, properties, or unrelated OOXML.
5. Validate every nested block independently. Reject cross-parent, paragraph-to-row, common-ancestor promotion, inline, cell-level, crossing, overlapping, mismatched, multi-marker, and mixed-content-marker layouts.
6. Support nested tables by cloning complete outer rows or complete tables as ordinary content and by applying nested row plans only to the exact inner table that owns them.
7. Preserve `w:gridSpan` horizontal merges without rewriting them.
8. Detect `w:vMerge` at the structural-operation level. Reject row regions that contain a vertical merge or cut through its restart/continuation extent, with table, row, and marker locations. Do not reject or rewrite an intact vertically merged table cloned by a paragraph block.
9. Continue processing the main document, headers, footers, footnotes, and endnotes. Reject blocks that attempt to cross parts or story roots.
10. Define hard text-span boundaries around hyperlinks and opaque Word fields. Preserve each subtree and its relationships, do not scan its visible text, instructions, or cached results as template syntax, and never join token text across either boundary. A remaining MERGEFIELD is opaque Word content, not legacy template input to process or remove.
11. Exclude shapes, floating text boxes, SmartArt, charts, and other embedded or specialized content from supported text-span discovery and preserve their XML and relationships unchanged. Their presence is not an error. Do not add a second scanner merely to diagnose expression-looking text inside them; emit a diagnostic only where the normal supported-content traversal identifies the unsupported placement reliably.
12. Detect recognized WordprocessingML revision markup as part of the existing supported-content traversal wherever practical, including insertion, deletion, move, and property-change elements. Report the part and nearest reliable paragraph, table, row, or cell location, and make the shared analysis or generation preflight fail before output publication or irreversible document mutation. Do not normalize, accept, reject, or render revision alternatives.
13. Do not add a mandatory revision-only traversal, service, abstraction, or public option. A separate validation pass is allowed only if implementation evidence from the actual adapter and fixtures demonstrates that it is simpler or necessary for reliable coverage; record that evidence and keep the pass local to DOCX validation.

## 5. Word list behavior

1. Treat list items as paragraphs under the same paragraph-region rules; do not introduce a separate list renderer.
2. Preserve each paragraph's complete `w:pPr`, `w:numPr`, `w:numId`, `w:ilvl`, indentation, styles, and formatting while cloning or selecting it.
3. Preserve the document's numbering and abstract-numbering parts unchanged. Do not allocate numbering instances, calculate displayed numbers, or force restarts.
4. Verify numbered, bulleted, multilevel, nested, repeated, and conditional list items. Document that Word controls continuation and that per-iteration restart is not guaranteed.

## 6. Analyzer and public contracts

1. Replace the analyzer `fields/formats` output and its public models with recursive `requirements` and flat dependency `occurrences`.
2. Define requirement nodes with only `name`, `type`, and optional `fields`, using only `object`, `array`, `value`, and `boolean`. For an object, `fields` represents dependencies on that object's properties. For an array, `fields` represents dependencies evaluated against properties of its elements, not properties of the array itself. Keep the result a dependency tree rather than complete JSON schema inference.
3. Define occurrences with `kind`, original `expression`, canonical `path`, document `part`, and source `location`; expose only field, `For`, and `If` dependencies, representing `{.}` with the existing field kind.
4. Preserve every occurrence, including duplicates, in source order. Deduplicate requirements by canonical path after resolving `{.}`, current-property, parent, root, object, and array-element scopes; place each requirement at its first occurrence and never reposition it when the path repeats. Apply the same insertion-order rule independently to the root requirements and every nested `fields` collection.
5. Preserve the existing deterministic DOCX story traversal convention: main document first; headers next in ordinal part-URI order; footers next in ordinal part-URI order; footnotes next when present; and endnotes last when present. Within each story, traverse expressions in logical document order. Use this total source order for requirements, occurrences, and Analyzer diagnostics so repeated analysis is stable.
6. Omit `fields` from serialized requirement nodes with no child dependencies; never emit `"fields": []`. Normalize `{.}` inside a loop such as `{For:Tags}` to the array-element path `$.Tags[]` and serialize its collection requirement exactly as `{"name":"Tags","type":"array"}`. Do not add an artificial `Value` field, a new requirement type, or a homogeneous element-type claim.
7. Treat array `fields` as expression dependencies only. Do not infer that every array element is an object, contains every field, or excludes primitives or null; concrete mixed elements continue through the approved rendering and missing-property rules.
8. Detect incompatible property requirement shapes before generation without treating mixed array elements or `{.}` occurrences as complete schema declarations. Emit location-bearing analyzer errors for every actual conflict without choosing a winning type; allow a Boolean to be used both as an `If` dependency and scalar output.
9. Do not add `items`, `elementType`, `required`, `nullable`, `format`, or other schema properties to public requirements, and do not convert the dependency contract into JSON-input validation.
10. Preserve the analyzer's package/error isolation and stable `valid`, `errors`, and `warnings` concepts while replacing MERGEFIELD-specific codes and descriptions.
11. Update HTTP serialization, OpenAPI schemas, CLI analysis output, UI analyzer types, and generated-sample JSON logic for the new contract without compatibility adapters. Generated sample JSON may illustrate one shape satisfying the dependencies but must not present requirements as a homogeneous, required, or nullable schema.

## 7. Legacy deletion and integration updates

1. Trace and remove MERGEFIELD parsing, Word field-code switches, cached-result writers, processed-template field preservation and re-merge mechanisms, `TableStart:` and `TableEnd:` interpretation, and obsolete models and services. Preserve opaque Word field XML independently of deleting those mechanisms.
2. Remove custom formatter extension plumbing and tests while retaining reusable formatting algorithms needed by the fixed language.
3. Update dependency injection to compose the new parser, semantic analyzer, DOCX adapter, renderer, and built-in formatting services through constructor-supplied dependencies.
4. Update Core, CLI, HTTP API, OpenAPI, and UI integrations together so no supported surface exposes the old syntax or analyzer schema.
5. Update `Samples/sample.template.docx` and `Samples/sample.json` as the only authoritative samples and retain `PlaygroundExamplesPlugin` bundling without duplicated UI assets.
6. Replace README guidance with the concise approved syntax, minimal usage, analyzer contract, structural/list limitations, removed formatter-extension note, and breaking-change note.

## 8. Incremental implementation sequence

1. Add focused red tests for shared grammar, `{.}` context rules, canonical paths, strict escapes, requirement conflicts, scalar and mixed arrays, and parse-once inert-output behavior.
2. Implement the shared parser, immutable plan, canonicalization, and analyzer requirement builder until those tests pass.
3. Add DOCX fixtures and red tests for text-span replacement, paragraph and mixed-block regions, row regions, nested tables, merged cells, lists, hyperlink and Word-field boundaries, unsupported embedded-content preservation, and tracked-revision rejection.
4. Implement the DOCX adapter and recursive renderer without clone reparsing, then verify zero/one/many, nesting, formatting preservation, and structural diagnostics.
5. Replace formatting and data-resolution behavior, including exact formatter input domains, static pattern validation, missing structural warnings, and removal of the formatter extension API.
6. Migrate analyzer, API, CLI, UI, OpenAPI, and sample-generation contracts.
7. Remove obsolete production paths, models, fixtures, tests, and documentation only after their replacements are exercised.
8. Update the authoritative DOCX/JSON sample and concise README, then perform final acceptance mapping and verification.

## 9. Behavioral verification mapping

1. Grammar and path tests verify all scalar/control forms, `{.}` only within `For`, the lexical distinction between `{.}` and `{./Name}`, Unicode identifiers, invalid identifiers, current/root/repeated-parent resolution, canonical paths, adjacent expressions, strict escaping, malformed syntax, and split-run tokens.
2. Formatter grammar tests verify every built-in formatter; exactly-one enforcement; unknown, missing, extra, and multiple formatter syntax; rejection on structural expressions; static rejection of invalid .NET number/date patterns even with null or empty data; and the absence of custom registry or pipeline behavior.
3. Parse-once tests use JSON and formatted values containing `{For:X}`, `{If:X}`, `{{`, and `}}` to prove rendered data remains inert. Repeated rendering and concurrent service reuse verify no state leakage.
4. Loop tests cover zero, one, and many elements; object, string, number, Boolean, null, and mixed elements; rejection of strings as collections; `{.}` scalar rendering; object/array scalar errors; nested arrays; empty inner arrays; parent restoration; root access; missing/null/wrong collection types; warning context; and warning deduplication.
5. Mixed-array property tests evaluate both `{Name}` and `{./Name}` against string, number, Boolean, and null elements under every supported `MergeOptions.MissingValues` policy. They separately assert present-null scalar output, missing structural warnings and empty/false behavior, present-null structural behavior without warnings, and present non-null structural type errors.
6. Condition tests cover true/false, optional `Else`, nested conditions, both loop/condition nesting directions, missing/null/wrong types, and unchanged data context.
7. Analyzer contract tests assert the exact public `requirements` JSON from the Feature Spec's normative `Customer`/`Orders` example, including nested array `fields` and the canonical path `$.Orders[].Items[].Product`. Focused cases verify exact `{"name":"Tags","type":"array"}` serialization with omitted `fields`; object-property versus array-element-property meaning; mixed object/primitive/null arrays without inferred restrictions; `{.}` normalization to `$.Tags[]`; root and recursive first-occurrence ordering; canonical-path deduplication without repositioning after repeated references; preservation of duplicate occurrences in source order with their original expressions and locations; closing-marker exclusion; compatible Boolean scalar use; all actual incompatible shapes; and every conflicting location.
8. Paragraph/mixed-block fixtures verify scalar expressions and supported controls in ordinary paragraphs, tables, and nested tables; single and grouped paragraphs; complete sections containing tables; conditional tables; nested controls; marker removal; and preservation of runs, styles, lists, numbering references, tables, cells, merged cells, images, drawings, relationships, and unrelated XML.
9. Table fixtures verify row repetition/selection, exact-table ownership, nested tables, horizontal merges, vertical merges inside and crossing affected ranges, unaffected external merges, and intact vertically merged tables cloned as complete blocks.
10. List fixtures verify numbered, bulleted, multilevel, nested, repeated, and conditional items; exact `numPr`, `numId`, `ilvl`, indentation, style, and numbering-part preservation; and zero/one/many ordering.
11. Invalid-structure tests verify inline controls, mixed marker text, multiple markers, duplicate or misplaced `Else`, unmatched/misordered markers, crossing regions, cross-parent boundaries, paragraph-to-row boundaries, common-ancestor promotion, and cell-level repetition.
12. Story and integration tests verify main document, multiple headers in ordinal part-URI order, multiple footers in ordinal part-URI order, footnotes, and endnotes in the documented total traversal order. They assert stable cross-part requirements, duplicate occurrences, and diagnostics across repeated analysis, plus Core/CLI/API parity, HTTP problem details, the minimal OpenAPI requirement-node schema with optional omitted `fields`, UI analyzer behavior, dependency-based generated sample JSON without added schema claims, error limits, and no partial outputs.
13. Deletion checks verify that MERGEFIELD processing, field-code switch parsing, cached-result mutation and re-merge paths, technical region-field interpretation, custom formatter bindings, and compatibility fallbacks are absent rather than dormant, while opaque Word fields remain preserved.
14. Text-formatter tests verify string, JSON-number, and Boolean textual representations; lowercase Boolean text; culture-sensitive `upper` and `lower`; Unicode word-initial `caps`; first-Unicode-letter `firstcap`; unchanged remaining characters; and object/array errors.
15. Number tests verify JSON numbers and numeric strings parsed with `NumberStyles.Float` and `CultureInfo.InvariantCulture`; invariant decimal notation; positive and negative exponent notation; leading and trailing whitespace; decimal precision; culture-sensitive output; valid standard and custom .NET patterns; and supported boundaries. They reject group separators, culture-specific decimal separators, non-finite values, invalid strings, Boolean/object/array values, and decimal overflow through normal aggregated diagnostics.
16. Date tests verify the exact date-only and offset-bearing ISO 8601 forms, `Z`, positive and negative offsets, one-through-seven fractional digits, date-only midnight at offset zero, preservation of supplied offsets, and culture-sensitive output. They reject missing offsets, machine-local interpretation, ambiguous or culture-specific dates, space separators, unsupported ISO forms, invalid calendar values, and incompatible JSON kinds through normal diagnostics.
17. Empty-input tests verify that explicit null and empty strings bypass conversion for every formatter, missing values retain all supported `MergeOptions.MissingValues` outcomes, and static syntax/pattern analysis remains data-independent.
18. `spellout` tests preserve RUB-only positive, negative, grammar, rounding-away-from-zero, numeric-kopeck, zero, and boundary behavior and prove output remains independent of `MergeOptions.Culture`. They repeat the `number` numeric-string matrix to prove identical acceptance of invariant decimal, exponent, and surrounding-whitespace forms and identical rejection of group separators, culture-specific decimal separators, non-finite values, invalid strings, and decimal overflow. They also verify that the first out-of-range rounded amount, Booleans, objects, and arrays produce normal aggregated generation diagnostics rather than uncaught exceptions.
19. Content-boundary fixtures verify that hyperlinks and opaque Word fields retain their XML, relationships, visible text, instructions, and cached results; expression-looking content inside them remains inert; and expression tokens are never assembled across their boundaries. Remaining MERGEFIELD instances are preserved without invoking legacy processing.
20. Unsupported-content fixtures verify unchanged shapes, floating text boxes, SmartArt, charts, and representative embedded content alongside successfully processed ordinary paragraphs and tables. Their presence alone produces no diagnostic; tests separately cover a reliably detectable unsupported expression diagnostic when the normal traversal exposes one and unchanged, undiscovered expressions when it does not.
21. Revision fixtures cover insertion, deletion, move, and representative property-change elements in ordinary paragraphs, tables, and supported stories. They assert detection through existing traversal where practical, deterministic Analyzer and Generate rejection before output or irreversible mutation, precise available locations, no partial output, and successful processing after revisions are absent. If implementation evidence justifies a separate pass, add focused tests proving why the normal traversal is insufficient and that the local pass does not alter revision-free behavior.

## 10. Verification commands and evidence

1. Run focused tests after each implementation stage using filters for the affected Core, analyzer, region, diagnostics, CLI, API, or UI behavior.
2. Run `dotnet build MonoReport.slnx -c Release` for the final backend build.
3. Run `dotnet test MonoReport.slnx -c Release` for the final backend test suites.
4. From `MonoReport.UI`, run `npm test -- --run`, `npm run typecheck`, and `npm run build`.
5. Run focused real CLI and HTTP smoke checks for analysis and generation because both public behaviors and contracts change. Real-host checks need cover the changed behavior only; browser automation remains excluded without separate Product Owner approval.
6. Validate representative generated fixtures and the authoritative sample with the Open XML validator.
7. When Microsoft Word is available, manually verify representative outputs for continuous numbering, bullets, multilevel indentation, mixed sections, intact merged tables, save/reopen behavior, and absence of repair warnings. Record this as unexecuted rather than passed when Word is unavailable.
8. Store raw logs or TRX evidence outside Git where practical and record exact commands, results, evidence paths, and limitations in Feature 010's `review.md`.

## 11. Risks and controls

1. The highest parser risk is accidentally interpreting output as syntax. Control it with the immutable parse-once plan and inert-output regression tests.
2. The highest DOCX risk is corrupting structure while cloning mixed blocks or table rows. Control it with exact sibling ownership, source-to-clone mappings, real fixtures, semantic property assertions, package validation, and manual viewer checks.
3. The highest table risk is vertical-merge corruption. Control it by rejecting every affected row operation and allowing vertical merges only when complete tables are moved intact.
4. The highest list risk is asserting visible numbering from XML alone. Control it with exact numbering-property assertions and manual Word verification without promising per-iteration restarts.
5. The highest contract risk is leaving partial MERGEFIELD or formatter-extension compatibility. Control it with coordinated Core/API/CLI/UI migration and explicit deletion checks.
6. Formatter parsing risks include culture-dependent input interpretation, accidental local-time assumptions, and uncaught value-kind exceptions. Control them with exact invariant parsers, explicit JSON-kind dispatch, separate static pattern validation, offset-preservation assertions, and boundary tests.
7. The highest content-boundary risk is accidentally scanning nested specialized XML or Word-field cached results as ordinary text. Control it with one allowlisted ordinary-content adapter, hard hyperlink and field boundaries, revision detection integrated into existing traversal where practical, preservation assertions, and fixtures containing supported and unsupported content together. Add a separate local revision pass only when implementation evidence shows it reduces complexity or is required for reliable detection.
8. If implementation requires a material deviation from the approved grammar, public schema, structural boundaries, or deferred list, stop and request Product Owner approval before continuing.

## 12. Completion handoff

1. Before implementation review, reconcile every Feature Spec acceptance criterion against a meaningful automated test or justified manual check, with special inspection of parse-once and DOCX structural assertions.
2. Report changed public contracts, removed formatter capabilities, removed legacy code, supported and rejected structures, exact verification evidence, unexecuted checks, and migration implications.
3. Request Product Owner acceptance only after required implementation, documentation, sample updates, and verification are complete.
4. Do not stage, commit, tag, push, merge, publish, or release without separate explicit Product Owner authorization.
