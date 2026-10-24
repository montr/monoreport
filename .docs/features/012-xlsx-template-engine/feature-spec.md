# Feature 012 — XLSX Template Engine

**Status:** Approved baseline and Owner amendments remain authoritative. M3.14 and M4 are Owner-accepted; M4 was accepted on 2026-10-24 after public activation and the verification recorded in implementation-progress.md. The earlier bounded activation acceptance and its Git integrity confirmation are historical; recorded fingerprints remain historical evidence and are not represented as matching. M5 implementation completed on 2026-10-24 under Owner-authorized transport completion; verification is recorded in implementation-progress.md. The malformed-quality review correction is implemented; independent M5 re-review and formal Owner M5 acceptance remain pending. Microsoft Excel E1, M6, and final Feature 012 acceptance remain pending. Milestone implementation completion does not imply final Feature acceptance.
**Scope:** XLSX template processing through existing MonoReport Core, CLI and HTTP interfaces, JSON binding, nested two-axis expansion, and Excel-compatible output. Preserve existing DOCX contracts.

## 1. Goal

Build a C#/.NET engine that generates `.xlsx` reports from Excel-authored templates and JSON input without requiring Microsoft Excel or Office Automation at runtime. Users control report layout and appearance in Excel. The engine populates values and repeats explicitly marked template regions; it must not invent formatting, merge cells automatically, or infer presentation decisions on behalf of the template author.

## 2. Inputs and outputs

1. Input: XLSX template and a JSON object; optional generation settings may be added where justified.
2. Output: generated XLSX plus structured diagnostics, including non-fatal warnings on successful generation.
3. Support multiple worksheets and multiple tables on a worksheet.
4. Allow **at most one horizontally expanding region per worksheet** in the initial scope.
5. Do not require Excel, COM, Office Automation, or a proprietary runtime dependency.
6. Use only permissively licensed dependencies (e.g., MIT, Apache-2.0, BSD); exclude GPL/AGPL and commercial-only dependencies. The recommended runtime uses the existing `DocumentFormat.OpenXml` dependency and an internal bounded reference transformer; no ClosedXML dependency is required by the design.
7. Extend the existing `MergeEngine.Merge`, `TemplateAnalyzer.Analyze`, CLI generation/analysis commands and HTTP document endpoints. Do not create separate public XLSX APIs or XLSX-specific HTTP routes. Preserve existing DOCX signatures, construction/composition compatibility, inputs, outputs, diagnostic behavior, MIME type and filename behavior.
8. Identify the actual OOXML package format; do not trust an extension alone or send an XLSX package through WordprocessingDocument. Select the corresponding internal handler and output MIME/extension. No Playground or new UI capability is included.

## 3. Template conventions

### 3.1. Report and control areas

1. Excel's `PrintArea` defines retained report content. Users must not place report content, formula inputs or other required content outside it. Control markers live above and/or to its left.
2. All content outside PrintArea is disposable and must be absent from the final report. Physically remove external control rows/columns and crop the other external cells; do not merely hide them in the final output. Temporary hiding is permitted only in explicitly incomplete development phases.
3. Initially require one finite rectangular PrintArea on every retained worksheet. Missing, invalid or multiple/disjoint areas produce an actionable error; do not infer boundaries or silently discard a worksheet. Preserve and correctly transform intended print settings, coordinates and supported references through expansion and physical cropping.
4. Reject surviving formulas/references to discarded content when a correct referent cannot be retained inside PrintArea. No implicit cached-value substitution, general repair or external helper retention is permitted.

### 3.2. Data fields

1. Field expressions use braces, for example `{Name}`, `{Price}`, and `Lot: {Name}`.
2. Reuse the existing MonoReport DOCX data-resolution and field-formatting capabilities wherever applicable, rather than designing an incompatible expression language.
3. In a logical vertical row with multiple intersection records, ordinary row fields resolve from the **first input record** in that logical group; fields within an intersection resolve from its corresponding record.
4. Reuse existing scalar paths, case-sensitive JSON lookup, brace escapes, missing/null distinctions and built-in formatters. A complete unformatted scalar field retains a JSON number as an Excel number, a Boolean as a Boolean and a string as text. Null/empty output retains authored cell styling. An explicit formatter or surrounding literal text produces text. Objects/arrays are not scalar values. Text beginning with `=` remains text.
5. Preserve authored Excel number formats; do not translate .NET formatter patterns into Excel formats. Numeric identifiers that cannot be represented faithfully within Excel's numeric precision must be supplied/rendered as text; reject unsafe conversion rather than silently truncate them.
6. Preserve existing native Excel dates and workbook date-system settings. An unformatted numeric serial bound to an authored date-formatted cell is a native date in that workbook's 1900/1904 system. ISO date strings remain text; existing `date` formatting remains textual. Automatic ISO-to-serial conversion, timezone normalization and a new typed-date binding convention are deferred. This phased boundary avoids inventing a JSON date type or template syntax.
7. Ordinary fields, including data fields in parent-only columns, use the first logical-row record. Leaf intersection fields use their corresponding record. Header fields use the header level's collection element. Vertical parent paths follow actual data-ancestor first records, including ancestors not rendered by the template; they never use nearest rendered/source parents or implicitly fall back to the document root. Explicit `$.` accesses the document root. Horizontal header parent paths keep their separate preceding-header-level semantics; the two axes do not silently add each other's contexts. Vertical ancestor availability is data-dependent and cannot be validated from source nesting depth alone.
8. An absent property is not inherently an error in schema-less input: ordinary missing properties retain the public `MergeOptions.MissingValues` contract, with no additional warnings introduced by this amendment. Traversing beyond the actual vertical data root is a distinct invalid-ancestor condition: render empty and warn by default, including under `EmptyString`/`Warning`; use existing `Error` mode for strict failure. Diagnostic information must distinguish unavailable ancestry from a missing property on an existing ancestor, identify the field/source/current indexed item and avoid fabricated ancestor paths. Keep existing missing/null property behavior, bounded diagnostics and no partial output; add no configuration or diagnostic framework.

### 3.3. Vertical markers

| Long form | Short form |
| --- | --- |
| `$level(10).Items` | `$l(10).Items` |
| `$level(20).Items` | `$l(20).Items` |

1. The collection name is arbitrary; `Items` is illustrative.
2. `$level` is an integer. Values need not start at zero or be consecutive; a greater value represents a deeper level.
3. Markers are placed in the control column to the left of the report area.
4. Consecutive template rows with identical markers form **one multirow template block**, repeated as a unit for each matching logical item.
5. A parent-level block before child blocks is a group header; a parent-level block after child blocks is a group footer. Relative source placement determines presentation roles; it does not establish data parenthood or require source/data parent chains to match. Select each data item's level roles, emitting header/body on entry and footer after its actual data descendants. No separate `footer` keyword is required.
6. A vertical-only report may omit `$path`.
7. Consecutive rows containing vertical markers form one independent vertical template region. A row without a vertical marker terminates that region and separates it from a following one, even when markers on both sides are identical. Separator rows inside PrintArea remain once, outside both repetitions, and may have an author-defined zero height. No named report regions or additional boundary syntax are used.
8. Within one region, header/deeper-body/footer placement must admit one unambiguous recursive interpretation. Reject ambiguous interleavings with the remedy to separate independent regions using an unmarked control row. Equivalent long/short marker forms are normalized before coalescing identical-marker multirow blocks.

### 3.4. Horizontal markers

| Long form | Short form |
| --- | --- |
| `$header(1).Bids` | `$h(1).Bids` |
| `$header(2).Criteria` | `$h(2).Criteria` |

1. Horizontal header row markers are placed to the left of the report area. Their collections and level numbers define horizontal hierarchy.
2. Top control rows contain `$h(1)`, `$h(2)`, `$h(3)` and so on. A different control row represents each depth so parent and child markers can coexist over the same physical column. Consecutive same-level markers define a multicolumn block. Bare `h` is superseded and is not the new template convention.
3. Child blocks are contained in parent blocks. A parent may contain columns outside its children; those parent-only columns repeat once with the parent, not with a child. Counts come from the corresponding JSON collections, not authored column counts. Zero-width parent-only columns are valid physical columns and retain their authored width.
4. Model horizontal regions recursively, without a hardcoded two-level limit. Validate containment, level/collection associations, one outer horizontal region per worksheet and Excel dimension/resource bounds. Multiple child blocks at a level use that level's collection; repeated template positions do not require extra positional data records.
5. The author defines all cells, widths, formatting and merges. The engine may copy/move authored structures and make the approved bounded adjustment of an existing parent-owned merge to follow a fully contained expanded child region. It must not invent merges, merge unrelated cells, invent styling or stretch unrelated structures.
6. An existing merge that collapses to one surviving cell becomes an ordinary cell, preserving its content and formatting. A merge with no surviving owner disappears. Reject partial/cross-owner projections that cannot be represented safely as one rectangle. Preserve nonanchor formatting and do not discard nonanchor content silently.
   A zero-width surviving anchor retains its value and authored width/hidden state; the engine does not widen it to make an empty-child heading visible.
7. Consecutive rows with equivalent horizontal header markers form a multirow header block. A parent heading renders once per parent instance; child headings render once per child inside it. Do not duplicate the parent heading per leaf.
8. Mandatory fixture: parent D:E has `$h(1)` on both columns, child E has `$h(2)` on the next control row, D7:E7 contains merged `{Company}`, E8 contains `{Name}`, and D may have width zero. Two companies and three criteria produce D7:G7 and H7:K7 before external-margin removal; each company appears once above its own three criteria. Test zero, one and multiple criteria; crop these coordinates through the final PrintArea transform.
9. Header-row long/short forms remain equivalent. Keywords are case-insensitive; JSON property/collection lookup remains ordinal and case-sensitive. Geometry markers carry levels only; collection associations come from header-row markers.

## 4. Data model

### 4.1. Vertical-only example

```json
{
  "Items": [
    {"$level": 10, "Name": "Lot 1"},
    {"$level": 20, "Name": "Item A", "Price": 100},
    {"$level": 20, "Name": "Item B", "Price": 200},
    {"$level": 10, "Name": "Lot 2"},
    {"$level": 20, "Name": "Item C", "Price": 300}
  ]
}
```

1. Input ordering is authoritative. The engine **must not sort by `$path` or `$level`**.
2. `$level` controls hierarchical interpretation, not business semantics.
3. Intentionally unrendered input levels are normal and produce no warning. Omit only their own blocks, continue through their actual data descendants and retain all data ancestry for scalar scopes and instance identity. Reject only a genuinely ambiguous/unsafe requested result, not a difference in source/data hierarchy. Never synthesize missing objects, reparent data or deduplicate by values.
4. Template levels absent from input yield zero instances of every authored role, with no placeholder objects/physical rows or warning solely for the absent level.
5. Assume otherwise valid nesting in the initial implementation; no comprehensive hierarchy-validation feature is required yet.

### 4.2. Two-axis example

```json
{
  "Bids": [
    {"Number": 1, "Company": "Alpha", "Address": "City A"},
    {"Number": 2, "Company": "Beta", "Address": "City B"}
  ],
  "Criteria": [
    {"Name": "Price"},
    {"Name": "Term"}
  ],
  "Items": [
    {"$level": 10, "$path": "001", "Name": "Lot 1", "Price": 1000},
    {"$level": 10, "$path": "001", "Term": 30},
    {"$level": 10, "$path": "001", "Price": 950},
    {"$level": 10, "$path": "001", "Term": 35},
    {"$level": 20, "$path": "001/001", "Name": "Item A", "Price": 600, "Amount": 720},
    {"$level": 20, "$path": "001/001", "Term": 20},
    {"$level": 20, "$path": "001/001", "Price": 570, "Amount": 684},
    {"$level": 20, "$path": "001/001", "Term": 25}
  ]
}
```

1. `Bids` and `Criteria` are separate horizontal-header collections; their names are template-defined, not hardcoded.
2. The initial model uses the **Cartesian product** of horizontal collections in input order: Alpha/Price, Alpha/Term, Beta/Price, Beta/Term.
3. `$path` is **only a logical-row grouping key**. It does not prescribe ordering and is not a sort key.
4. All intersection records for a logical row must arrive together in the correct horizontal order. This contiguity requirement must be validated or otherwise enforced without sorting.
5. The first record of each `$path` group supplies ordinary row fields. Other records supply their respective intersection fields.
6. For a nonempty horizontal product, expect a **complete rectangular matrix**: one record per leaf-level tuple per logical row, including placeholders for missing values. Parent-only columns do not add intersection records. When the product is empty, use the degenerate normalization in section 4.3.
7. There is no fixed `Value` field. Templates may use arbitrary fields such as `Price`, `Amount`, or `Term`.
8. No `$h1` / `$h2` identifiers are required in input. Positional mapping is the baseline.
9. Sparse matrices, independent horizontal IDs, per-bid criteria lists, and cross-collection lookups are deferred.
10. Validate contiguous groups, consistent integer levels and expected group cardinality without sorting. Positional ordering is the producer's responsibility: without IDs the engine cannot detect a permutation of otherwise complete records. Reject a path that reappears after its group closes. Paths are opaque keys, not ancestry syntax; use the numeric level stack for hierarchy.

### 4.3. Empty horizontal collections

1. Empty collections at any horizontal level are valid. If the full product L is zero, require exactly one descriptor record per logical vertical row, grouped by `$path`; this record supplies ordinary row fields and hierarchy. Do not require the vertical collection to be empty and do not manufacture intersection records.
2. Normalized record arity is L when L is positive and one when L is zero. Preserve input ordering, nesting and multirow vertical rendering in both modes. An empty vertical collection still renders no vertical repetitions.
3. A nonempty parent collection with an empty child still repeats its parent-only columns and parent headings. Empty child blocks emit zero columns. An empty outer collection emits no descendant columns or orphan child headings. Zero-width surviving columns remain physical columns with width zero.
4. Ordinary row fields outside removed blocks continue to render from the descriptor. Leaf cells/formulas disappear with their empty block. A surviving formula with an empty or ambiguous referent follows the safe error policy; the engine does not replace every empty aggregate with zero.

## 5. Repetition semantics

1. Repeat vertical template blocks by logical input item, preserving data order and instance multiplicity. The data tree is authoritative; templates project selected levels while transparently traversing unrendered ancestors. Equal names/field values do not merge separate logical instances. Existing contiguous `$path` matrix grouping remains the logical-item contract.
2. Recursively repeat authored horizontal blocks by their own collection cardinalities, preserving source order, parent-only columns and authored formatting.
3. Repeat horizontal headers using their respective source collections and authored template rows.
4. A multirow vertical block remains one logical item, regardless of how many physical Excel rows it generates.
5. Support vertical headers and footers and their Excel formulas across the horizontally expanded columns.
6. Static report rows/columns inside PrintArea remain present and correctly positioned. Content outside PrintArea is removed under section 3.1.
7. Independent tables are marked vertical row runs under section 3.3. The one worksheet horizontal tree applies consistently through the sheet's printed row bands. Multiple tables sharing it use its horizontal collections where they intersect its leaf blocks; tables outside that strip remain vertically independent. Worksheet widths/heights are global: reject conflicting ownership or independently moving content rather than copy/delete a neighboring report accidentally.
8. Formulas and separator rows may be outside any vertical region; their reference scope is global rather than the preceding group. A separator never belongs to both adjacent regions. Displacement of later tables is calculated from source coordinates together, not by repeatedly reinterpreting a mutated worksheet.

## 6. Excel formulas and calculation

1. Preserve formulas in normal cells, headers, footers, and grand-total areas, including formulas placed **above** the repeated data.
2. Support ordinary Excel functions without restricting the formula language to an allowlist; required examples include `SUM`, `SUBTOTAL`, and `COUNT`.
3. Maintain correct relative, absolute, and mixed references when rows/columns are repeated.
4. Ensure dynamic ranges refer to the intended generated rows/columns, including nested group subtotals and grand totals. Do not replace `SUBTOTAL` with `SUM` or otherwise rewrite the user's chosen function semantically.
5. The engine does not need to evaluate formulas; the output must request recalculation by a compatible spreadsheet application.
6. Use parsed reference transformation against final workbook geometry, including external-margin removal. The recommended first implementation is an internal bounded A1 reference-token parser/transformer using existing dependencies. It must distinguish strings, numbers, identifiers, function calls and sheet qualification; regex-only replacement is insufficient. No formula evaluator or general Excel expression grammar is required.
7. If a formula cannot be safely preserved/transformed, fail with actionable diagnostics rather than silently returning a mathematically incorrect workbook.
8. The supported reference forms (cross-sheet references, named ranges, structured references, etc.) must be determined by a capability investigation, not presumed supported without tests.
9. Initially support explicit A1 cells/rectangles, relative/absolute/mixed flags, same-workbook sheet references and safely resolved simple names. Owner-approved F-M312-1 clarification (2026-10-23): finite range endpoints may descend independently in rows and columns. Preserve lexical endpoint order and each endpoint's flags through copying and relocation; normalize bounds only for geometry and retain range syntax when endpoints coincide. `$` controls copying, not structural relocation. Bind group-header/footer ranges to their own descendants, leaf references to their horizontal instance and static totals to their intended complete region. Reject ambiguous repeated absolute anchors, disjoint or empty images instead of selecting an arbitrary instance or taking a bounding rectangle.
10. Reject affected constructed references (`INDIRECT`, unsafe `OFFSET`), unresolved formula-valued names, structured/external/3-D/whole-axis references, spill/array/data-table formula structures and other unproven forms with a specific remedy. Arbitrary ordinary function names over supported operands remain permitted; the boundary is reference safety, not a function allowlist. Known untouched structures can be preserved only when their dependencies remain safe through expansion and crop.
11. Normalize shared scalar formula followers from their authored master before copy/relocation; verify a real package fixture during implementation. Remove stale formula caches/calculation chains as appropriate and request recalculation. Never present cached template results as newly calculated values.

## 7. Fidelity and compatibility

1. Preserve visible report appearance: fonts, fills, borders, alignment, number formats, cell content, row heights, column widths, hidden states, merges, and print layout.
2. Preserve relevant workbook/worksheet settings when they do not conflict with expansion.
3. Non-print-oriented artifacts (comments, review metadata, etc.) are not primary requirements: discard those wholly outside PrintArea, preserve safe retained ones, otherwise reject unsupported templates clearly rather than silently corrupting output. An artifact straddling a crop/repeat boundary is not safe merely because its XML is valid.
4. Investigate interactions with images, charts, conditional formatting, named ranges, tables, validation, and other Excel constructs; establish supported/unsupported cases based on evidence.
5. Do not invent presentation choices or automatically merge cells.
6. Support/preserve/reject decisions must account for all moved cells and references, including static sheets that reference expanded sheets. Initial support preserves safe structures and rejects affected tables, charts, drawings, conditional formatting or validation without a verified relocation handler. Schema validation must be accompanied by semantic range/relationship assertions.

## 8. Empty inputs and error handling

1. An empty repeat collection produces zero repetitions; the repeated region is absent in output. Static content remains unless structurally dependent on the removed region.
2. No conditional template constructs (`If`/`Else`) are required for this feature.
3. Prefer producing the best **correct** result possible. Non-fatal anomalies should produce warnings even on success.
4. Structural ambiguities, unsupported operations, or unsafe formula transformations that could invalidate the result are errors.
5. Diagnostics should include severity, stable code, description, and where available worksheet and cell/range context.
6. Exact error taxonomy and recovery behavior should be refined through concrete test cases rather than invented wholesale upfront.
7. Preserve no-partial-output behavior across Core, CLI and HTTP. Use deterministic, severity/code/location-bearing warnings on success and established bounded error aggregation on failure. CLI keeps binary output separate from diagnostics and does not overwrite output on generation failure. HTTP keeps the existing document routes and DOCX binary contract; its additive structured XLSX success-warning representation must be specified in the Implementation Plan, since a warning-count header alone is insufficient.

## 9. Architecture constraints

1. Separate parsing, template semantics, data resolution, expansion planning, formula handling, XLSX rendering, and diagnostics at least conceptually.
2. Design the vertical engine with horizontal expansion in mind; avoid an implementation that must be replaced to add the second axis.
3. Build an explicit final workbook plan before mutation, with recursive horizontal/vertical instance paths, parent-owned ranges, source-to-output identities, later-table displacement and final cropping. Formula and merge transformations consume the same plan.
4. Reuse existing project conventions and common data/formatting components when appropriate.
5. Keep the report-generation engine deterministic. AI-assisted data preparation is a possible future feature, not part of generation runtime.
6. Extend existing public façades; keep DOCX/XLSX package and rendering internals separate. Supply dependencies externally and retain per-operation mutable state. Preserve existing DOCX construction and DI behavior; no new dependency defaults or internally constructed service dependencies.

## 10. Acceptance scenarios

1. Static XLSX with field replacement preserves styles and print settings.
2. One-level vertical repetition.
3. Arbitrary nonconsecutive nested `$level` values.
4. Multirow repeated vertical blocks.
5. Parent header, children, parent footer, repeated across multiple parents.
6. Group `SUM`/`COUNT`/`SUBTOTAL` and a grand total, including totals located above the data.
7. One horizontal level and an authored multicolumn repeated block.
8. Mandatory nested D:E/E fixture with zero-width parent-only D; two companies and zero, one, three and seven criteria; one company heading per parent, existing merge boundaries adjusted only through owned children, single-cell collapse handled without invented merges.
9. Two vertical and two horizontal levels with four positional intersections per logical row.
10. Multiple tables and worksheets, with at most one horizontally expanded region on each worksheet.
11. Empty collections and missing template/data levels per section 4 and 8.
12. Formula references remain correct after both-axis expansion.
13. Successful generation may return warnings; unsafe expansion fails with actionable diagnostics.
14. Excel opens the generated workbook without repair prompts and recalculates formulas as expected.
15. Multiple independent tables separated by a blank vertical-control row, including identical marker names on both sides, an author-defined zero-height separator, nested levels, formulas and a shared horizontal tree.
16. Physically absent external control/content rows/columns; correct report-relative formulas, PrintArea, supported references and static content after cropping.
17. Empty inner/outer horizontal collections retain ordinary vertical rows from one descriptor each; surviving unsafe formulas fail explicitly. Native numeric/Boolean values, date-formatted serials and both workbook date systems retain their intended types.
18. Existing Core, CLI and HTTP generation/analysis interfaces accept XLSX without separate format-specific public APIs/routes; correct MIME/filename/diagnostics and unchanged DOCX compatibility. Automated transport and OpenAPI/CLI regression checks belong to implementation.
19. An SDK-only research runner validates supported formula transformations independently of ClosedXML; the implementation tests its lexical boundaries, unsupported references, shared scalar formulas, deterministic diagnostics and no partial output.

Each accepted scenario should have an XLSX template, JSON fixture where applicable, and structural/value/formula/style assertions. Focused unit/integration tests suffice. Manual Microsoft Excel verification is acceptable for final visual appearance, recalculation, repair prompts and compatibility; automated Excel execution is not required.

## 11. Initial limitations and deferred extensions

1. At most one horizontal expansion region per worksheet.
2. Full positional Cartesian-product matrix when nonempty; one descriptor per logical row when the product is empty. No sparse intersection lookup.
3. No arbitrary horizontal ID matching or cross-collection lookup expressions.
4. No conditional blocks in XLSX templates.
5. No formula evaluation in the engine.
6. No AI dependency at runtime.
7. No automatic layout invention or merging.
8. No automatic ISO-to-native-date inference in the initial delivery. Preserve native dates/serial bindings and both date systems.
9. No separate XLSX public APIs/routes, new UI, general Excel evaluator or speculative OOXML repair framework.

## 12. Items requiring architecture evidence before implementation

1. Revised isolated evidence in [Architecture Assessment](architecture-assessment.md) establishes nested geometry, zero/one/many criteria, bounded parent-merge projection, physical crop, two tables/sheets, both empty-axis modes and selected formulas/date types. It is supplied-plan research, not a production implementation.
2. Internal reference-transformer results match the existing parser on 61 prior and 110 new supplied-plan formula transformations. Broader parsing acceptance does not establish semantic support; the SDK-only reference envelope is explicit in section 6.
3. Implementation must still verify marker parsing/inference, shared formula package normalization, scoped named/cross-sheet references through crop, public-interface compatibility, native date edge cases, advanced-feature rejection and realistic resource limits.
4. Manual Excel appearance and final compatibility checks are unexecuted and remain acceptance work. No large-workbook throughput estimate is inferred from the tiny fixtures.

These are **research/verification tasks**, not authorization to change the agreed data contract or invent new template syntax without review.

## 13. Product Owner decisions and planning handoff

1. On 2026-10-20 the Product Owner approved nested `$h(level)` horizontal regions, parent-only/zero-width columns, marker/PrintArea-only ownership, physical external-region removal, valid empty horizontal input with surviving vertical descriptors, native types, parsed reference transformation, focused automated checks and manual Excel acceptance.
2. The subsequent explicit merge clarification approves bounded adjustment of existing parent-owned merges and collapse to an ordinary cell. This supersedes the earlier fixed-leaf/repeated-parent-label recommendation and any conflicting interactive choice.
3. The Product Owner explicitly approved blank-control-row boundaries, including zero-height separators retained once inside PrintArea, and Core + CLI + HTTP delivery.
4. The final interface clarification requires extending existing MonoReport interfaces with DOCX backward compatibility, not separate XLSX APIs. Internal format dispatch and additive compatible metadata are implementation mechanisms to specify in the Developer-owned plan.
5. The internal SDK-only reference transformer is the revised architecture recommendation under the dependency investigation, not an already implemented production capability. Detailed API success-warning transport, DI/construction compatibility and resource budgets must be concrete and reviewable in the Implementation Plan.
6. The research/documentation task does not approve that plan, start implementation, authorize production dependency changes or authorize Git operations. Present the revised specification and plan for the required review before production work.
7. **2026-10-22 approved architecture amendment:** the Owner explicitly adopts data hierarchy as authoritative and source-template structure as presentation. Sections 3.2.7, 3.3.5, 4.1.3–4 and 5.1 now record that decision, replacing the earlier unmatched-input warning requirement and any requirement to match source/data parent chains. Full instance ancestry, ordering/multiplicity, quiet omitted levels, explicit root and actual data-parent lookup are approved; no synthetic levels or generalized IR is authorized. This is an explicit specification amendment based on the Owner decision, not an implied change from an implementation plan.
8. **2026-10-22 approved diagnostic clarification:** section 3.2.8 distinguishes invalid ancestor traversal from missing properties. Default warning/empty output and strict failure through existing `MissingValues.Error` are approved for unavailable vertical ancestors; existing public missing-property behavior remains. The broader necessity/usage of MissingValues across engines is a mandatory retrospective question, not permission to remove or redesign it now. Optional data schemas and code-configurable diagnostic severity remain future backlog work only.
9. **2026-10-22 implementation authority:** the Owner approves M3.5-before-M3.4 as an implementation choice and authorizes M3.5 hierarchy/identity/binding/scopes and focused verification only, delegated to the Feature Developer. Record these approved changes in plan/progress, stop after M3.5 for Owner review, and preserve all six deferred M3.4 nested rendered integration requirements. No M3.4 projection/rendering or nested preflight guard removal, Git mutation, AGENTS.md change or retrospective document creation is authorized. M6 retains the mandatory internal architecture document and reviews of MissingValues and product/architecture priorities in general rules.

10. **2026-10-22 Owner acceptance and M3.4-only authorization:** the Owner explicitly accepts M3.5 based on its completed implementation, verification evidence and independent Reviewer PASS in [reviews/m3.5-codex/review.md](reviews/m3.5-codex/review.md). The Owner authorizes M3.4 implementation through the established Architect → Developer → independent Feature Reviewer workflow, preserving the approved data hierarchy/full structural identities and addressing deferred source-compatible merge-candidate filtering/owner deduplication. All six nested rendered-integration conditions in implementation-plan.md §5.3 remain mandatory and must use actual generated XLSX output. Record changes, verification and limits in existing progress documents, then stop for Owner review of M3.4; review findings require Owner authorization before fixes. This authorization does not activate public nested generation, authorize M3.6 or later work, unrelated refactoring, Git staging/index changes or commits. The existing public nested preflight guard remains active; public activation retains its separate authorization boundary.

11. **2026-10-22 explicit M3.4 Owner acceptance:** the Owner accepts M3.4 based on the completed implementation, all six nested rendered-integration proofs, [independent Reviewer PASS](reviews/m3.4-codex/review.md) and recorded verification evidence. This task records acceptance only, directly without subagents. M3.6 and later milestones remain unauthorized, and the public nested preflight guard must remain active. The failing RealHost test is deferred to a separate task; do not investigate or modify it here. No Git staging/index change or commit is authorized. Stop after recording acceptance; this is not final Feature acceptance.


### 13.12. Owner-approved bounded public nested activation — 2026-10-24

1. Activate accepted M1–M3 nested behavior through existing Core, CLI and HTTP generation paths. On a worksheet containing a parsed nested vertical region, allow no horizontal expansion or a chain of at most two authored horizontal blocks, with one physical header row per level. Reject deeper horizontal trees, multiple child blocks and multirow horizontal headers before binding, including empty/null input. This temporary delivery boundary does not restrict existing flat worksheets or source analysis and does not cap collection cardinalities or vertical depth.
2. Preserve all existing source/data, formula, ownership, geometry, formatting, package, diagnostics and no-partial-output guarantees. No new public API, setting, dependency or unsupported capability is approved.
3. The Owner manually opened all three retained M3.14 generated workbooks in Microsoft Excel without repair warnings or visible layout problems; Excel calculated formulas automatically on first opening. These are Owner-reported observations, not verified numeric results, save/reopen or print acceptance. Excel version was not supplied. Full E1 remains pending.
4. Implement the approved plan directly without subagents or Git index changes. Run focused, affected, Release build and normal solution checks with representative SDK save/reopen. Stop after implementation for independent review and Owner assessment; do not implement M4–M6 or infer final Feature acceptance.
5. **Owner acceptance — 2026-10-24:** the Owner formally accepts the bounded public activation following independent review PASS WITH CONDITIONS, which found no implementation defects, 1,048 passing solution tests with zero failures/skips, a Release build with zero warnings/errors, and Owner confirmation that staging changes were made manually and commit `c70e524` contains only Feature 012 changes. This resolves the integrity condition by Owner confirmation, not by matching the recorded fingerprints. Microsoft Excel E1 checks, M4–M6, and final Feature acceptance remain pending.
