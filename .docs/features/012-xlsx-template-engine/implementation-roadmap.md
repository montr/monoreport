# Feature 012 — Implementation Roadmap

**Status:** Revised 2026-10-20 following Product Owner decisions and follow-up research. This is an architecture roadmap for a Developer-owned Implementation Plan, not an approved implementation plan or authorization to implement.
**Planning principle:** Use recursive two-axis source/instance geometry and final cropping from the start; deliver a working combined slice early, including correct formulas and authored merge projection. Extend existing MonoReport interfaces with DOCX backward compatibility.

## Phase 0 — Research consolidation and implementation planning

1. Use the revised [Feature Spec](feature-spec.md), [Architecture Assessment](architecture-assessment.md) and [isolated research evidence](research-spikes/README.md) as the baseline. Record approved nested blocks, parent-only/zero-width columns, bounded existing-merge adjustment, blank-control-row boundaries, empty-product descriptors, physical crop and Core + CLI + HTTP scope.
2. Select the existing Open XML SDK and a bounded internal reference transformer. No ClosedXML runtime dependency is recommended; retain its isolated comparator only for optional differential research. Define the reference/rejection envelope and do not build an evaluator or complete Excel grammar.
3. Prepare a concrete Implementation Plan, including existing façade signatures, constructor/DI compatibility, package-format dispatch, CLI/HTTP routing through current interfaces, format metadata and additive HTTP structured success-warning access. Preserve current DOCX contracts; no separate XLSX APIs/routes.
4. Include focused shared-formula, crop-linked reference, deep nesting and unsupported-feature cases as implementation verification. Reuse the valid research evidence, but do not mistake supplied row/column bindings for a generic parser or engine.
5. Present the revised planning baseline and Implementation Plan for the required approval. The Product Owner selects roles and gates; no independent review workflow or implementation authorization is inferred from research.

**Exit gate:** reviewed specification/architecture baseline and approved Developer-owned Implementation Plan, with concrete delivery/compatibility and warning contracts. Research demonstrates feasibility but does not itself satisfy this gate.

## Phase 1 — Existing-interface foundation and complete plan model

1. Extend the existing MergeEngine/TemplateAnalyzer composition and DI to select DOCX/XLSX from the actual package. Preserve existing public DOCX signatures/construction behavior and operation/stream ownership. Keep format-specific internals separate without a plug-in framework.
2. Inventory workbook parts, one finite rectangular PrintArea per retained worksheet, external controls, native values, string/style tables and affected coordinate-bearing features. Do not infer named report regions.
3. Parse scalar syntax with shared MonoReport grammar and neutral data-scope boundaries. Parse top `$h(level)` geometry and left `$header(level).Collection`/`$level(level).Collection` controls separately; coalesce equivalent long/short forms.
4. Model recursive horizontal nodes, own columns, children and instance paths; vertical marked runs and group/body/footer ownership; source cells, range images and final crop. No fixed leaf-tile or hardcoded two-level geometry.
5. Return deterministic code/severity/location-bearing diagnostics, including warnings on success and no partial result on error. Add bounded unsupported-feature/size checks now.
6. Render static XLSX fields with native numbers/Booleans and existing explicit text formatters; preserve authored native dates and workbook date-system flags. Use a separate cell adapter for shared/rich strings as needed.

**Exit gate:** static XLSX and retained DOCX paths behave through existing Core entry points; recursive two-axis/crop types exist and native/diagnostic contracts have focused checks. Any temporary external-margin hiding is explicitly incomplete.

## Phase 2 — Early working nested two-axis slice with formulas

1. Implement one vertical logical block together with parent D:E, child E and the corresponding nested horizontal collections. Parent-only D may have width zero; repeat counts come from data, not template width.
2. Render one parent heading per company and criteria within it. Project only existing owned merge boundaries through children; cover zero, one, three and seven criteria and ordinary-cell collapse. Preserve anchor content, authored styling and physical zero-width columns.
3. Normalize a positive product as one record per leaf tuple and a zero product as one descriptor per logical row. Keep vertical rows and safe static content under empty outer/inner collections.
4. Apply the bounded internal reference parser/transformer to source-plan bindings; include formulas above data, in details and below. Preserve functions and copy flags, reject unsafe images, and request recalculation without evaluating formulas.
5. Physically remove external control/content rows/columns in the planned final geometry. Transform supported formulas, merges, PrintArea and other coordinates through that same crop; reject retained references to disposable content.
6. Prove the blank-control-row region boundary with two independent tables. Retain the author-defined zero-height separator once inside PrintArea; no extra boundary token or named range.

**Exit gate:** combined XLSX fixture passes automated values, native types, coordinates, merge, width/height, crop and formula checks. External recalculation/manual Excel can validate calculated results; automatic Excel execution is not required. Do not defer the first two-axis/formula interaction until a late phase.

## Phase 3 — Nested vertical groups and scoped formula correctness

1. Add arbitrary nonconsecutive integer levels, multirow blocks, parent headers/deeper bodies/footers, unmatched-level warnings and otherwise safe descendant rendering. Reject ambiguous grammar rather than guess.
2. Support independent marked runs with identical markers on both sides of a separator, multiple tables/worksheets and displacement of later printed content. Reject incompatible side-by-side row ownership/shared heights.
3. Bind group subtotals to their own descendants, global totals to complete intended regions, and leaf operands to their tuple. Distinguish copying from structural relocation; classify unique/rectangular/disjoint/empty images.
4. Support SUM/COUNT/SUBTOTAL and arbitrary ordinary functions over safe explicit references without function substitution or a function allowlist. Add actionable rejections for disappearing/disjoint/ambiguous targets and constructed references.
5. Normalize real shared scalar formula masters/followers; verify same-workbook cross-sheet and simple-name bindings through expansion and crop. Preserve no-partial-output and remove stale caches/calculation chains appropriately.

**Exit gate:** nested two-axis groups, top/bottom totals, blank boundaries, safe references and known failure cases pass structural and calculated-value assertions. Neither schema validation alone nor parser acceptance is sufficient.

## Phase 4 — Deeper horizontal hierarchy and bounded fidelity

1. Extend the same recursive model to deeper levels and multiple contained child blocks, preserving own-column order and data associations. Do not replace the vertical planner or duplicate the renderer into unrelated axis engines.
2. Add a rendered three-level fixture; the research currently verifies the three-level width recurrence only. Enforce justified dimension/resource limits rather than an unexplained depth-two limit.
3. Complete multirow headers, owned merges, one-horizontal-root enforcement and zero-product cases across more vertical regions.
4. Implement only agreed reference/metadata handlers and preservation rules. Preserve safe retained structures, discard external disposable artifacts, and reject affected unsupported tables/charts/drawings/CF/validation/comments with a remedy.
5. Verify native serial/date formatting in both workbook systems; keep ISO-to-native inference deferred. Test relevant native-date edge cases without inventing a typed JSON date contract.

**Exit gate:** full two-vertical/two-horizontal and deeper-model acceptance cases pass, with an explicit supported/preserved/rejected feature matrix and no invented formatting/merges.

## Phase 5 — CLI and HTTP through existing interfaces

1. Extend existing CLI generate/analyze commands to XLSX; preserve current switches, DOCX behavior, binary stdout, exit/error conventions and output-on-failure protection. Put warnings in the diagnostic stream.
2. Extend existing HTTP document generate/analyze endpoints to dispatch the package format and return the correct XLSX MIME type/filename. Do not add XLSX-specific route families or silently change DOCX success/error schemas.
3. Implement the plan-approved additive structured success-warning access; warning count alone does not satisfy that requirement. Keep diagnostic data safe and bounded rather than place unbounded arrays in headers.
4. Update OpenAPI/contract and focused Core/CLI/API integration tests for XLSX, DOCX compatibility, diagnostics, format mismatch, stream ownership and no partial files/responses. Do not assume file extension reliably identifies the package.
5. Preserve compatible public result construction/deconstruction and existing DOCX DI/composition paths while adding required format metadata internally or additively. Avoid service dependency defaults or internal dependency construction.

**Exit gate:** both formats work through the same supported public interfaces; DOCX compatibility and XLSX warning/error/output contracts have evidence. No new UI or Playground behavior is included.

## Phase 6 — Capacity, documentation and final acceptance

1. Measure representative approved report sizes, logical groups, physical columns including zero-width columns, populated cells, formula/reference counts and package complexity. Use checked arithmetic and compact intervals before materialization; optimize based on actual evidence.
2. Complete meaningful negative tests and documentation for incomplete/empty matrices, invalid ownership/markers, unsafe merge/reference/crop cases, unsupported structures, data conversion and resource limits. Avoid a combinatorial OOXML test framework.
3. Verify the resolved dependency closure/license notices. The current recommendation introduces no ClosedXML runtime package; revisiting advanced formula grammar requires separate evidence and review.
4. Run appropriate solution/transport regression checks after implementation changes. Reuse already valid evidence and disclose unexecuted checks.
5. Perform manual Microsoft Excel opening, no-repair-prompt, recalculation, zero-width merged-heading appearance, date display and print-layout acceptance. Automated Microsoft Excel execution is not a prerequisite. Record any accepted verification limitation explicitly.

**Exit gate:** approved acceptance criteria pass with known limits documented and Product Owner acceptance recorded. Research alone does not establish final Excel compatibility.

## Cross-phase requirements

1. Maintain a current nested XLSX + JSON golden fixture from Phase 0. Historical bare-h/fixed-leaf/MR-named fixtures remain comparison evidence only and are not current authoring examples.
2. Use one immutable final geometry/ownership plan for both axes, formulas, existing merges, later-table displacement and crop. Do not repeatedly rescan/reparse a mutated workbook.
3. Keep both template analysis and generation on the same data-independent preflight. Data-dependent matrix/layout/reference checks run before mutation; state and diagnostics remain per operation.
4. Preserve valid supported content, native values and authored formatting; reject unsafe transformations rather than return plausible but incorrect workbooks.
5. Keep investigation separate from implementation. No production code/dependency changes, Git mutation or browser automation are authorized by this research task.
6. The earlier Core-only effort range is not a current estimate. Estimate the reviewed implementation plan including existing-interface/constructor compatibility, CLI/API integration and structured warnings; small research fixtures do not establish production performance.
