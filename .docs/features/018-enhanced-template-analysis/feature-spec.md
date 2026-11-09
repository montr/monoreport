# Feature Spec — Enhanced Template Analysis & Sample JSON Generation

**Status:** Accepted (2026-11-09). The Product Owner approved Feature acceptance with the known large-Dataset Chrome crash tracked separately; the crash is unresolved and is not accepted as correct behavior.

**Feature directory:** `.docs/features/018-enhanced-template-analysis/`

## 1. Intent

MonoReport authors need useful, understandable sample data for their DOCX and XLSX templates without knowing the engine-specific JSON conventions. The current Playground constructs Sample JSON from Analyzer requirements in the frontend. This approach lacks enough information to produce appropriate XLSX data for vertical levels and horizontal intersections, and sometimes produces scalar values unsuitable for formatters.

The goal is to make Sample JSON generation a reusable MonoReport library capability, improve template analysis with nonbinding semantic type hints, and let users inspect and copy a generated sample before replacing their Dataset.

This Feature must preserve existing template syntax, generation semantics, and DOCX/XLSX compatibility contracts.

## 2. Scope

1. Generate Sample JSON from a supplied DOCX or XLSX template through a reusable .NET library service, with format-specific behavior owned by the corresponding engine.
2. Expose sample generation through a separate HTTP operation, `POST /api/documents/sample`.
3. Improve representative sample structure and values for supported DOCX/XLSX scenarios.
4. Add optional, advisory semantic type hints to template-analysis occurrences without changing existing requirement types or generation restrictions.
5. Update the Playground Sample JSON flow to preview, copy, and explicitly accept generated JSON.
6. Add copying of analysis results from the Playground Analyze tab.

## 3. Sample Generation Requirements

1. The operation accepts a template, an optional positive integer `itemsPerCollection` setting (default: 10), and an optional value mode (`filled` by default; `empty` as an alternative).
2. Reuse existing data-independent template inspection before constructing a sample. Inspection errors, including unavailable formatters and incompatible requirement shapes, prevent a replacement candidate; return the available actionable diagnostics. Preserve existing compatibility restrictions, all supported DOCX story coverage, and Raw/opaque-content exclusions. Do not introduce a second parser or common validation engine.
3. The library's Sample Generator owns example values. `filled` produces varied, deterministic values across generated fields/items, including generic nonempty strings for unformatted fields without stronger evidence. Use conservative numeric values, supported ISO date strings, and Boolean values where supported by semantic or structural evidence. Do not infer meaning from field names or require every conditional branch to be exercised.
4. Repeated requests produce reproducible samples for identical templates and settings under unchanged formatter registrations/metadata and generator version. Use stable internal generation rules; no public seed setting, wall-clock-dependent values, or unseeded randomness is required. Repeated references to the same generated property reuse its value.
5. `empty` creates the same applicable structure, using empty strings for ordinary scalar placeholders and `false` for Boolean dependencies. Structural metadata necessary for valid XLSX data, including `$level` and required `$path` grouping keys, must still be populated.
6. For DOCX, respect loop scopes, nested collections, conditional dependencies, parent/root references, and primitive current-element expressions, without imposing stricter collection rules. Infer primitive examples where current-element usage provides evidence; otherwise leave unknown-shape arrays empty and report the specific unresolved sample limitation where relevant.
7. For XLSX, honor actual authored `$level` values and parent-before-child record order, including templates whose footer layout appears first. Do not invent unrendered data ancestors or infer hierarchy from `$path` text.
8. For XLSX horizontal repetition, generate header collections and contiguous `$path` record groups with the positional multiplicity required by the engine, including combined repetition. Preserve first-record behavior for ordinary/parent-only fields and corresponding tuple values for leaf fields. Handle supported empty-product descriptors when a collection remains empty under the rules above; no zero-count setting is introduced.
9. Apply `itemsPerCollection` to logical elements: DOCX items per collection instance, XLSX horizontal items per distinct collection, and XLSX hierarchy roots/children per represented level. Nested DOCX collections receive the count within each generated parent. Header/body/footer layouts do not independently multiply logical items. Physical XLSX intersection records expand as needed and count toward the total allocation limit.
10. Construct each shared logical collection once and reuse its data across repeated references, regions, and worksheets where required. Report incompatible collection demands rather than silently adjusting counts. When a meaningful sample cannot be inferred, report the specific limitation rather than inventing data hierarchy or binding semantics.
11. Resolve semantic evidence only for uses sharing an actual generated value in the engine's scope. Identical hints agree; unspecified uses do not override a known hint. With conflicting known hints, use the first known hint in source order and emit a targeted sample warning. Existing Boolean structural dependencies take precedence for sample values; conflicting formatter hints produce a warning. Do not merge different XLSX levels/records solely because their flattened canonical paths match.
12. Unknown formatter semantics use the deterministic generic fallback; missing optional metadata alone does not require a warning. Coarse hints do not establish formatter-specific ranges or constraints. Report identified conflicts or limitations without introducing constraint solving or generation restrictions.
13. Preserve the aggregate limit of 1,000,000 generated array elements, counting nested and expanded records before allocation. Count overflow or exceeding the limit fails with an actionable diagnostic and no candidate. Never return an arbitrarily truncated sample or silently reduce the requested count to fit.
14. Return syntactically valid, complete JSON for every useful candidate, including candidates with unresolved sample limitations. Such limitations produce targeted warnings with affected template locations where available. A candidate is not guaranteed to render successfully without editing.
15. If no useful candidate can be produced, return actionable failure diagnostics and no replacement candidate. Keep template/sample failures distinct from malformed requests and unexpected processing failures.
16. Do not invoke document generation automatically for every sample request. Verify representative samples through focused end-to-end tests instead.

## 4. HTTP API

1. Provide a separate `POST /api/documents/sample` multipart endpoint accepting the template and sample settings; do not add sample generation as a parameter of Analyze.
2. On useful sample construction, return HTTP 200 with `data` containing the generated JSON value and `warnings` containing ordered diagnostics, including an empty array when none exist. Candidates with warnings remain successful results. No `generated`/`partial`/`unavailable` status taxonomy is required.
3. Report inability to produce a useful sample as a failed operation with structured, actionable diagnostics, following existing API error conventions. Exact error-response shape and status code should align with established MonoReport API patterns during architecture design.
4. Invalid request settings, unsupported formats, excessive upload sizes, and unexpected failures must use the existing corresponding API error handling. Finalize the value-mode form field name, duplicate/blank setting handling, and allocation-limit failure mapping during architecture design; preserve existing Analyze/Generate request validation behavior.
5. Do not expose engine-specific planning geometry or require the frontend to submit Analyzer internals.

## 5. Template Analysis Enhancements

1. Preserve existing requirement categories (`object`, `array`, `value`, `boolean`); specifically, do not reinterpret `value` as `string`.
2. Add an optional `semanticType` hint to individual field occurrences, including current-element fields. Initial values are exactly `number`, `date`, and `boolean`; omit the property when unspecified. Preserve existing occurrence fields and source order.
3. Extend existing `FormatterMetadata` with an optional nullable, init-only `SemanticType` property, using a small Core-owned enum for Number, Date, and Boolean. Keep `IFormatter`, the existing metadata positional constructor, and deconstruction unchanged. Existing custom formatters continue working without supplying hints.
4. Formatter-related hints come exclusively from the selected formatter implementation's metadata, including replacements and aliases. Analyzer must not infer them from registration names. Number/date/Spellout implementations declare their own applicable hints; casing formatters leave the hint unspecified.
5. Hints are fixed metadata, independent of invocation arguments. A formatter whose input meaning varies by arguments may omit the hint. Do not introduce argument-dependent hint callbacks, `AcceptedTypes`, formatter-provided `SampleValue`, or a general type/schema system.
6. Semantic hints describe advisory input meaning, not formatter output, accepted JSON kinds, or guaranteed conversion success. Numeric strings remain accepted where supported; date examples remain JSON strings. Hints do not constrain data or introduce generation errors.
7. Preserve per-occurrence formatter hints even when uses of one dependency disagree; do not aggregate them into requirement types or add Analyze errors for semantic disagreements. Ordinary scalar fields without a formatter hint may receive `boolean` from an existing Boolean structural dependency only where it applies to the same engine scope; do not propagate it between distinct XLSX records solely by canonical path. Resolve sample choices under Section 3, item 11.
8. Preserve existing validation of formatter availability and invocation syntax, including inactive branches and empty loops. Unregistered formatters in inspected expressions continue to produce Analyze diagnostics; existing opaque/Raw exclusions remain.
9. Retain each occurrence's expression, including formatter syntax. No new structured formatter argument API is required merely for sample generation.
10. Make the new hints available for user inspection in Playground Analyze where present.

## 6. Playground UX

1. Extend the existing Generate Dataset dialog; do not introduce a second full JSON editor.
2. Default the value mode to **Filled** instead of Random, while retaining **Empty** as an option. Show **Filled** before **Empty** in the selection list. The sample count defaults to 10.
3. The first action generates a candidate. Show formatted, scrollable, selectable, read-only JSON in the existing Monaco editor in a review state, together with warnings or other relevant diagnostics visibly above the preview. Explain that the sample can be reviewed and copied here; editing and Generate are available after explicitly replacing the Dataset.
4. Provide **Copy** without changing the current Dataset.
5. Provide **Replace Dataset** as an explicit acceptance action. Until acceptance, the existing Monaco Dataset remains unchanged; cancellation preserves it.
6. When the current Dataset has unsaved changes, make the replacement consequence clear at the acceptance step without a second routine confirmation; do not ask for discard confirmation before sample generation. Keep the warning visible in the review state. Permanently describe below the collection-count field that nested collections multiply the total generated elements and that the overall limit is 1,000,000; do not use a tooltip.
7. Allow the user to adjust settings and regenerate. Changing settings invalidates the previous candidate until regeneration. Do not allow a stale candidate to replace data if its template, Dataset context, or settings have changed. Cancel/close and superseding requests must prevent late responses from restoring or applying obsolete candidates.
8. A useful candidate with warnings can still be copied or accepted. No replacement action is available if sample generation fails.
9. Reuse the existing Dataset update behavior so document previews/results are invalidated appropriately after replacement.
10. Add **Copy** to the Analysis tab toolbar to copy the complete stored analysis JSON, including errors, warnings, occurrences, and semantic hints, without reconstructing visible rows or altering the result. Keep it available whenever an analysis result exists, including when it has no errors; disable copying when no analysis exists.
11. Clarify in the Analyze UI that requirement types describe JSON structure while occurrence semantic types are advisory formatter hints; do not change the analysis contract to align the labels.
12. Keep Sample JSON diagnostics visible in the dialog and associated Problems view while that candidate/request is current. Invalidate sample diagnostics when the request is canceled, settings/context change, or the candidate is accepted. Diagnostics from a subsequent document Generate remain visible under the existing Generate behavior. Diagnostic contents are session state; restoring the Problems pane preference after reload does not restore prior messages.

## 7. Out of Scope

1. Data-aware Analyze or a separate dataset compatibility validator.
2. Extraction of a common validation engine solely for this Feature.
3. JSON Schema import, export, generation, or schema-based validation.
4. Automatic document generation after every sample request.
5. General-purpose template/data schema inference, new restrictions on accepted data, or changes to supported DOCX/XLSX features.
6. New caching infrastructure, compiled-template APIs, or frontend implementations of engine-specific sample rules.
7. Unrelated improvements to Generate diagnostics; assess separately.

## 8. Acceptance Criteria

1. Existing template-only Analyze, document Generate, and their compatibility behavior remain unchanged except for additive advisory occurrence metadata.
2. A DOCX template with nested loops, primitive-element use, and formatter-supported values produces a structurally appropriate sample; relevant cases successfully render through the existing engine.
3. XLSX templates with vertical hierarchy, horizontal repetition, and combined repetition produce correct `$level` and `$path` structure and expected record counts in representative supported cases.
4. Samples honor `filled`/`empty`, varied reproducible values including unformatted fields, count semantics, and aggregate allocation limits. Verify exact-limit success and pre-allocation rejection of excess/overflow without truncation.
5. Invalid templates produce no candidate. Unknown collection shapes, conflicting hints, and incompatible shared collection demands follow Section 3 with targeted diagnostics; absent optional formatter metadata alone does not require warnings. Verify deterministic fallback and Boolean precedence without new generation restrictions.
6. The API returns useful candidate data and warnings on success, and actionable diagnostics on failure.
7. The Playground shows the generated JSON before replacement, supports copy and cancel without data loss, and applies replacement only after explicit acceptance.
8. Pending requests or stale review results cannot overwrite a changed template or Dataset; cover settings changes, cancellation/closure, and superseded requests.
9. The Analyze tab can copy its analysis results, and any available semantic hints are visible without misleading users about strict JSON types.
10. Existing DOCX/XLSX regression behavior and appropriate API/UI tests remain passing. Focus coverage on formatter replacements/aliases and unchanged custom implementations, omitted/occurrence-specific hints, primitive/unknown collection shapes, footer-first layouts, shared collections, and supported empty products. Representative samples render through existing Generate; no universal render-success guarantee or broad performance suite is required.

## 9. Architecture Review Topics

1. Confirm the smallest library/API contract and placement of DOCX/XLSX sample logic, without exposing internal planners publicly.
2. Confirm the additive metadata implementation preserves existing constructor/deconstruction and custom formatter compatibility, with no argument-aware hint or formatter sample-value infrastructure.
3. Define sample-generation failure diagnostics and status-code details consistently with existing API conventions.
4. Review representative XLSX corner cases, collection reuse, ambiguous ancestors, and resource limits.
5. Verify that UI stale-result protection aligns with existing Dataset editing and generation behavior.
6. Identify necessary updates to user-facing documentation, but do not edit those documents without separate Product Owner approval.

**Approval:** The Product Owner approved this Feature Spec, the [architecture](architecture-assessment.md), and the [Implementation Plan](implementation-plan.md) on 2026-11-08, then accepted the completed Feature and all six stages on 2026-11-09. See the [acceptance record](implementation-review.md#product-owner-acceptance--2026-11-09) for acceptance-criteria evidence and verification boundaries.

## 10. Acceptance record — 2026-11-09

1. **Feature: Accepted (2026-11-09).** Acceptance includes the recorded Product Owner-approved corrections and final UI refinements. The later approved empty Analysis Copy behavior supersedes the initial disabled-without-analysis wording in Section 6, item 10: Copy remains enabled and safely copies `{}` when no analysis exists; an existing result copies its complete stored wire payload.
2. Chrome can crash during Generate with a very large Dataset, particularly after HTTP 413. Removing the redundant client-side JSON parse did not resolve the observed crash. The root cause remains unconfirmed; acceptance does not mark this defect fixed or accept it as correct behavior.
3. The remaining crash is tracked outside Feature 018 in [the separate follow-up task](../../tasks/large-dataset-chrome-crash.md). The possibility that an incomplete or erroneous XLSX response reaches the preview pipeline is an investigation hypothesis, not an established cause. No crash correction or broader performance guarantee is included in this acceptance.
4. Remaining user-facing documentation work is recorded in the [final acceptance review](implementation-review.md#remaining-user-facing-documentation). No user-facing documentation or compatibility contract is changed by this acceptance task.
