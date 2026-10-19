# MonoReport — Codebase Simplification Review (October 2026)

**Date:** 2026-10-19
**Scope:** Whole repository after Feature 011 acceptance (commit `bdb6e00`): Core, CLI, API, UI/Playground, tests, and supporting infrastructure.
**Nature:** Analysis only. No application code, tests, samples, or feature documentation was modified. No Git-mutating command was executed.

## 1. Method and verification basis

1. Read `AGENTS.md`, `docs/docx-template-compatibility.md`, and the Feature 011 artifacts (Feature Spec, Architecture Assessment, Implementation Plan including sections 9.1–9.11).
2. Read every production source file of `MonoReport` (3,071 lines total), `MonoReport.Api`, and `MonoReport.Cli`, and all test suites (4,386 lines), including the maintained fixture layout under `MonoReport.Tests/Fixtures/`.
3. Inspected transport and bundling infrastructure: `MonoReport.Api/Program.cs` middleware pipeline, OpenAPI transformers, `MonoReport.UI/webpack.config.cjs` with `PlaygroundExamplesPlugin`, `MonoReport.UI/src/services`, `.github/workflows/quality.yml`, `Dockerfile`, and `compose.yaml`.
4. Traced the end-to-end flow: `MergeEngine.Merge` → `TemplateAnalyzer.Inspect` (story/package reading, planning, Raw extents, identity preflight, requirement building) → `DocumentRenderer`/`MergeContext` rendering and cleanup; and the analyze flow used by CLI and HTTP.
5. Verified dead-code and vocabulary claims by repository-wide symbol search (grep), not by compiler warnings alone.
6. Executed `dotnet test MonoReport.slnx -c Release` once during this review: 348 passed (Core 311, CLI 13, API 24), 0 failed, 0 skipped, matching the acceptance evidence in Implementation Plan section 9.5–9.9. No frontend checks, container smoke test, or browser automation was run; none is required for this analysis.

## 2. Executive summary and overall assessment

1. **The codebase is already lean.** Feature 011 completed the scheduled removals: comment/bookmark cloning, anchor remapping, and identity-frame machinery are gone; the production core is eight small services totaling roughly 3,100 lines, and the CLI, API, and Playground layers contain no independent parser or compatibility logic. The strongest design property is that Analyzer and Generate share one plan (`StoryPlan`) and one compatibility preflight, enforced by diagnostic-parity regressions rather than by duplicated code.
2. **Remaining complexity is concentrated in one place.** `MonoReport/Services/DocxTemplatePlanner.cs` (543 lines) plus its `Raw` (133) and `Identities` (130) partials carry the essential DOCX/OOXML complexity: opaque-boundary handling, field-state tracking, container pairing, merge preflight, and identity validation. Most of this is essential rather than accidental, and the `Identities` partial was recently rewritten under independent review (Implementation Plan sections 9.8–9.10); restructuring it now would add churn for little benefit.
3. **The genuine simplification opportunities are modest and mostly small.** The findings below are: two structurally similar marker-pairing implementations, two parallel token-stream traversals with duplicated control-stack state, a string-encoded canonical path that the requirement builder re-parses, a near-duplicated code-to-message map between Core and API, a small amount of Feature 011 leftover vocabulary, a stale OpenAPI diagnostic list, and one explicitly deferred architecture item (`MergeContext` hosting DOCX allocators). No broad rewrite is recommended anywhere.
4. **Everything recommended preserves externally observable behavior** except three items that touch the public Core surface (removing an unused public method, unifying public warning diagnostics, reshaping the public `MergeContext`) and one that changes published API problem text; these four require explicit Product Owner decisions and are collected in section 6.

## 3. Findings (ordered by priority)

### F1 — Duplicated marker-container pairing between Raw handling and block building

1. **Current complexity and evidence.** Two implementations decide whether an opening and closing marker form a permitted sibling layout. `DocxTemplatePlanner.CreateBlock` (`MonoReport/Services/DocxTemplatePlanner.cs:185–221`) validates same-parent paragraphs, otherwise dedicated sibling rows under one row parent, with `MarkerOnlyRow`/`MarkerOnlyContainer` checks and crossing/misorder rejection. `DocxTemplatePlanner.Raw.RawContainers` (`MonoReport/Services/DocxTemplatePlanner.Raw.cs:43–57`) re-implements the same paragraph-vs-row decision, adding the independent-root check. A third near-copy of the independent-root name list appears in `DocxTemplatePlanner.InspectIdentities` (`DocxTemplatePlanner.cs:76–80`), duplicating `IndependentRoot` (`DocxTemplatePlanner.Raw.cs:37–41`). Both pairing paths also share a second hidden rule (marker-only container suitability) that is easy to drift. History supports the concern: the Feature 011 corrections in Implementation Plan sections 9.8 and 9.9 were both subtle placement/traversal defects in this area.
2. **Proposed simplification.** Extract one shared helper on the planner, for example `TryPairContainers(XElement start, XElement end, out XElement first, out XElement last)`, implementing the paragraph-or-row decision with the marker-container checks; `CreateBlock` supplies the extra Else handling, ordering, and Raw-exemption checks, and `RawContainers` adds the independent-root condition. Replace the inlined independent-root name list in `InspectIdentities` with the existing `IndependentRoot` helper (or a shared name list). No marker placement rules change.
3. **Expected benefit.** One authoritative placement decision used by both consumers; removes the highest-drift-risk duplication in the pipeline; slightly shrinks both files.
4. **Estimated effort:** M. 5. **Regression risk:** Medium (it touches placement rules that gate many rejections).
6. **Relevant tests and suggested verification.** `Feature011RawTests` (all placement/precedence and cross-cell/cross-table/cross-root theories), `UnifiedTemplateTests` paragraph/row block cases, `Feature011CompatibilityTests` invalid-placement cases, `Feature010IdentityTests.Scenarios`/`ControlsAndComments`, and CLI/API parity suites. Run the Core project with the same focused filters used in Implementation Plan section 7.2–7.3; then the full solution once.

### F2 — Two parallel traversals of the same active token stream with duplicated stack state

1. **Current complexity and evidence.** After `ParseOutsideRaw` has stripped Raw bodies, the planner walks the same `ParagraphPlan` token stream twice, each with its own stack: `BuildBlocks` (`DocxTemplatePlanner.cs:119–183`) maintains a marker stack to build `BlockPlan` trees, and `BuildDependencies` (`DocxTemplatePlanner.cs:252–307`) maintains both a `loops` string list and a `controls` stack to canonicalize field/for/if paths. Both encode the same nesting semantics; the `controls` stack in `BuildDependencies` exists almost solely to keep `loops` consistent on `EndFor`.
2. **Proposed simplification.** Merge the two walks into one structural pass that maintains a single marker/context stack and emits both the `BlockPlan` tree and the `PlannedDependency` list. `BuildDependencies` keeps its current public output ordering because the merged pass visits tokens in identical document order. The two expressions of block nesting then live in one place.
3. **Expected benefit.** Removes one full traversal and the duplicated stack bookkeeping; makes future grammar changes (e.g., a new control kind) a single-place edit.
4. **Estimated effort:** M. 5. **Regression risk:** Medium (dependency and occurrence ordering is part of the public contract and is asserted in tests).
6. **Relevant tests and suggested verification.** `Feature011RawTests.Raw_payload_is_not_tokenized_and_active_dependencies_resume_in_source_context` (asserts exact occurrence and requirement order), `UnifiedTemplateTests` requirement/context cases, `TemplateLanguageParserTests`, `Feature010AcceptanceTests`, `HttpApiTests.Analyze_returns_ordered_requirements_and_all_occurrences`, and the CLI/API parity tests. Full solution run at the end.

### F3 — Canonical loop paths encoded as strings and re-parsed by the requirement builder

1. **Current complexity and evidence.** `TryCanonical` (`DocxTemplatePlanner.cs:309–331`) produces paths as strings, pushing `arrayPath + "[]"` onto `loops`; `RequirementBuilder.Add` then re-splits those strings (`TemplateAnalyzer.cs:74–98`, with `Components` at `TemplateAnalyzer.cs:112–117` decoding the `"[]"` suffix) to reconstruct array-vs-object structure; `DataResolver.DataPath` (`MonoReport/Services/DataResolver.cs:12–24`) again rebuilds strings for diagnostics.
2. **Proposed simplification.** Introduce a small structured canonical path (ordered segments with an array marker per segment) produced by the planner and consumed by `RequirementBuilder`, keeping the existing public string fields on `TemplateRequirement`/`TemplateOccurrence` unchanged (public schemas remain exactly as approved).
3. **Expected benefit.** Removes the string-format round trip and the `[]`-suffix decoding; makes the requirement tree's semantics explicit in types rather than in string conventions.
4. **Estimated effort:** S. 5. **Regression risk:** Low–Medium (internal only; public shapes unchanged).
6. **Relevant tests and suggested verification.** Same as F2: all requirement- and occurrence-order tests, plus `ValidationSchemaTransformer`-covered OpenAPI schema contract tests.

### F4 — Near-duplicated code-to-message maps between Core and the API

1. **Current complexity and evidence.** `TemplateException.SafeMessage` (`MonoReport/Models/TemplateException.cs:79–92`) and `ApiProblemFactory.SafeMessage` (`MonoReport.Api/Services/ApiProblemFactory.cs:31–55`) maintain two almost identical mappings for the `template.*` and `generation.*` codes, with slight wording drift (for example, "The DOCX document structure is unsupported." versus "The DOCX structure is unsupported."). Every new diagnostic code must be added in two places.
2. **Proposed simplification.** Move the shared `template.*`/`generation.*` code-to-message mapping into Core and have `ApiProblemFactory` use it, keeping the request-level codes (`request.*`, `data.*`, `culture.*`) local to the API. This slightly changes API problem `detail` prose for a few codes; no test asserts that prose today (verified against `HttpApiTests` and `ApiHostTests`), but the wording is published API behavior, so it needs Product Owner approval (section 6, Q2).
3. **Expected benefit.** Single source for user-facing diagnostic phrasing across CLI and API; removes a taxonomy of drift risk.
4. **Estimated effort:** S. 5. **Regression risk:** Low (code/field/part/location fields unchanged; only `detail` wording equalizes).
6. **Relevant tests and suggested verification.** `HttpApiTests`, `CliTests` (stderr output), `ValidationContractParityTests` (JSON analysis body unaffected), `OpenApiContractTests`.

### F5 — Dead and Feature 011-orphaned vocabulary left behind after simplification

1. **Current complexity and evidence.** Repository-wide search shows:
   1. `DocxStoryReader.Read(WordprocessingDocument)` (`MonoReport/Services/DocxStoryReader.cs:32–37`) has no callers in any project; only internal `Inspect` is used. It is a public method, so removing it is a public-surface change (section 6, Q1).
   2. `WordXml` (`MonoReport/Services/WordXml.cs`) contains constants with zero non-definition references: `SpaceAttributeName`, `DirtyAttributeName`, `FieldSeparatorValue`, `FalseValue`, `PreserveSpaceValue`, `DefaultHeaderFooterTypeValue`, `TypeAttributeName`, `FooterElementName`, `HeaderReferenceElementName`, `FooterReferenceElementName`.
   3. Additional constants are referenced only by test helpers, not by production code: `RunFontsElementName`, `BoldElementName`, `ItalicElementName`, `ColorElementName`, `FontSizeElementName`, `HighlightElementName`, `UnderlineElementName`, `LanguageElementName`, `AsciiFontAttributeName`, `HighAnsiFontAttributeName`, `WidthAttributeName`, `ValueAttributeName`, `InstructionTextElementName`, `HeaderElementName`, `SingleUnderlineValue`, `TableGridElementName`, `GridColumnElementName` (production uses raw `"trPr"`/`"gridBefore"` strings in `DocxTemplatePlanner.Raw.cs:114` instead). These are legacy of the pre-011 formatting-preservation and identity code.
   4. `MergeContext` redefines `WordprocessingNamespace` locally (`MonoReport/Models/MergeContext.cs:105–106`) and defines `DrawingPropertiesName` privately, duplicating vocabulary that `WordXml` (and the tests) already carry.
2. **Proposed simplification.** Delete the zero-reference constants; optionally move the test-only styling names into the `Docx` test helper class so production vocabulary stays production-facing; point `MergeContext` at `WordXml.WordprocessingNamespace`. Removing `DocxStoryReader.Read` requires the Product Owner decision in section 6, Q1; the internal-only vocabulary pruning is ordinary maintenance.
3. **Expected benefit.** Smaller, truthful vocabulary; removes the impression that formatting rewrites still exist.
4. **Estimated effort:** S. 5. **Regression risk:** Low.
6. **Relevant tests and suggested verification.** Compilation of the solution plus the full Core test project once (test helpers touch some constants).

### F6 — `MergeContext` blends per-operation data state with DOCX allocation, clone mapping, and part writeback

1. **Current complexity and evidence.** `MonoReport/Models/MergeContext.cs` is a public class in the Models project that, in addition to the shared data stack, warnings, and error recording, performs: the physical-copy `Clone` with XElement mapping and drawing/permission ID allocation (`MergeContext.cs:55–83`), package-wide drawing/permission reservations (`MergeContext.cs:42–52`), and the comments-part delete writeback (`MergeContext.cs:87–95`). It therefore contains DOCX-specific machinery in the shared Models layer, which sits awkwardly with Coding Guideline 4 (isolate DOCX-specific logic from shared APIs). Feature 011's Implementation Plan section 3.9 explicitly deferred "moving the remaining DOCX allocator out of `MergeContext`".
2. **Proposed simplification.** Keep a slim, format-neutral operation context (data stacks, warnings, errors, missing-value behavior) in Models, and move the DOCX parts (clone mapping, drawing/permission allocation, comments writeback) into an internal Services-side per-operation companion object constructed per `MergeEngine.Merge` call. `MergeContext` is public, so the reshape is a breaking change to the Core public surface and needs the Product Owner decision in section 6, Q1; it is implementation-proportional only if combined with the F1–F3 planner work in a single maintenance pass.
3. **Expected benefit.** Clearer boundary between shared language/data state and DOCX mechanics; restores the Models layer to format-neutral contracts.
4. **Estimated effort:** L. 5. **Regression risk:** Medium (the whole render path touches this state).
6. **Relevant tests and suggested verification.** `Feature010IdentityTests.ControlsAndComments.Reused_engine_isolates_Raw_cleanup_drawing_and_permission_allocations` (sequential and concurrent), `Feature011CompatibilityTests` Raw cleanup tests, `Feature010DocxTests` drawing/permission cases, then the full solution.

### F7 — Stale diagnostic codes in the published OpenAPI 422 description

1. **Current complexity and evidence.** `ApiOperationTransformer.TransformAsync` (`MonoReport.Api/Services/ApiOperationTransformer.cs:59`) lists the 422 codes as `..., template.invalid_syntax, ..., template.invalid_format, ...`. Neither code exists; the actual codes produced by Core are `template.invalid_package`, `template.invalid_xml`, `template.unsupported_structure`, `template.invalid_expression`, `template.unmatched_brace`, `template.invalid_context`, `template.requirement_conflict`, `template.invalid_structure`, `template.tracked_revision`, `generation.invalid_data`, and `generation.invalid_format` (see `TemplateException.cs:79–92`). This is a documentation-accuracy defect introduced when the code families evolved, not new behavior.
2. **Proposed simplification.** Correct the description string to the actual produced codes. Because this is a published response description, record it as a small maintenance correction.
3. **Expected benefit.** Accurate API documentation; prevents clients from coding against nonexistent failure codes.
4. **Estimated effort:** S. 5. **Regression risk:** Low.
6. **Relevant tests and suggested verification.** `OpenApiContractTests` (extend to assert the 422 description lists exactly the produced codes, preventing recurrence).

### F8 — Source-order bit-packing idiom is opaque and repeated

1. **Current complexity and evidence.** The `(long)index << 32` / `sourceOrder & uint.MaxValue` packing of element index plus intra-paragraph offset appears independently in `DocxTemplatePlanner.Plan` (`:40`), `InspectIdentities` (`:85`), `RevisionDiagnostics` (`:499`), `ValidateRawDeletion` (`DocxTemplatePlanner.Raw.cs:83`), `ParagraphPlan.PositionAt` (`MonoReport/Models/TemplatePlan.cs:16–21`), and `SourceContext` (`DocxTemplatePlanner.cs:426–443`).
2. **Proposed simplification.** Introduce two or three named static helpers (pack, element index, in-paragraph offset) in the planner, and reference them everywhere the idiom appears. No ordering behavior changes.
3. **Expected benefit.** Readability and fewer pack/unpack mistakes when the idiom is next extended (Raw handling showed how sensitive these orders are).
4. **Estimated effort:** S. 5. **Regression risk:** Low.
6. **Relevant tests and suggested verification.** Full Core test project (ordering assertions exist across all feature suites).

### F9 — `MarkerOnlyRow` is a pure alias, and `BuildBlocks` end-kind matching is needlessly indirect

1. **Current complexity and evidence.** `MarkerOnlyRow` (`DocxTemplatePlanner.cs:394–395`) simply forwards to `MarkerOnlyContainer`; and `BuildBlocks` computes the expected closing kind through an inline nested switch inside the "does not match" condition (`DocxTemplatePlanner.cs:160–166`), which is the single densest control-flow expression in the core.
2. **Proposed simplification.** Inline the alias (or rename the shared helper to reflect both roles), and hoist the kind-to-closer mapping into a small static expression or local function with a short explanatory comment. Purely mechanical.
3. **Expected benefit.** Marginal readability; included only because both edits occur naturally inside F1/F2 and should not be separate passes.
4. **Estimated effort:** S. 5. **Regression risk:** Low.
6. **Relevant tests and suggested verification.** Same as F1/F2.

## 4. Recommended implementation sequence

1. **Quick wins (independent, low risk):** F7 (OpenAPI code list), F5 (vocabulary pruning except `DocxStoryReader.Read`), and the `MergeContext` namespace deduplication inside F5. Each is a small maintenance change verified by compile plus the full solution suite once.
2. **Planner consolidation pass (recommended as grouped maintenance):** F1 + F9 together, then F2, then F3 and F8. F1 must land before F2 so the pairing decision is already shared when the traversal merges. F3 and F8 are small and can ride along. This is the highest-value chunk; run the focused Core filters after each step and the full solution suite once at the end.
3. **Cross-layer cleanup (after Product Owner decision in section 6, Q2):** F4, which equalizes a small amount of published API problem wording.
4. **Public-surface decisions first, then implementation:** `DocxStoryReader.Read` removal and F6 (`MergeContext` reshape) require explicit Product Owner approval (section 6, Q1 and Q3). If approved, F6 should be its own maintenance pass with the operation-isolation and concurrency tests as the primary regression net.
5. **Not recommended now:** any restructuring of `DocxTemplatePlanner.Identities.ValidateIdentities` (including the `Do/while` definition-retention loop and the `Removed` predicate), the renderer's containment scans, or the test-class names. See section 5.

## 5. Areas reviewed where no changes are recommended

1. **`TemplateLanguageParser`** (`MonoReport/Services/TemplateLanguageParser.cs`): a small, stateless, single-file grammar; the Raw extent preprocessing in the planner correctly sits outside it per Feature Spec 9.1. Fine as designed.
2. **`ValueFormatter` and `RubSpellOutFormatter`**: self-contained, deterministic, diagnostically guarded; the exception-to-`TemplateException` conversion is narrow. Leave unchanged.
3. **`DataResolver`**: minimal and correct; its path-string rebuild remains justified while `DataPath` strings are public-adjacent diagnostics (F3 removes the remaining re-parse).
4. **The transport layers** (`MultipartRequestReader`, `DocumentEndpoints`, `ApiProblemFactory` structure, `CliApplication`, `CliOptionsParser`): clean, linear, with contract-parity tests (`ValidationContractParityTests`, `OpenApiContractTests`) and no language reimplementation. No structural change beyond F4/F7.
5. **The render loop** (`DocumentRenderer`): the repeated `Contains` containment scans are intentional (source-map-protected rendering with parse-once copies); the D6 shared cell normalization already satisfies the Feature Spec's no-duplication requirement for one correction path.
6. **`DocxTemplatePlanner.Identities.ValidateIdentities`**: dense but recently rewritten, independently reviewed (Implementation Plan sections 9.8–9.10), and its complexity maps directly to the approved D2/D3 contract. Restructuring now would trade verified code for churn at real regression risk.
7. **`MergeEngine`, `TemplateAnalyzer`**: thin orchestrators with the right responsibilities; `RequirementBuilder` itself is fine once F3's structured path lands.
8. **Test infrastructure**: the shared `Docx` helpers, the `Reject`/`Prepare` parity helpers, and focused fixture layout under `MonoReport.Tests/Fixtures/` are proportionate. The `Feature010*` test-class names are retained deliberately per Implementation Plan section 3.6 ("renaming is optional; avoid cosmetic churn") — keep them. Hand-built composition in `Docx.CreateEngine` is a deliberate isolation choice over `AddMonoReport`; keep it.
9. **Frontend/Playground** (`MonoReport.UI`): no template parser duplication (bundling via `PlaygroundExamplesPlugin` reads `Samples` directly); the documented `docx-preview` workaround in `docxPreviewNormalization.ts` has an explicit upstream removal condition. Leave unchanged.
10. **Infrastructure** (`Dockerfile`, `compose.yaml`, `quality.yml`, `scripts/container-smoke.sh`, solution layout with `Directory.Build.props`, project-linked fixtures): coherent; the test-project fixture `None`-includes mirror each other intentionally so each test project is self-sufficient.

## 6. Open questions requiring Product Owner decisions

1. **Q1 — Public Core API surface.** Removing the unused public `DocxStoryReader.Read(...)` and any reshaping of the public `MergeContext` (F6) are breaking changes to the Core public surface. Confirm whether the Core surface is frozen between releases (and whether a deprecation cycle is required) or whether unused public members may be removed directly.
2. **Q2 — Published API problem wording.** F4 changes `detail` wording for a few codes in HTTP problem responses. Approve the wording consolidation, or require byte-stable API response text.
3. **Q3 — Deferred deferrals.** Implementation Plan section 3.9 deliberately excluded (a) the allocator's location in `MergeContext` and (b) tree-based dependency traversal. This review recommends (a) now as F6 and (b) as the F2/F3 combination. Confirm both are now in scope as Maintenance, or record the continued deferral in `.docs/features/backlog.md` (note: rule 5 requires Product Owner approval for backlog additions).
4. **Q4 — OpenAPI prose corrections.** F7 alters a published response description. Confirm it can proceed as part of normal Maintenance.

## 7. Review limitations

1. Coverage of application code is complete; the OpenAPI transformer bodies, Dockerfile, workflows, and frontend sources were inspected for structure rather than line-by-line behavior.
2. Binary DOCX samples (`Samples/*.template.docx`) were not unpacked or re-validated here; their correctness is covered by the executed acceptance tests, and their content influences no recommendation above.
3. The full .NET suite was executed once; frontend checks, container smoke tests, and Word/manual rendering checks were not run (none was needed for analysis-only conclusions, and none is proposed as a verification obligation by these findings).
4. Symbol-usage claims are based on repository-wide text search across all five C# projects; they are disclosed as grep-based evidence rather than compiler-verifiable dead-code analysis (the repository intentionally carries no coverage/dead-code tooling per AGENTS.md Test Quality rule 2).
