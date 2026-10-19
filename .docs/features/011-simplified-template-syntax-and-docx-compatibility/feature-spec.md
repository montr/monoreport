# Feature 011 — Simplified Template Syntax and DOCX Compatibility

**Status:** Approved on 2026-10-19; implementation verification in progress  
**Feature ID:** 011  
**Proposed directory:** `.docs/features/011-simplified-template-syntax-and-docx-compatibility/`

## 1. Goal

Simplify MonoReport's template language and DOCX implementation to focus on reliable generation of printable documents. Remove unsupported OOXML complexity rather than maintaining elaborate transformations for rare Word editing constructs. Introduce concise closing markers and a literal, unprocessed region.

## 2. Product principles and authority

1. MonoReport is a document generation engine, not a general-purpose OOXML editor.
2. `docs/docx-template-compatibility.md` is the target compatibility policy for Feature 011. This specification incorporates the Product Owner's approved decisions D1–D5; both documents describe the intended Feature 011 contract, not all current engine behavior or retroactive changes to accepted Feature 010 criteria.
3. For Word constructs, distinguish **support and transform**, **preserve unchanged**, and **reject with an actionable diagnostic**. Do not reject unrelated opaque OOXML merely because MonoReport does not process it.
4. If a newly discovered scenario would materially expand supported behavior or require significant implementation complexity, present options and a recommendation to the Product Owner before implementation. Technical feasibility alone does not justify support.
5. The approved Feature Spec governs implementation and review. Unapproved compatibility extensions are not defects merely because they are technically possible.

## 3. Scope

1. Replace `{EndIf}` with `{/If}` and `{EndFor}` with `{/For}` throughout the engine, Analyzer, documentation, tests, examples, and DOCX template assets.
2. Add `{Raw}` ... `{/Raw}` to preserve the enclosed document content without template-language processing.
3. Apply the target DOCX compatibility policy, simplifying or removing unnecessary handling of comments, bookmarks, internal hyperlinks, and related OOXML identity/remapping behavior where the policy does not require it.
4. Ensure consistent template validation and behavior across Core, CLI, HTTP API, and Playground.
5. Update user-facing documentation and repository examples to describe only the current syntax once the Feature is implemented.

## 4. Template language

### 4.1 Closing markers

1. Conditional blocks use `{If:Path}` ... `{/If}`, optionally with `{Else}`.
2. Repetition blocks use `{For:Path}` ... `{/For}`.
3. Existing expression paths, scalar formatting, escaping, data semantics, and nesting remain unchanged except where explicitly modified by this Feature.
4. `{EndIf}` and `{EndFor}` cease to be control markers. Under the unchanged identifier grammar they are ordinary scalar expressions referencing properties named EndIf and EndFor; they do not close blocks. Apply ordinary parsing, dependency extraction, data resolution, and diagnostics, with no aliases, migration handlers, dedicated legacy warnings, or compatibility code.
5. Controls remain case-sensitive. Exact `{Raw}` is reserved as an opening control; a property named Raw remains accessible through ordinary qualified paths such as `{./Raw}`. Existing brace-escape precedence applies outside Raw.
6. Update maintained templates and executable examples, including binary DOCX fixtures and Playground samples. Preserve historical Feature 010 documents and evidence unchanged. Where current tests load historical binaries, migrate their executable fixture sources or references rather than modifying archived evidence or retaining legacy runtime syntax support.

### 4.2 Raw blocks

1. `{Raw}` ... `{/Raw}` defines a literal region. All enclosed template-looking content is inert, including scalar expressions, malformed braces or expressions, `{If}`, `{Else}`, `{For}`, opening/closing controls, formatters, and brace escapes. Do not validate, evaluate, unescape, extract dependencies from, or emit actionable occurrences for that content.
2. Preserve the region's original Word content and formatting, subject to compatibility validation and transformations imposed by enclosing active blocks, such as cloning a containing loop body or excluding an unselected conditional branch. Raw itself removes only its delimiter containers and safe delimiter-attached annotations under section 5.2; it does not clone or reconstruct its body.
3. Raw may occur inside `If` and `For`, including nested combinations. It neither changes the data context nor creates a requirement, occurrence kind, or warning for literal template-looking content. Expressions following Raw resume the enclosing context and normal source ordering.
4. Opening and closing markers must occupy dedicated sibling paragraphs under the exact same parent, or dedicated sibling table rows under the same table. A row marker must be wholly in one paragraph in one cell, with all other visible row content empty. Apply existing marker-container rules, including whitespace and split-run support. Reject inline active delimiters, mixed-content marker containers, cross-parent and paragraph-to-row layouts, and crossing parts or independent story roots. Do not promote markers to a common ancestor.
5. Support multiple paragraphs within one table cell and mixed paragraph/table sibling content. Complete tables, including merged cells, images, and opaque Word content may be preserved inside paragraph-delimited Raw. Row-delimited Raw preserves its body rows and complete nested tables.
6. Inside Raw, only the first structurally valid, exact, unescaped `{/Raw}` in an eligible closing marker container ends it. Inline, mixed-content, wrong-parent, escaped, or opaque lookalikes remain literal. A closing lookalike inside a table cannot close a paragraph region outside that table. If no valid closer exists, report an unclosed Raw diagnostic; do not diagnose the literal lookalikes as template errors. Match delimiters by eligible container placement; a subsequent integrity/compatibility failure is an error, not a reason to skip that closer and consume a larger region.
7. Raw does not nest: an additional `{Raw}` inside it is literal. Other control-looking text, including `{Else}`, `{/If}`, and `{/For}`, cannot alter the enclosing block stack. Actual enclosing branch/end controls must occur outside Raw, and the resulting Raw node must lie wholly within one enclosing body or branch. An exact eligible standalone Raw closer cannot itself be literal at that level; no additional Raw escape syntax is introduced.
8. Unclosed Raw, active stray `{/Raw}` outside Raw, and invalid active delimiter placement produce shared Analyzer/Generate diagnostics. Opaque Word fields, hyperlinks, drawings, and content controls remain hard recognition boundaries; their contained marker text is not an active delimiter. Preserve existing field-state tracking through Raw.
9. DOCX structural-integrity and compatibility validation apply independently of literal template content. Apply section 5 restrictions even in inactive branches and zero-iteration loops. Raw is not an exemption from revision, range, relationship, or enclosing-operation checks.
10. Unlike `For`/`If` row transformations, Raw does not reject merged body rows merely for containing vertical merges. Reject only layouts where removing its delimiter rows would compromise vertical merges or required DOCX structure. Reject delimiter deletion that would remove a section boundary or leave a cell without required block content; provide an authoring remedy instead of merge repair, section-property relocation, or synthesized replacement content.

## 5. DOCX compatibility simplification

The compatibility policy and the following requirements distinguish `For`/`If` body transformation from Raw delimiter removal. An operation's footprint includes its opening, optional Else, closing containers, both bodies, and their descendants; restrictions are independent of runtime data.

1. **Comments and bookmarks around `For`/`If` (D3):** preserve unaffected constructs and complete ranges fully enclosing an operation when endpoints and all required attachments remain intact outside its footprint and the resulting range stays valid. Do not clone, rename, relocate, or reconstruct them. Reject local attachments inside a body or marker container, split-branch ranges, and crossing/incomplete arrangements that cannot be preserved safely. Pair identities within independent story roots, including separate footnotes/endnotes. Enclosing-range support is conditional on a simple implementation: if substantial extra logic proves necessary, recommend actionable rejection and obtain Product Owner confirmation before changing this contract.
2. **Comments and bookmarks in Raw (D2):** allow unchanged constructs within or spanning Raw when identities, ranges, attachments, and targets remain valid. Remove a complete bookmark or comment attachment, including a reference-only point comment, wholly attached to one removed Raw delimiter container together with that container, without an error. Reject deletion that splits an attachment/range or otherwise leaves an invalid surviving structure. Handle associated comment references and definitions safely: retain definitions needed by surviving attachments, and remove a deleted annotation's definition only when no surviving attachment or required reference depends on it. Preserve unrelated definitions, relationships, and package content; do not add general package cleanup. Reject retained internal links to any target removed or invalidated by this deletion. If safe deletion requires significant additional machinery, report the product tradeoff before proceeding; do not silently broaden support or substitute blanket rejection. Enclosing `For`/`If` restrictions take precedence, including annotations on Raw delimiters inside those operations.
3. **Internal hyperlinks:** preserve explicit internal links only when the link is outside `For`/`If` footprints and its target remains valid and unchanged, including a safely preserved enclosing bookmark. Raw body links are allowed under item 2. Reject links located inside `For`/`If` even when their target is stable, and links to removed, repeated, ambiguous, or otherwise invalidated targets. Apply source/reference checks across already inspected package XML where necessary; do not rely on runtime cardinality. Preserve unrelated pre-existing unresolved anchors outside affected operations as today. Remove dynamic target remapping; do not add broken-link repair or scan opaque field instructions for references.
4. **External hyperlinks:** preserve external relationships, addresses, and fragments/anchors unchanged, including relationship-based hyperlinks carrying `w:anchor`. Do not reinterpret an external fragment as a bookmark target.
5. **Fields, TOC, and content controls:** preserve Word field instructions and cached results without evaluating or recalculating `REF`, `PAGEREF`, TOC, or field-based hyperlinks. Preserve opaque content-control subtrees, properties, bindings, and relationships; do not process contained expressions or join outside expression fragments across their boundaries. Opacity does not suppress independent compatibility checks within the existing inspection footprint.
6. **Tracked revisions (D5):** reject recognized unresolved revisions in the currently inspected main document, headers, footers, footnotes, and endnotes. Retain existing recognized content, cell, move/custom-XML range, deleted-payload, property-change, and Office 2010 conflict families and existing drawing/picture exclusions. Validate independently of language evaluation, including Raw, inactive branches, and content controls within that footprint. Do not expand detection to auxiliary parts or package-wide inspection. The Track Changes setting alone is not unresolved revision markup and must not cause rejection. Diagnostics direct authors to accept or reject changes in Word.
7. **Permission ranges:** retain accepted Feature 010 behavior: complete source-root-local ranges may be preserved, cloned within one loop body/conditional branch with matching local IDs, fully enclose operations, or disappear with a self-contained deleted marker. Reject affected split, crossing, duplicate, missing, or misordered endpoints. Keep permission matching and per-copy ID allocation; removal of bookmark/comment support does not authorize removal of permission support. Raw alone makes no copy and must not split surviving permission endpoints.
8. **Drawing identities and other supported content:** retain package-wide drawing-ID reservations, including unchanged comments and glossary parts, and allocate unique drawing IDs for physical copies, including images inside repeated Raw. Preserve relationships, image bytes, formatting, supported stories, tables, and list numbering definitions/references without forced restarts. Existing `For`/`If` vertical-merge limitations remain unchanged; Raw uses section 4.2.10.
9. **Simplification boundary:** remove comment-definition cloning, bookmark cloning/renaming, and dynamic internal-anchor remapping used solely for rejected cases. Retain source identity/range validation, permission copying, drawing safety, and the narrow comment-definition update needed for approved Raw deletion. No general annotation relocation, identity cloning, link remapping, or OOXML repair may be added to implement Raw.
10. Do not corrupt output or silently discard unrelated Word content. Analyzer and generation must apply the same policy; failures must produce useful locations and remedies, with no partial DOCX returned.

## 6. Validation and user experience

1. The Analyzer identifies unsupported DOCX structures before generation when detectable without data, and reports actionable errors with Word-oriented remediation.
2. Generation enforces the same compatibility restrictions and returns no partial DOCX on failure.
3. Raw regions do not create JSON requirements for expressions contained solely inside them.
4. CLI, API, and Playground expose the same language and compatibility behavior as Core, without independent parser variants.
5. The README links to `docs/docx-template-compatibility.md`. Once Feature 011 is accepted, remove transitional wording that describes the policy as future-only.

## 7. Acceptance criteria

1. `If`/`Else` and `For` use the new closing markers in paragraph and table-row blocks with retained nesting, path, scalar, formatter, and data semantics. Former closing spellings behave as ordinary scalar paths and never close blocks; maintained executable assets and current documentation do not depend on their former control meaning. Historical Feature 010 documents and evidence remain unchanged.
2. Raw preserves malformed and valid expression-looking content, formatters, doubled braces, and nested-looking controls without language diagnostics, evaluation, unescaping, requirements, occurrences, or content-derived warnings. Opening and closing containers are removed. Verify ordinary processing and dependency/context ordering resume afterward, including a path used inside and outside Raw.
3. Exact split-run delimiters and allowed whitespace obey existing dedicated paragraph/row rules. Verify first-valid-close precedence, non-nesting, literal inline/wrong-parent/escaped/opaque lookalikes, and shared errors for unclosed Raw and active stray or invalidly placed delimiters. Paragraph delimiters in different cells, row delimiters in different tables, mixed paragraph/row delimiters, and cross-part/root layouts cannot form a region through common-ancestor promotion.
4. Raw works in both conditional branches and nested `For`/`If` combinations for zero, one, and multiple iterations. Cover multiple paragraphs in one cell, complete tables and nested tables, empty Raw where structural validity permits, and source-to-copy rendering without reinterpreting literal or inserted content. Literal controls do not modify the outer stack.
5. Safe merged Raw bodies and complete merged tables are preserved; deletion that compromises vertical merges, section boundaries, or required cell content is rejected with an authoring remedy. Existing `For`/`If` row merge restrictions remain in force. No merge repair or section relocation is introduced.
6. Preserve unaffected and valid fully enclosing comment/bookmark ranges without identity changes. Reject unsupported local, split, or crossing arrangements in `For`/`If`, including Raw nested inside them, irrespective of runtime branch/cardinality. Existing fixtures may supply enclosure coverage; substantial new preservation logic requires the D3 Product Owner confirmation.
7. In standalone Raw, preserve valid comments/bookmarks/internal links and safely remove complete delimiter-local bookmarks and ranged/point comment attachments. Verify surviving definitions/references remain valid, unrelated comments and media relationships survive, split deletion is rejected, and links to removed targets fail without remapping. Significant cleanup machinery triggers the D2 tradeoff review rather than implicit scope expansion.
8. Reject an internal link inside `For`/`If` even with an unaffected target. Preserve supported stable links, unrelated pre-existing unresolved anchors outside operations, external relationship fragments, and opaque Word fields/content controls. Retain a meaningful permission-copying regression and package-wide drawing uniqueness using reservations in unchanged auxiliary parts and drawings inside repeated Raw.
9. Existing revision families fail in the current inspection footprint, including Raw, inactive branches, and controls, with shared deterministic diagnostics and Word remediation. Retain existing exclusions and accept revision-free templates with only the Track Changes setting enabled; no package-wide revision scan is introduced.
10. No cloning/remapping mechanism remains solely for rejected comment/bookmark/link cases. Document retained permission matching/ID allocation, drawing allocation, range/reference preflight, and narrow Raw comment cleanup as required dependencies. Mutable operation data remains isolated across repeated/concurrent use.
11. Supported generated documents pass applicable Open XML validation and preserve relevant literal content, formatting, relationships, identities, and semantic references; checks must detect data loss or dangling structures as well as schema errors. Analyzer and Generate agree on restrictions and deterministic locations, and generation failures return no partial DOCX.
12. Update and verify all five maintained sample template sources, their output expectations, current fixtures/helpers, README, and Playground bundling. Cover affected Core/CLI/API/frontend contracts proportionately, reusing valid evidence and existing regressions rather than expanding every case across every surface. Public requirement and occurrence schemas remain unchanged. Disclose unexecuted checks and use representative manual Word inspection where needed; browser automation requires explicit Product Owner approval.

## 8. Out of scope

1. Compatibility aliases or migration warnings for the old closing markers.
2. Expanding support for complex Word editing constructs beyond the compatibility policy.
3. Processing expressions inside Word fields, content controls, or other explicitly opaque constructs.
4. New data types, formatters, custom formatting extensibility, or XLSX support.
5. A new orchestration framework, new Feature status file, or unrelated repository restructuring.
6. General comment/bookmark relocation or reconstruction, identity cloning or target remapping to support Raw, table-merge repair, section-property relocation, package-wide revision detection, and unrelated OOXML cleanup or repair.

## 9. Architecture constraints and Developer handoff

1. Keep one shared language grammar and source plan for Analyzer and Generate. Establish Raw extents before ordinary tokenization of their contents; filtering dependencies or render operations afterward is insufficient. Keep the DOCX adapter's opacity, field state, source locations, exact sibling rules, and independent story roots.
2. Render Raw through existing source-to-output mappings, removing delimiters without a new body copy or JSON resolution. Retain parse-once rendering for outer copies and inserted values. Keep DOCX-specific processing separate from shared language/data APIs and mutable state per operation.
3. Reuse source-root range ownership for simple enclosing-range preservation and selective rejection. Apply the distinct Raw deletion rules before removing obsolete clone/remap code. A local comment-definition update may remain; comment cloning and ID allocation may not be retained under that justification.
4. The subsequent Implementation Plan must trace required removals and retained dependencies, describe safe delimiter/annotation deletion and the narrower Raw merge check, and map focused verification to section 7. Optional tree-based dependency traversal or allocator extraction must remain proportionate; no general framework or broad rewrite is required.
5. If evidence shows D2 cleanup needs significant machinery or D3 preservation needs substantial extra logic, present the practical value, simpler alternatives and actionable rejection, implementation cost, and integrity risks for Product Owner confirmation. Do not silently change the approved behavior. This architecture documentation task does not create that plan or authorize implementation.

## 10. Approval

1. **Decision approval:** The Product Owner explicitly approved D1–D5 in the architecture continuation request on 2026-10-19. Sections 4–5 incorporate those decisions as normative requirements, including the D2/D3 complexity conditions.
2. **Complete revised Feature Spec approval:** The Product Owner approved the complete revised Feature Spec and Implementation Plan on 2026-10-19. This does not constitute Feature acceptance or Git authorization.

## 11. Architecture assessment for Product Owner review

1. [Architecture assessment](architecture-assessment.md) records the original recommendation dispositions, current approach, source evidence, retained dependencies, and conditional escalation points. It is not an Implementation Plan.
2. D1–D5 have been incorporated into this specification and the target compatibility policy. No additional material product decision is currently established by source inspection; the conditional D2/D3 tradeoffs remain subject to the explicit confirmation process in section 9.5 if triggered by implementation-planning evidence.
