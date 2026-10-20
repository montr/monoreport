# Feature 012 — Architecture Research Brief

**Audience:** Architecture agent
**Mode:** Investigation and executable spikes only. Do not implement the full feature.
**Primary input:** `feature-spec.md`
**Revision:** 2026-10-20; findings are consolidated in [architecture-assessment.md](architecture-assessment.md), including Product Owner clarifications and the SDK-only formula investigation. This brief does not authorize production implementation.

## 1. Objective

Determine a technically sound approach for XLSX template expansion on two axes while preserving user-authored layout, merges, styles, and formula semantics. Identify constraints early enough to estimate and stage implementation without avoidable rewrites.

## 2. Research questions

### 2.1. Expansion geometry

1. Define nested top-control `$h(level)` blocks, parent-only columns, separate control rows for each depth and data-driven repetition counts. Do not use the superseded fixed leaf-tile/bare-h model.
2. Show how `$header(1).Bids` and `$header(2).Criteria` map onto parent D:E/child E, including zero-width D, one authored company merge and zero/one/many criteria.
3. Demonstrate the approved bounded adjustment of existing parent-owned merges and collapse to an ordinary cell, preserving content/formatting without creating unrelated merges or inventing styling.
4. Test multiple horizontal header rows and vertical multirow blocks.
5. Determine how a single horizontal region coexists with multiple tables and with independent vertical regions on one worksheet.
6. Use the approved unmarked-control-row boundary between vertical regions; retain printed separators once, including author-defined zero-height rows. PrintArea/controls are the only region mechanisms; no named report regions or new boundary tokens.
7. Physically crop disposable external rows/columns/content and transform affected references/settings. Verify both axes using the same final geometry plan.

### 2.2. Formulas

Compare at least these approaches:

1. Rewrite formula references and ranges using parsed Excel/OOXML formula semantics.
2. Use hidden helper/boundary rows or cells so formulas naturally cover expanded ranges.
3. Use a hybrid approach.
4. Compare ClosedXML.Parser with a bounded internal reference parser/transformer using the existing Open XML SDK. Identify the SDK's storage/validation capabilities separately from token/reference semantics. Do not implement an evaluator or general Excel grammar.

Test:

1. `SUM`, `COUNT`, and `SUBTOTAL` in group footers.
2. The same functions in group headers **above** the detail rows.
3. Grand totals across multiple nested groups.
4. Absolute, relative, and mixed references.
5. References spanning both row and column expansion.
6. Cross-sheet references and other reference types, identifying unsupported cases explicitly.
7. Recalculation flags and resulting workbook behavior in a compatible spreadsheet application.

For each approach report correctness, complexity, maintainability, unsupported constructs, performance, and diagnostic requirements. Never silently accept wrong formulas.

### 2.3. OOXML fidelity and library choice

1. Evaluate permissively licensed C#/.NET libraries and native OOXML operations.
2. Verify preservation of styles, number formats, merges, row/column dimensions, print area, print settings, and hidden states.
3. Probe advanced features: images, charts, conditional formatting, named ranges, Excel tables, validation, comments, and other nonessential metadata.
4. Record exact licenses and dependencies, with no GPL/AGPL or proprietary requirement.
5. Establish whether generation can be performed without Excel installed.
6. Investigate native numbers/Booleans, serial dates and 1900/1904 workbook systems; distinguish preservation/serial binding from new ISO conversion semantics.

### 2.4. Data and integration

1. Confirm positional mapping of a complete `Bids × Criteria` matrix grouped by contiguous `$path` values.
2. Confirm that `$path` is a grouping key only; input order is preserved.
3. Confirm first-record row-field resolution and intersection-record field resolution.
4. Inspect existing project data-resolution and formatting APIs for reuse; distinguish actual reusable functionality from assumptions.
5. Propose safe diagnostics for missing levels, incomplete intersections, unsupported structures, and formula failures.
6. Normalize an empty horizontal product as one descriptor per logical vertical row; retain ordinary rows and safe parent-only columns. Do not require empty vertical data.
7. Extend existing Core, CLI and HTTP interfaces with DOCX backward compatibility; do not propose separate public XLSX API/routes. Identify additive format metadata and structured success-warning transport work for the Implementation Plan.

## 3. Required executable spikes

1. **Geometry spike:** real XLSX parent D:E/child E, two header/vertical levels, authored parent merge and zero-width column. Assert zero/one/three/seven criteria, merge projection/collapse, ordinary rows under empty horizontal input and physical crop.
2. **Formula spike:** a real XLSX with `SUM`, `COUNT`, `SUBTOTAL`, subtotals, and top/bottom totals. Generate nested expansion and assert final formula references/ranges and recalculated results using a suitable test environment.
3. **Multi-table spike:** two independent marked runs separated by an unmarked control row, retained zero-height separator, nested groups/formulas, two worksheets and one horizontal region maximum per worksheet. No named report regions.
4. **Compatibility spike:** workbook with representative print and formatting features; record preservation and failure behavior.
5. **Dependency spike:** replay existing and nested formula requests through the SDK-only internal transformer and optional external comparator; assert supported results and deliberate rejection boundaries. Verify the internal runtime graph has no ClosedXML dependency.

Use small, reproducible fixtures and automated assertions. Do not treat a visually plausible output as proof of correct formulas.

## 4. Deliverables

1. `architecture-assessment.md`: current decisions/design, experimental evidence, alternatives/trade-offs, bounded capability matrix, risks and remaining verification. The original split deliverable filenames are superseded.
2. Updated `feature-spec.md` and `implementation-roadmap.md` incorporating approved decisions consistently; the roadmap remains input to an approved Developer-owned Implementation Plan.
3. Spike source code and XLSX/JSON fixtures, with commands to reproduce.
4. Proposed internal component boundaries and data/geometry model.
5. Recommended implementation sequence, risk-based estimates or estimate ranges, and validation gates.
6. Suggested **minimal** spec amendments where experiments demonstrate a real ambiguity or impossibility.

## 5. Decision constraints

1. Do not change the requirement that the user authors all merges and formatting.
2. Do not introduce `$h1` / `$h2` input fields as a prerequisite.
3. Do not sort `$path` groups.
4. Do not require nested input JSON for vertical records.
5. Do not assume a single table or worksheet.
6. Do not start full implementation until findings are reviewed and accepted.
