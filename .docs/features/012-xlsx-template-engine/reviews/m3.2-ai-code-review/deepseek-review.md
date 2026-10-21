# DeepSeek — Independent Review of Feature 012 M3.1 + M3.2

**Verdict: Needs revision.** One confirmed reproducible P2 finding; one cosmetic P3 finding.

## F1 — P2: Source-row scopes disagree with numeric-level tree

**Locations:**

1. `MonoReport/Services/Xlsx/XlsxTemplatePlanner.Controls.cs:63–91` (especially 78–88).
2. `MonoReport/Services/Xlsx/XlsxTemplatePlanner.cs:86–94`.
3. Contrast with `XlsxTemplatePlanner.Vertical.cs:32–74` and `XlsxTemplatePlanner.Rows.cs:10–22`.

**Cause:** `CellScopes` walks control rows top-to-bottom and derives ancestor scopes from source marker order. M3.1 accepts nested grammar where child blocks precede their numeric-level parents (footer-only or mixed shapes). `Prepare` still validates `ParentDepth` and canonicalizes dependencies against source-row scopes rather than the parsed tree.

**CLI reproduction (reported as executed):** Based on `research-spikes/nested-evidence/nested-template.xlsx`, replace rows 9–10 with:

Variant A (parent first):

```text
A9=$l(10).Items    B9={Name}
A10=$l(40).Items   B10={../../Name}
```

Variant B (child first):

```text
A9=$l(40).Items   B9={../../Name}
A10=$l(10).Items  B10={Name}
```

Both yield the same numeric-level hierarchy (level 10 root, level 40 child). Expected for both: `template.xlsx_unsupported_feature` from nested preflight. Actual: variant A gets `template.xlsx_unsupported_feature`; variant B gets `template.invalid_context` at B9. `--analyze` and generation reproduce; no partial output is produced.

**Impact:** Misleading validation for accepted nested grammar; latent analyzer/runtime scope disagreement. Nested emission is still gated, so no generated report corruption is currently reachable.

**Proposed fix:** Compute vertical field scopes from parsed `XlsxVerticalRegion.Root`, or run nested preflight before field-scope validation and explicitly defer full analyzer semantics to M3.14. Add paired regression tests.

## F2 — P3: Stale diagnostic

`XlsxGeometryPlanner.Reports.cs:23–28` still says hierarchy binding is missing, although M3.2 implemented it. Update message to say placement/emission is pending.

## Missing coverage

1. High: footer-only / child-before-parent grammar combined with `../`, `../../`, `../..` and canonical analysis paths.
2. Moderate: five levels; mixed header/footer four-level chains.
3. Moderate: multiple nested regions or nested plus single-level region on one sheet.
4. Moderate: explicit `Root.Body` assertions for multirow leaf body.
5. Low: empty/null nested `Items`; 2×2 horizontal product and ancestor scopes (future M3.14 relevance).

## Executed checks (as reported)

1. Nested + binding tests: 51 passed.
2. Affected selection: 279 passed.
3. Release build: zero warnings/errors.
4. CLI probes for both variants under `/private/tmp/f012-m3-review`; different diagnostic codes, no partial output.
5. Clean Git working tree after review.

Full 693-test suite was not rerun; E1 not run. M3.3+ out of scope. M3.5 must revisit single-level regions with multi-level data.

## Original review note

This is a structured preservation of the review shared in the conversation, not a verbatim full transcript.
