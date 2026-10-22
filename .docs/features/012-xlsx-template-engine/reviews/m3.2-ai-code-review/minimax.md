## Feature 012 — Independent Review of M3.1 and M3.2

**Reviewed:** `48ceb65` (M3.2), parent `ad253d4` (M3.1), base `868dd6a` (M2). Working tree clean before and after; no repository file was modified (probe harness lives in `/private/tmp/f012probe`).

### 1. Verdict

**Needs revision — minor, non-blocking for output correctness.**

The M3.1 numeric-level grammar and the M3.2 hierarchy/ancestor binding are substantively correct: I independently reproduced 5- and 6-level parsing, extreme/nonconsecutive levels, footer-only ancestors, multirow coalescing, opaque `$path` keys, empty/single/equal-level inputs, missing-vs-null ancestors, sibling transitions and 5-level ancestor resolution, and all are correct. No confirmed defect produces wrong data, wrong output or an M1/M2 regression (623 Core + 17 transport tests pass).

Two confirmed defects remain, both in the M3.1/M3.2 interaction area: **F1** (row-order field scopes contradict the new level-based grammar) and **F2** (stale user-visible capability message). F1 must be fixed before M3.4 enables nested emission; F2 is cosmetic. Neither blocks starting M3.3 if the PO accepts F1 as an explicit M3.4 dependency — that is a PO decision, not mine.

---

### 2. Confirmed findings

#### F1 — Medium (P2): field-context scopes are derived from source row order, not from the parsed numeric-level tree

**Locations**
- `MonoReport/Services/Xlsx/XlsxTemplatePlanner.Controls.cs:63-91` (`CellScopes`, esp. `:81-88` — vertical branch pops/keeps ancestors by *row order*).
- `MonoReport/Services/Xlsx/XlsxTemplatePlanner.cs:86-88` (`ParentDepth > scope.Length` → `template.invalid_context`) and `:92-94` (canonical basis `scope[ParentDepth] ?? "$"`).
- Contradicted by `MonoReport/Services/Xlsx/XlsxTemplatePlanner.Vertical.cs:32-74` (level-based tree) and `MonoReport/Services/Xlsx/XlsxTemplatePlanner.Rows.cs:10-22` (level-stack ancestor chain at runtime).

**Cause.** M3.1 parses nested grammar by *numeric level*, and the M3.2 runtime scope is built from the level stack (`Rows.cs:15-19`). But the pre-flight field validation and the analyzer's canonical paths still use the M1-era row-order scope chain. For a layout where a child block precedes its parent block in row order, the two disagree.

**Reproduction** (executed, out-of-repo probe; template `PrintArea B3:D4`, one control column A):

| Layout | Rows | Public Analyze/Merge result | Canonical path of the ancestor field |
|---|---|---|---|
| A (parent first) | A3=`$l(10).Items`, A4=`$l(40).Items`, B4=`{../../Name}` | `template.xlsx_unsupported_feature` @ A3 (nested guard) | `$.Name` (root) |
| B (child first) | A3=`$l(40).Items`, A4=`$l(10).Items`, B3=`{../../Name}` | **`template.invalid_context` @ B3** ("A template path is invalid in its current loop context") | — (rejected earlier) |
| C (child first, `../`) | A3=`$l(40).Items`, B3=`{../Name}` | `template.xlsx_unsupported_feature` @ A3 | **`$.Name` (root)** — should be the level-10 ancestor (`$.Items[].Name`) |

Both A and B parse to the *same* level hierarchy under the M3.1 grammar (verified via the internal `XlsxVerticalRegion.Root` tree: root level 10 → child level 40), yet they surface different public codes and locations.

**Expected vs actual.** The scope chain for a row must follow the parsed tree. Expected for B: the same nested-grammar rejection as A. Actual: an unrelated `template.invalid_context` at the field cell. Expected for C: canonical ancestor path `$.Items[].Name` (the level-10 ancestor, which is what `RowScope` actually resolves at runtime). Actual: `$.Name` (root) — the reported analysis path contradicts the runtime value.

**Impact.** No incorrect report is reachable today: nested runs are rejected by the preflight guard in both layouts and no partial output is produced (verified). The user-visible effect is misleading/divergent diagnostics for two grammatically equivalent layouts, plus a wrong canonical analysis path. It becomes a *value* correctness issue when nested emission is enabled (M3.4) if the row-order scope is not replaced by a tree-based one.

**Note on provenance.** The row-order scope model is M1-era (`CellScopes` introduced in `9711d60`, unchanged since) and the same `invalid_context` was already reachable in M2, because `Prepare` (field validation) runs before `Structure` (nested rejection). M3.1 is what made the grammar *accept* these layouts, so the inconsistency is now inside M3.1's remit. This finding was also reported by the DeepSeek review stored in `.docs/features/012-xlsx-template-engine/reviews/m3.2-ai-code-review/`; I reproduced it independently and additionally characterized the wrong canonical path (case C), which that review did not report.

**Suggested correction (PO/implementer decision).** Derive vertical field scopes from `XlsxVerticalRegion.Root`, or run the nested preflight guard before field-context validation and defer analyzer scope projection to M3.14. Add paired regression tests for parent-first and child-first layouts.

#### F2 — Low (P3): stale nested-guard diagnostic

`MonoReport/Services/Xlsx/XlsxGeometryPlanner.Reports.cs:23-28` still says *"Nested vertical groups require M3 hierarchy binding and placement; only their source grammar is implemented."* M3.2 delivered hierarchy binding; only placement/emission remain pending. The message is published verbatim through `Analyze`, `Merge`, CLI and HTTP, so it misstates the current capability state.

---

### 3. Missing test coverage (prioritized by risk)

1. **High — F1 pair.** No test covers a child-before-parent nested layout *with ancestor field references*. `Feature012NestedTests.cs:131-144` builds only plain-text column B ("Authored row N"), so the grammar tests never exercise field scopes; the two child-first sequences that do exist (`"10,90,40,10"`, `"90,40,10"`) assert only the tree. Add: Analyze/Merge parity for parent-first vs child-first layouts, and canonical-path assertions for `{../}`/`{../../}`.
2. **High — missing-level ancestor divergence (see §5 R1).** `Feature012BindingTests.Hierarchy.cs:167-181` currently *pins* the root fallback (`"{../Name}"` → `"ROOT"`) for a single-level template with multi-level data, which contradicts plan §4.2.5 ("keep their logical ancestry where needed for parent fields") and spec §3.2.7. M3.5 must revisit this test; it should at least be marked as expected-to-change.
3. **Moderate — five nesting levels.** Only 3 levels are tested. I verified 5 and 6 levels parse correctly (`L10→L40→L90→L140→L190`, `L190` body) and that a 5-level data chain binds correctly (`$.Items[4] | $.Items[3] | $.Items[2] | $.Items[1] | $.Items[0] | $` → `L190/L140/L90/L40/L10/ROOT`). Add grammar and binding regressions.
4. **Moderate — nested region × horizontal product.** M3.2 covers a 2×1 product; the 2×2 four-intersection ancestor case (spec §10 A9) is untested.
5. **Moderate — multiple regions on one sheet** (nested + independent single-level region; nested region next to a second marked run separated by an unmarked row).
6. **Moderate — interleaved level orders** (e.g. template `10/40/20`, see R2): no diagnostic and no test; a negative test would force an explicit decision.
7. **Low — cosmetic:** `XlsxTemplatePlanner.Rows.cs:111` renders "Supply exactly **1 records**…".
8. **Low — documentation:** `docs/xlsx-template-compatibility.md` is still titled "Feature 012 M2" and item 8 only says nested vertical hierarchy "remains M3"; it does not mention that nested runs are now parsed and rejected with a specific ambiguity diagnostic. `README.md:5` remains accurate.

---

### 4. Tests actually executed

| Command (repo root) | Result |
|---|---|
| `dotnet build MonoReport.Tests/MonoReport.Tests.csproj -c Release --no-restore` | Build succeeded, 0 warnings / 0 errors |
| `dotnet test MonoReport.Tests … --filter 'Feature012NestedTests\|Feature012BindingTests\|Feature012AnalyzerTests'` | **57 passed**, 0 failed |
| `dotnet test MonoReport.Tests … --filter 'Feature012\|Feature011\|UnifiedTemplateTests\|Feature010'` | **623 passed**, 0 failed, 0 skipped |
| `dotnet test MonoReport.Api.Tests … --filter 'Feature012'` | **12 passed** |
| `dotnet test MonoReport.Cli.Tests … --filter 'Feature012'` | **5 passed** |
| Out-of-repo probe (`/private/tmp/f012probe`, assembly `MonoReport.Tests` for internals access) | Parsed trees for 20 level sequences; Analyze/Merge codes, locations, occurrences and canonical paths for 4 nested layouts; 5-level bind/scope/resolve check |

Probe highlights: `[10,40,90,140,190]` → `L10[3/–/–] L40[4/–/–] L90[5/–/–] L140[6/–/–] L190[–/7/–]`; `[10,40,20]` → `L10[3] → L20[footer5] → L40[body 4]`; `[40,10,40]`, `[10,40,10,40]`, `[10,40,90,40,90,40,10]`, `[90,10,40]` rejected with `template.invalid_structure` + "unmarked control row" remedy; `[2147483647,0,-2147483648]` parsed with the minimum level as root.

No test was created or modified in the repository; no Git operation was performed.

---

### 5. Limitations and unverified assumptions

1. **Microsoft Excel E1 was not run** and remains pending; no calculated-value evidence exists for M3.1/M3.2.
2. I did not rerun the full 693-test solution suite; my regression evidence is the623-test Core Feature010–012/shared-language selection plus the API/CLI Feature012 groups. The developer's claimed full-suite result was not independently reproduced.
3. **M3.2's binding is not publicly reachable.** The nested guard in `XlsxGeometryPlanner.Reports.cs:23` runs in `Preflight` (`XlsxTemplateHandler.cs:50`) *before* `BindSheet` (`:59`), so `BindHierarchy`/`RowScope` ancestor chains are exercised only through internal APIs. M3.2 evidence is therefore unit-level; no user-visible nested behavior exists yet (correctly stated in `implementation-progress.md`).
4. **R1 — deferred divergence (not an M3.2 acceptance failure, but tracked).** `XlsxTemplatePlanner.Rows.cs:57` binds the hierarchy only when `region.Root.Children` is non-empty. For a single-level template with multi-level data, the *omitted* ancestor level is therefore not in the scope chain, and an ancestor field silently resolves to the **root object** instead of the omitted ancestor record — contradicting plan §4.2.5 and spec §3.2.7. The M3.2 progress defers this to M3.5; the newly added test at `Feature012BindingTests.Hierarchy.cs:167-181` currently locks the divergent behavior.
5. **R2 — template tree vs. data hierarchy can disagree.** For interleaved level orders (e.g. template `10/40/20`) the M3.1 tree yields `L10 header → L20 footer → L40 body`, while the M3.2 level stack makes `L40` and `L20` siblings for data ordered `10,40,20`. No diagnostic is raised. This falls under the spec's "assume otherwise valid nesting" allowance (§4.1.5), but M3.3/M3.4 must reconcile the two hierarchies or reject the mismatch.
6. **R3 — `ChildFirstIndices` is lossy.** It stores the child's first *record* index only (`XlsxTemplatePlanner.Hierarchy.cs:20,31`); with arity > 1 the child's records span `[FirstIndex, FirstIndex+arity)`, so a consumer cannot recover the child's level, row or records from the index alone. Also, `XlsxLogicalRow` is a `record` whose generated equality now includes the `Parent` chain (recursive comparison, acyclic but depth-proportional).
7. No browser automation, spreadsheet-application execution, capacity measurement or RealHost V10 evidence was produced or required for these internal checkpoints.
8. I did not re-derive the accepted M1/M2 baselines from scratch; regression conclusions rest on the existing suite plus the targeted runs above.