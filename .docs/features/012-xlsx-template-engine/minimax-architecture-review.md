# Independent Architecture Review — Feature 012, Bounded Vertical Logical Block Schedule (Alternative B)

**Scope of this review:** architecture only, read-only. I read the Feature 012 spec, the amended `implementation-plan.md`, `preprocessing-architecture-assessment.md`, `bounded-logical-block-schedule-proposal.md`, the current XLSX production code and the M3.3 working-tree changes, and the M3.3 test additions.

**What I did not do:** I did not build, run tests, modify any file, or touch the Git index/working tree. All statements about runtime behaviour are derived from source reading, except one item (§4.1) which I verified empirically in an isolated scratch project outside the repository. The current working tree's compile/test state is therefore **unverified by me**.

---

## 1. Reconstruction: old and new architecture

### 1.1 Before the schedule (the "old" architecture)

The pipeline was:

```
Read → Plan/Prepare → Structure → Preflight → BindSheet → ExpansionGeometry → EmitCells
      → ProjectMerges → PlanHeaderOwners → BindInstances → Formula Finalize → Render
```

Responsibilities were distributed as follows:

- **Parsed template hierarchy** — `XlsxTemplatePlanner.Vertical.cs:32` (`VerticalTree`) peels numeric levels from the ends of a marked run into `Header`/`Body`/`Footer` blocks per `XlsxVerticalNode`. This is the single authoritative source tree (`XlsxTemplatePlan.cs:58-63`), corrected in the F1 cycle and shared with static scope projection (`XlsxTemplatePlanner.Controls.cs:75`).
- **Bound data hierarchy** — `XlsxTemplatePlanner.Hierarchy.cs:8` (`BindHierarchy`) builds `XlsxLogicalRow.Parent` / `ChildFirstIndices` via a numeric-level stack. Never sorts, never interprets `$path` slashes.
- **Vertical repetition** — `XlsxGeometryPlanner.Expansion.cs` did this **inside the placement layer**: it counted rows matching `region.Level` for the height budget (old lines 16-20) and then re-selected them for emission (old lines 28-40), emitting the `generation.xlsx_missing_level` warning from the same place.

The genuinely mixed responsibilities were therefore narrow and identifiable:

1. **The placement layer decided which logical items exist** — omission diagnostics were produced by the geometry planner, not by the component that knows the hierarchy.
2. **A duplicated count/select predicate** — the same `row.Level == region.Level` test ran twice for two different purposes.
3. **No vertical *instance identity* concept** — `XlsxAxisSpan` carried only a `VerticalInstancePath`, and the flat model conflated "the region's level" with "the item that owns this block". There was simply no way to say "this physical row belongs to *this* logical item under *this* parent", which is exactly what nesting requires.

Point 3 is the important one. The old design did **not** merely "mix responsibilities"; it lacked the concept that nesting forces into existence. Any correct extension of the old design (option A) would have had to introduce an owner identity anyway. I think the proposal understates this: it presents B as an optional seam, when in fact the *shape* of B's central idea was already unavoidable.

### 1.2 After the schedule (the "new" architecture)

```
Read → Plan/Prepare → Structure → Preflight → BindSheet → ScheduleSheet → PlanReports
      → ExpansionGeometry(schedule) → EmitCells → ProjectMerges → PlanHeaderOwners
      → BindInstances → Formula Finalize → Render
```

New types (`MonoReport/Models/Xlsx/XlsxLogicalSchedule.cs`, 10 lines):

```csharp
internal enum XlsxLogicalBlockRole { Header, Body, Footer }
internal sealed record XlsxLogicalBlockInstance(XlsxVerticalBlock Source, XlsxLogicalBlockRole Role, VerticalInstancePath Owner);
internal sealed record XlsxVerticalRegionSchedule(int RegionId, ImmutableArray<XlsxLogicalBlockInstance> Blocks);
internal sealed record XlsxLogicalSheetSchedule(uint SheetId, ImmutableArray<XlsxVerticalRegionSchedule> Regions);
```

`ScheduleSheet` (`XlsxTemplatePlanner.Scheduling.cs`) currently implements only the flat case: it throws for nested trees, matches `row.Level != region.Level`, emits the omission warning, and emits one `Body` block per matching row. `ExpansionGeometry` now consumes `schedule.Regions[index].Blocks` and no longer touches bound rows (confirmed: `data.Rows[...]` no longer appears in `Expansion.cs`; the only remaining `XlsxMissingLevel` emission site in the whole production tree is `Scheduling.cs:24`).

**Net cost of the change: ≈43 lines** (10 model + 33 service) replacing ≈12 lines of selection logic inside `ExpansionGeometry`, plus a changed signature on `PlanReports`/`ExpansionGeometry` and one new call site in `XlsxTemplateHandler.cs:60`.

---

## 2. Strengths of the adopted approach

1. **The schedule is the right *size*.** It is not an IR, not a second template tree, not a wide-record schema, not a synthetic worksheet, not a new injected service. For a product that is ~2,350 lines of XLSX code today, +43 lines and +3 type names is proportionate. The explicit "stop and fall back to A" clause in the amended plan (§4.2.9) is the correct governance for a decision of this size and is rare enough to be valuable.
2. **It moved a real decision to a real owner.** The `generation.xlsx_missing_level` warning now has exactly one emission site, in the component that knows the hierarchy. That is a genuine responsibility improvement, not a relabelling.
3. **It created the right seam for the hardest upcoming work.** M3.4's enter/exit traversal (footer-first ordering, header-before-children-before-footer) is the single most intricate algorithm left in the feature. Producing an ordered `(Source, Role, Owner)` list makes it assertable **without** crop, worksheet dimensions, XML or horizontal data. Under option A the same testability is only reachable by exposing an internal helper that geometry happens to call — which is, in substance, B minus the model type. The marginal difference between A and B is therefore smaller than the assessment's framing suggests.
4. **It matches the spec's mandated separation better than A does.** Spec §9.1 requires parsing, template semantics, data resolution, expansion planning, formula handling, rendering and diagnostics to be separated "at least conceptually". B puts *logical* ordering in the planner and *physical* placement in the geometry planner. A puts both in the geometry planner. B is the more spec-aligned of the two, and the assessment under-uses this argument.
5. **It did not smuggle in speculative generality.** No horizontal scheduler, no eager field values, no copied ancestors, no per-cell structure, no synthetic field names. The proposal explicitly rejects horizontal schedule extraction, which is the single most likely place for this architecture to have ballooned.
6. **The owner path is the right shape for M3.7.** A root-to-current index sequence is prefix-extensible, so "is this cell inside this group?" is answerable later by prefix, which is exactly what group header/footer range binding will need. A flat "owner id" would have been wrong.
7. **Honest about evidence limits.** Both documents state repeatedly that nothing was executed and that no performance claim is made. I checked: §8.4, §2.3, §9.6 of the proposal and §2.3/§7.8 of the assessment all disclaim measurement rather than implying it.

---

## 3. Weaknesses and risks, ranked

### 3.1 HIGH — Confirmed defect: instance identity has reference semantics, not value semantics

`XlsxTemplatePlan.cs:27`:

```csharp
internal sealed record VerticalInstancePath(int RegionId, ImmutableArray<int> LogicalRows);
```

`ImmutableArray<T>` does **not** implement structural equality — `ImmutableArray<T>.Equals` compares the underlying array by reference. Because `VerticalInstancePath` is a `record`, its synthesised `Equals`/`GetHashCode` therefore compare the `ImmutableArray` member by reference.

I verified this empirically in an isolated scratch project (outside this repository):

```
record equal (separate arrays): False
record equal (shared array):   True
ImmutableArray.Equals(other):   False
HashSet<P> count for two equal-content instances: 2
```

Consequences inside this codebase — every one of these types inherits the defect:

- `XlsxLogicalBlockInstance`, `XlsxAxisSpan`, `CellInstanceId`, `XlsxContentOwner`, `XlsxDisplacedSource`.

Live consequences:

- `XlsxGeometryPlanner.Owners.cs:60` — `owners[i].Vertical == group.Key.Vertical` is a **record equality** comparison used to *remove* a stale content owner. It only works because both sides happen to be the *same object*. If a merge owner path and a placement owner path are ever reconstructed as equal-content-but-distinct objects, the removal silently fails and a duplicate content owner survives into `ProjectReference`.
- `XlsxGeometryPlanner.Merges.cs:85-86` — the merge candidate list is `RowSpans.Select(span => span.Vertical)...ToArray()`. The M3.4 dedup rule the plan mandates ("distinct compatible owners") will be implemented with `Distinct()`, i.e. with this equality. It will dedup correctly **only** if header and footer blocks of one item share one path *object*, which the proposal requires by discipline ("share one canonical Owner path per scheduled logical item") rather than by construction.
- `XlsxTemplatePlanner.Instances.cs:58-63` — `prepared` is a `Dictionary<CellInstanceId, …>`; `prepared.Add` throws on a duplicate key. Two cells with the same source, equal-content owner paths and the same horizontal path would collide.

This is not a hypothetical: M3.4 is precisely the change that starts constructing owner paths in several new places (header, body, footer, merge candidates, formula scope). The proposal itself flags the hazard in §3.1.5 — but it treats it as a *usage discipline* problem ("do not recursively hash/compare `XlsxLogicalRow` records"). It is actually a *type* problem: the type does not express the identity the architecture relies on.

**Smallest practical correction (≈10 lines, one file):** give `VerticalInstancePath` value semantics over the sequence — either override `Equals`/`GetHashCode` to compare `LogicalRows` element-wise, or replace `ImmutableArray<int>` with a value-equality sequence. Cost: one record plus one test asserting that two independently constructed equal paths compare equal and hash equal. Benefit: the merge-dedup rule, the owner-removal rule and the placement-key rule become correct *by construction* rather than by caller discipline, and the whole class of "silently wrong owner" failures closes. There is no downside: region id + index sequence *is* the identity, and two paths with the same region and same sequence denote the same logical item.

### 3.2 HIGH — Credible risk: the compatibility check must be a general invariant, not the single R2 example

The approved R2 policy (plan §4.2.8) covers one shape: source `10/40/20` versus input `10 P,40 X,20 Y`. That is the right *diagnostic* for that case, but as an *implementation rule* it is too narrow, and the underlying hazard is broader than the documents acknowledge.

The hazard: **the static scope depth comes from the source tree; the runtime scope depth comes from the data hierarchy; nothing checks that they agree.**

- `XlsxTemplatePlanner.Controls.cs:75-86` (`CellScopes`) builds each row's scope as `[collection] + ancestors` walking the **source tree** — so a level-90 node gets scope length 3, and `ParentDepth ≤ scope.Length` is the static validation (`XlsxTemplatePlanner.cs:87`).
- `XlsxTemplatePlanner.Rows.cs:10-22` (`RowScope`) builds the runtime scope by walking `row.Parent` — i.e. the **bound data hierarchy** — and its length is the data's ancestor count.
- `DataResolver.cs:37`: `value = path.ParentDepth < scope.Values.Count ? scope.Values[path.ParentDepth] : scope.Root;`

Consequence: if the data's ancestor chain is *shorter* than the source tree's depth, an over-deep `../..` does not error — it silently resolves to the **document root**. A template analysed clean would render a wrong value with no diagnostic.

Concrete scenario: source levels `10/40/90` (scope lengths 1/2/3). Input `{"$level":10}`, `{"$level":90}` — a level-90 item whose nearest shallower ancestor is the level-10 item (the level-40 level was skipped). `BindHierarchy` gives it **one** ancestor; the source tree gives its block **two**. A cell in the level-90 block containing `{../../Name}` passes static validation (depth 2 < scope length 3) and at runtime `ParentDepth 2` is out of range for `Values.Count == 3` → **root's `Name`**, silently, where the author expected the grandparent's.

The narrow R2 rule does not catch this: it is about *sibling* incompatibility, not *depth* incompatibility. The general invariant that catches both is:

> When a data item is paired with a source node, the data item's ancestor chain length must equal the source node's tree depth, and each ancestor's `$level` must equal the corresponding source node's level.

Applied to the documented R2 case (data `10 P, 40 X, 20 Y` binds X and Y as siblings under P, both at depth 1, while the template places level 40 at depth 1 and level 20 at depth 0 as the level-10 node's footer), the general invariant rejects the shape too — with a slightly different message than the proposal's. **Recommendation:** keep the proposal's actionable message for the documented case (it is genuinely more useful to the author: "re-author the source hierarchy or supply input in the intended parent-before-descendant order"), but implement the check as the general depth/level invariant so other shapes cannot slip through. The general rule also gives M3.5's missing-level policy a defined base instead of leaving it to be discovered later.

### 3.3 MEDIUM — Confirmed inconsistency: two vertical-matching rules now exist

`XlsxTemplatePlanner.Scheduling.cs:20` matches data rows to template blocks with `row.Level != region.Level`. That is a **level-comparison rule**. It happens to coincide with the node-scoped rule for a flat region (root level == region level, no children), but M3.4 must replace it with a node-scoped rule (data item ↔ specific source node). The proposal's own invariant §3.2.7 says "No component reconstructs parentage from level comparisons or `$path` text after `BindHierarchy`" — the current implementation already violates the letter of that invariant, in the one place where it matters.

This is small in code but significant in kind: it is exactly the "duplicated decision" that the architecture was supposed to eliminate. **Smallest correction:** implement the node-scoped matching from the start and assert that it reproduces the flat behaviour byte-for-byte (identical block order, identical warning text, order, location and data path). The cost is one rewrite of ~8 lines and one equivalence test; the benefit is that M3.4 does not have to carry two matching rules and a mental note about which one applies where.

### 3.4 MEDIUM — Credible risk: the merge-owner model cannot express "one owner, several blocks"

`XlsxAxisSpan` (`XlsxWorkbookPlan.cs:8`) carries only `VerticalInstancePath? Vertical` — **not** which scheduled block it came from. Once a parent has both a header block and a footer block (the normal nested case), the same owner path appears on two spans. The proposal acknowledges the immediate consequence (`§5.4`: `ProjectMerges` will enumerate the same owner twice and can falsely report overlapping merge projections). I verified that this is the correct reading of the code: `Merges.cs:85-86` takes every span's `Vertical` without dedup, and `Merges.cs:95-97` raises "Remove overlapping merge projections" on any rectangle overlap.

I also checked whether the row-image filter saves this. It does, for the *specific* case of a merge authored wholly inside one block: `AxisImage` (`Merges.cs:32-42`) filters spans by source-row overlap, so a merge in the header row does not pick up the footer span. So **dedup by owner path is sufficient** — but only if the dedup is implemented with *value* equality (see §3.1). With reference equality, `Distinct()` will not dedup two independently reconstructed equal paths, and the false "overlapping merge projections" error will appear.

The deeper point: merges and formulas scoped to a group need "the descendants of this owner", and the current predicates are exact-owner equality (`Owners.cs:85`, `Owners.cs:109`, `Owners.cs:115` all use `LogicalRows.SequenceEqual`). M3.7 will have to extend that to a descendant relation. The path shape supports it; the predicates do not yet.

**Recommendation:** decide now, explicitly, whether a merge/formula is owned by *(owner)* or *(owner, block)*. The code supports *(owner)* today and it is sufficient for M3.4 if merges stay inside one block. Write that decision into the plan as an invariant so M3.7 is not blocked by an unstated assumption. Optional low-cost insurance: add the originating block (or its identity) to `XlsxAxisSpan` — one field — so spans stop relying on source-row-range reverse lookup.

### 3.5 MEDIUM — Structural oddity, currently masked: the data hierarchy exists only when the *template* is nested

`XlsxTemplatePlanner.Rows.cs:57`:

```csharp
rows.Add(region.Id, region.Root.Children.IsEmpty ? grouped : BindHierarchy(grouped));
```

Whether a logical row has ancestor links is a function of the **template's** vertical nesting depth, not of the data. Two consequences:

- `{../Name}` means "parent record" in a nested template and "document root" in a flat one, for the same input. (Pinned deliberately by `Feature012BindingTests.Hierarchy.cs:167`; the proposal defers it to M3.5/R1.)
- Every remaining confusing case — R1, R2, and the scope-depth divergence in §3.2 — traces back to this one coupling.

The proposal is right that this must not change during M3.3 (it would break accepted M2 behaviour). But it should be **recorded as architectural debt with a named owner**, not left implicit, because M3.5 will otherwise have to rediscover it. **Recommendation:** one sentence in the plan's M3.5 row stating that M3.5's recovery policy must decide whether ancestor links are bound unconditionally, and that the flat/nested divergence in `{../}` semantics ends there.

### 3.6 MEDIUM — Contract risk: the moved M3.3 criterion is not restated in M3.4

The amended plan's M3.3 row still says *"Retain the existing two-row child repeated per item, own styles/merged names and no extra matrix record per physical row"*, while §5.3.2 permits deferring that proof, and **M3.4's row does not restate it**. As written, the criterion can be dropped without failing either checkpoint. Two of its clauses are load-bearing product guarantees, not implementation details: "no extra matrix record per physical row" is a statement about the *input data contract* (physical repetition must not multiply matrix arity), and "own styles/merged names" is about authored-content preservation.

Separately, note that the clause is currently **half-true**: the *flat* two-row repetition is already implemented and tested (`Feature012TwoAxisTests.cs:154`, plus the new `Multirow_Flat_output_preserves_…` test), but the *child* (nested) part is not — the new `Multirow_Parsed_nested_child_…` test asserts the bound hierarchy and then asserts that `ScheduleSheet` **throws** "Nested block scheduling requires M3.4 traversal". So M3.3 as currently written would be accepted with a schedule that can only ever produce `Body` blocks.

**Recommendation (documentation only, cheap):** in the M3.3 row, split the already-met flat property from the unmet nested property; in the M3.4 row, restate the nested proof verbatim as an M3.4 precondition including the "no extra matrix record per physical row" data-contract clause. Moving the proof is defensible; losing it is not.

### 3.7 LOW — Confirmed: `Region.Root.Body!` is an unchecked invariant, and `Role` has one producible value

`Scheduling.cs:27` uses `region.Root.Body!`. I traced the parser to check whether `Body` can be null for a flat region: it cannot — a region with a single level group always ends in the `first == last` break at `Vertical.cs:59-63` with `Body` set, and two same-level blocks in one region must be consecutive and therefore coalesced at `Vertical.cs:38-39`. So the `!` is currently safe.

But the invariant is implicit and undocumented, and it **does not survive into M3.4**: the source tree legitimately contains nodes with `Header` + `Footer` and **no** `Body` (e.g. source levels `10,40,10` parses to a level-10 node with `Header` and `Footer` and a level-40 child — `Vertical.cs:64-68`). `Role` currently has exactly one producible value. Both are fine *now*; both need to be honest in M3.4. **Recommendation:** replace `!` with an explicit invariant check that produces a coded diagnostic, and either produce `Header`/`Footer` roles in M3.4 or state in the plan that `Role` is currently single-valued.

### 3.8 LOW — Confirmed: an internal contract violation is reported as a corrupt package

`ExpansionGeometry` throws `ArgumentException` when the schedule does not match the sheet's regions (`Expansion.cs:11-12`). The handler's catch filter (`XlsxTemplateHandler.cs:76`) maps `ArgumentException` to `template.xlsx_invalid_package` — "The XLSX package is inconsistent." An internal planner/geometry disagreement is not a corrupt package, and this message would send the user (or support) hunting for a template-repair problem that does not exist. Low severity (it should never fire), but the mapping is wrong.

### 3.9 LOW — Confirmed: the nested guard throws an uncoded exception

`Scheduling.cs:16` throws a raw `NotSupportedException`. It is unreachable in production today (the public preflight guard at `Reports.cs:23-28` rejects nested templates before binding), and the new test asserts on it. But `NotSupportedException` is **not** in the handler's catch filter, so if it ever became reachable it would escape as an unhandled exception → HTTP 500 rather than a coded 422. **Recommendation:** use the existing `template.xlsx_unsupported_feature` diagnostic path (as `Reports.cs:25` does) so a future reachability slip degrades to an actionable error.

### 3.10 LOW / pre-existing, not B's fault: the cell-emission scan

`XlsxGeometryPlanner.Cells.cs:35-37` scans **all** source rows for **every** row span. With nesting, the number of spans grows with the number of logical items, so the cost becomes O(items × source rows). B neither causes nor fixes this; the assessment notes it (§7.8) without acting. For the product's current scale this is acceptable. Flagging it so it is not mistaken for a B regression, and so M6's capacity measurement has a known hotspot.

### 3.11 Naming / framing (preference, not a defect)

The proposal calls this a "logical schedule". It is not a symbolic or logical plan — it is a concrete, ordered **instance list** with physical row identities. "Vertical block instance list" would describe it accurately. The name matters slightly because "logical schedule" invites the reading that it is a general intermediate representation, which it is not, and which the proposal explicitly disclaims. Low value, no code impact.

---

## 4. Comparison of A, B and C (my own assessment)

| Criterion | A — extend geometry directly | **B — bounded vertical instance list (adopted)** | C — wide records + rewritten template |
| --- | --- | --- | --- |
| Where the hard traversal lives | Inside `XlsxGeometryPlanner` (already 8 partial files: cells, expansion, merges, owners, reports, static, core) | Inside `XlsxTemplatePlanner` (semantics owner), result crosses a boundary | Requires **both** a scheduler *and* a preprocessing layer |
| Testability of M3.4's ordering | Reachable only via emitted spans/placements, or by exposing an internal helper | Direct: assert `(Source, Role, Owner)` sequences with no coordinates | Hard: ordering still needs a scheduler; preprocessing adds a second layer to test |
| Separation vs spec §9.1 | Logical decisions and physical placement co-located | Logical ordering and physical placement separated across services | Adds a stage the spec does not ask for |
| Marginal code over "done properly" | 0 | ≈43 lines + 3 type names | Wide schema + synthetic fields + expanded template + provenance for two coordinate systems |
| Cost of being wrong | Low: everything lives in one place, revert is a local edit | Low now, rising: once M3.4's traversal feeds geometry, reverting means rewriting it into geometry (the proposal's §10.6 says this honestly) | Very high |
| Performance | No schedule list; traversal may run twice | +O(I) entries; **no measured difference** (correctly disclaimed) | +output-sized structures; avoidable copies |
| Future horizontal work | Would need a second traversal if horizontal scheduling is ever needed | Same — B is deliberately vertical-only | Same, plus the expanded template would have to be rebuilt |
| Proportionality for this product | Fine | Fine | Over-built: Feature 012 needs a *report generator*, not a spreadsheet engine |

**My conclusions differ from the documents' in three places:**

1. **A is more nearly "A done properly" than the assessment suggests.** The old design's fatal gap was not mixed responsibilities; it was the *absence of a vertical instance-identity concept*. Any correct A must introduce one. So the real comparison is "A with an owner identity" vs "B with an owner identity", and B's delta is ~43 lines plus a named boundary. B is not competing against a naive A; it is competing against a correct A that costs roughly the same.
2. **B is better spec-aligned than A, and the assessment under-uses this.** Spec §9.1 mandates conceptual separation of "expansion planning" from other stages. B places the logical plan in the semantics owner and the coordinates in the geometry owner. A does the opposite of what §9.1 asks for, and the assessment's own §5.1.3 concedes it ("geometry must also answer logical questions: which template level matches a data row, what group owns a footer…").
3. **C is strictly worse than B, and the decisive reason is not memory.** C *contains* B: flattening the template does not determine footer ordering, so C still needs the scheduler, plus a wide value table, plus synthetic field names, plus a second coordinate system (original vs expanded template) that every formula, merge and print setting must be transformed through *twice*. The assessment is right that "once source/node/parent identities are retained and template mutation is avoided, that view approaches B" — that is effectively a proof that C collapses into B plus overhead. **C should be rejected permanently** for this product.

**Proportionality verdict:** B is proportionate. The product is a template-driven report generator with a deliberately narrow envelope (one horizontal region per sheet, Cartesian product matrix, no evaluator, no repair framework). B adds three small internal records and one stage. That is not spreadsheet-engine over-engineering. The risk in B is not its size; it is that the *identity semantics* (§3.1) and the *compatibility rule* (§3.2) are currently under-specified, and M3.4 is about to be written against them.

---

## 5. Recommendations

### Keep

1. **Keep the schedule as a named model crossing the planner→geometry boundary.** It earns its cost: one owner of vertical selection and omission, one inspectable artifact for the hardest traversal, and alignment with spec §9.1.
2. **Keep the vertical-only scope.** Horizontal scheduling remains deferred; the documents are right that adding it now would be speculation against explicitly deferred requirements (spec §11.1-11.3).
3. **Keep the "A fallback" clause** in plan §4.2.9 and the proposal's §10.4 acceptance condition ("geometry no longer accesses bound vertical rows or chooses blocks; BindInstances uses the carried current owner"). That condition is exactly the right tripwire; keep it as an explicit M3.3 exit criterion.
4. **Keep the single omission-warning emission site** and its exact message/order/location/data-path contract.
5. **Keep the public nested-generation guard** unchanged until a verified subset exists.

### Change now (before M3.4 is written)

6. **Give `VerticalInstancePath` value equality over `LogicalRows`** (§3.1). ≈10 lines in `XlsxTemplatePlan.cs` + one test. *Highest-value single change in this review.* It converts the merge-dedup rule, the owner-removal rule and the placement-key rule from "correct only if callers share objects" into "correct by construction".
7. **Replace the flat level-comparison matching with the node-scoped rule now**, and assert flat equivalence (identical block order, warning text/order/location/data path) (§3.3). ≈8 lines + one equivalence test. Removes the duplicated decision the architecture exists to eliminate.
8. **Generalize the R2 check into the ancestor-chain compatibility invariant** (§3.2), while keeping the proposal's actionable message for the documented case. This is what prevents a silent root fallback for `../..` when the data's hierarchy is shallower than the template's. This must be decided *before* the traversal is written, because it determines what "matching" means.
9. **Restate the moved M3.3 nested criterion in M3.4's row** (§3.6), splitting the already-met flat property from the unmet nested property, and keep the "no extra matrix record per physical row" data-contract clause. Documentation only.

### Simplify

10. **Replace `region.Root.Body!` with an explicit invariant** producing a coded diagnostic (§3.7). Also decide whether `Role` stays single-valued until M3.4.
11. **Replace the raw `NotSupportedException` in `ScheduleSheet` with the existing `template.xlsx_unsupported_feature` diagnostic path** (§3.9) so a future reachability slip yields a 422, not a 500.
12. **Do not let internal contract violations surface as `template.xlsx_invalid_package`** (§3.8). Either narrow the handler's catch filter or use a distinct internal-error code.

### Defer (explicitly, with a named owner)

13. **Merge/formula descendant scoping** — M3.7. Record the decision "a merge/formula is owned by *(owner)*, and must be authored within one block for M3.4" in the plan so M3.7 is not blocked by an unstated assumption.
14. **The flat-vs-nested data-hierarchy coupling** (`Rows.cs:57`) — M3.5. Record it as architectural debt with a named owner, because R1, R2 and the scope-depth divergence all trace to it.
15. **The emission scan hotspot** (`Cells.cs:35-37`) — M6 capacity measurement. Not B's regression; do not fix speculatively.
16. **Horizontal scheduling, wide records, synthetic field names** — permanently rejected for this product unless the Owner approves a scope change against the spec's explicit deferrals.

### Do not do

17. **Do not widen owner matching to prefixes** to shortcut M3.7 (§5.7 of the proposal). The path shape allows it later; doing it now would silently change what a formula means.
18. **Do not add eager field values, per-cell ancestors, or a second horizontal planner** to the schedule. The proposal's §10.4 acceptance condition correctly forbids these; keep them as rejection criteria in M3.3/M3.4 review.

---

## 6. Verdict

# **Modify B** — keep the adopted approach, fix three specific things before M3.4.

**Why not "Keep B":** the instance-identity model is defective (§3.1, empirically confirmed), the vertical-matching rule exists in two incompatible forms (§3.3), and the approved compatibility policy is too narrow to prevent a silent wrong-value failure mode (§3.2). None of these require abandoning the architecture; all three are small and local.

**Why not "Return to A":** the old design's real deficiency was the absence of a vertical instance-identity concept, which any correct A must introduce anyway. B's marginal cost over a correct A is ≈43 lines and three internal type names, and B is *better* aligned with the spec's mandated separation (§9.1) than A, which would co-locate logical ordering with physical placement inside an already eight-partial geometry planner. The traversal that makes M3.4 hard is identical under both options; only its owner and its inspectability differ.

**Why not C:** C contains B and adds a second coordinate system, a synthetic field schema and an output-sized template copy, for a product that needs none of them.

**Why B is proportionate:** 43 lines and 3 records against a ~2,350-line engine that already renders Excel output correctly for static, flat-vertical and nested-horizontal reports. B is not a general-purpose spreadsheet engine; it is a seam in a report generator. The proportionality risk in B is not size — it is that the *semantics* of identity and compatibility were left implicit when the types were introduced. Fixing them now costs a fraction of what fixing them after M3.4 costs.

---

## 7. The top three issues to resolve before or during M3.4

1. **Instance identity must be value-based, not reference-based.** `VerticalInstancePath` embeds `ImmutableArray<int>` in a `record`; equality is therefore reference equality on the underlying array (verified). M3.4 will construct owner paths in several new places, and the merge-owner dedup, the content-owner removal (`Owners.cs:60`) and the placement dictionary key (`Instances.cs:58`) all depend on this equality being correct. *Fix:* value equality over the sequence in `XlsxTemplatePlan.cs` + a test that two independently constructed equal paths are equal and hash equal. *Cost:* ≈10 lines. **Must be done before M3.4 traversal code is written.**

2. **The hierarchy-compatibility rule must be the general ancestor-chain invariant, and it must guarantee that the runtime scope depth can never be shorter than the statically validated depth.** The approved R2 rule covers one shape; the general rule — *data item depth == source node depth and ancestor levels match* — also closes the case where the data skips a level, which currently makes `{../..}` silently resolve to the document root (`DataResolver.cs:37`) with no diagnostic. *Fix:* implement the general check, keep the proposal's actionable message for the documented case, and assert the depth agreement as part of it. *Cost:* small; it *subsumes* the narrow case. **Must be decided before the traversal is written**, because it defines what "matching" means.

3. **There must be exactly one vertical-matching rule, and it must be node-scoped.** The current `row.Level != region.Level` test (`Scheduling.cs:20`) is a second rule that only coincides with the node-scoped rule for flat regions; leaving it in place means M3.4 carries two matching rules and a mental note about which applies. *Fix:* implement node-scoped matching now and prove flat equivalence (identical block order and identical warning text, order, location and data path). *Cost:* ≈8 lines + one equivalence test.

**And, additionally, before M3.4 is accepted (documentation, not code):** the moved M3.3 nested criterion — "two-row child repeated per item, own styles/merged names and **no extra matrix record per physical row**" — must be restated verbatim in M3.4's checkpoint row, and M3.3's row must distinguish the already-met flat property from the still-unmet nested one. Moving the proof is sound; the current amendment allows it to be lost.

**Two smaller items to fold into the same M3.4 change:** replace `region.Root.Body!` with an explicit invariant check, and replace the raw `NotSupportedException` at `Scheduling.cs:16` with the existing `template.xlsx_unsupported_feature` diagnostic path so a future reachability slip degrades to a 422 rather than an unhandled 500.

---

*Report ends. No files, tests, documentation, Git index or working tree were modified. No implementation of M3.4 was performed. Awaiting Owner review.*