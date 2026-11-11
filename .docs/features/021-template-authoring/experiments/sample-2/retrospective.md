# Sample 2 authoring retrospective

## 1. Conclusion and evidence limits

1. Externally measured elapsed time was **608 seconds**, versus **502 seconds** for Sample 1: **106 seconds longer, or 21.1%**. Both templates worked on the first attempt. Nested repetition increased inspection and semantic verification work, but the evidence does not attribute the additional 106 seconds to nesting.
2. Evidence: session/tool history, artifact timestamps, [findings](findings.md), [verification](evidence/verification.json), [conversion timing](evidence/preview/timing.json), and [Sample 1 retrospective](../sample-1/retrospective.md). No checks were rerun. Billed-token totals and a complete timestamped model trace are unavailable.
3. Approximately **476.5 seconds elapsed between template creation and final script cleanup**. This includes validation, reasoning, user steering, evidence handling, reporting, and cleanup; it is not a measurement of verification CPU time. It nevertheless makes post-authoring work the clearest optimization target.

## 2. Reconstructed timeline

Times are **12 November 2026, Europe/Moscow**. File timestamps measure wall intervals, not exclusive active work. HTTP dates corroborate ordering, with a small local-clock offset. Rows overlap: **do not sum them**. Exact session start/end and isolated documentation/authoring durations were not recorded.

| Stage | Available anchors | Measured duration or interval | Interpretation |
|---|---|---|---|
| Rules, documentation, API discovery | Before first source inventory at 00:56:48.787; saved OpenAPI/formatters at 00:56:51 | Not isolated | Read rules, all three Feature Skills, guide/contract, Feature documents, API source; permission retry for localhost. |
| Document inspection and template authoring | Source inventory 00:56:48.787 → template 00:58:12.482 | **83.7 s wall interval** | Includes further structure/property inspection, architecture rereading, inference, script composition, extraction, and editing. Execution of the finished authoring script itself was approximately 0.2 s in the tool trace. |
| Analyze and five REST generations | Template 00:58:12.482 → final render headers 00:59:58.507 | **106.0 s wall interval** | Analyze saved at 00:58:28.010; original render headers at 00:58:37.074; original attachment extracted at 00:59:29.146. Much of the interval was MIME/dataset scripting, not API execution. Four later renders completed within approximately 0.1 s between local header-file timestamps. |
| Generated values, preservation, OOXML verification | Final render headers → SDK evidence complete at 01:02:46.966 | **168.5 s wall interval** | Included writing the verifier, preparing/extracting artifacts, SDK setup and other overlapping work. Build log measures **21.20 s**; SDK validation output was written over approximately **0.9 s**. No assertion or validation failure occurred. |
| LibreOffice conversions, overlapping earlier stages | Source PDF 00:58:50; original PDF 01:00:01; variant PDFs 01:00:41 | Source command approximately **1.0 s**, original **0.6 s**, four-variant batch approximately **1.1 s** | Already installed. Initial failed attempt approximately **0.19 s**; first-attempt log to source PDF **27.4 s**, including retry and intervening work, not isolated conversion time. |
| Existing-PDF inspection; no visual review | Swift script 01:03:17.976 → PDF metadata/text evidence 01:03:52.959 | **35.0 s wall interval** | Includes Swift startup/compilation and PDFKit extraction; these costs are not separately timed. No new conversion or visual rendering after user steering. |
| Evidence and findings | Package hashes 01:03:57.827 → findings 01:05:51.713 | **113.9 s wall interval** | Included byte comparison, reasoning, and composing a 17,285-byte report. Final cleanup followed at 01:06:08.956. |

## 3. What was actually harder than Sample 1

1. Sample 1 had one table, three material records, one collection, and five parts. Sample 2 had **20 tables, 152 lines, two collection levels, and 23 parts**. Order ownership and subtotal placement required additional reasoning. Uniform formatting reduced those tables to **one prototype**, avoiding 20 independent edits.
2. Report-header extraction added organization, warehouse, period, currency, and summaries; order headers added number/date/customer/count/total. These are body paragraphs. Six physical Word header/footer parts remained unchanged; no header-part editing or relationship repair was needed.
3. Five renders instead of three checked mixed zero/one/24 inner records, outer expansion, all-inner-empty, and outer-empty cases alongside original data. Membership, summaries, root currency, and all cells needed verification. The separate all-inner-empty case partly overlaps mixed-empty coverage and could be omitted when it adds no distinct invariant.
4. Authoring grew from **100 to 142 lines**, verification from **149 to 180**, plus an **81-line** case/transport script. These are work-product sizes, not time measurements. Additional loop/header effort cannot be assigned a defensible duration.

## 4. Failures, repeated work, and token cost

1. **No template defects:** Analyze, five renders, and seven SDK validations passed without repair. Sample 1's erroneous preservation assertion, validator build retry, and preview-background failure did not recur.
2. **Environment:** sandboxed localhost access and LibreOffice launch failed initially; permitted escalation resolved both. Missing optional Python packages and failed path/glob probes added discovery branches without blocking standard-library editing. Isolated retry costs are unavailable beyond LibreOffice's approximately 0.19-second failed command; its 27.4-second setup interval includes intervening work.
3. **Oversized reads:** tool-reported output estimates were **18,193 tokens before truncation** for broad Feature/API reading; **7,717** for architecture rereading plus inspection; **6,457** for OOXML/header/footer dumps, mostly repeated namespaces; and **4,320** for the full source inventory. These are neither billed-token totals nor mutually additive measurements.
4. All three Feature Skills were read unnecessarily for this experiment. Architecture was reread, and source inspection was repeated for structure, properties, and uniformity. One bounded inventory with selected prototype expansion could answer those questions without repeated large outputs.
5. OpenAPI was saved once, improving on Sample 1's duplicate retrieval. **Nine successful HTTP calls** covered root, OpenAPI, formatters, Analyze, and five renders; one root request failed. No Analyze/Render was repeated. Reconstructing local helpers mattered more than request latency.
6. Ad hoc preservation comparisons, the full verifier, separate hashes, and a newline comparison could have emitted evidence together. Custom relationship checks overlap SDK responsibilities. The 17 KB findings report repeated preservation conclusions and limitations instead of linking evidence once.

## 5. LibreOffice value

1. **No visually discovered defect:** no screenshots, manual visual review, or Word opening occurred. PDFKit found identical source/original pagination and page text. This adds renderer-specific loadability/pagination evidence, not visual-fidelity proof.
2. Six conversions before structural verification had limited incremental value given exact original-data XML reconstruction and preserved supporting parts. PDF inspection added a **35-second** wall interval plus incompletely isolated setup overhead; the approximately one-second exports themselves were not the main bottleneck.
3. Make visual checks **optional and risk-based**, after structural checks, targeting a specific concern such as table splitting or changed geometry. Conversion alone is not visual inspection. No unresolved layout issue justified more rendering here.

## 6. Prioritized improvements

The savings below are **planning hypotheses**, not measured causal effects, guarantees, or additive estimates. They exclude the fixed cost of independently understanding the report.

| Priority | Destination | Concrete change | Estimated opportunity, hypothesis |
|---|---|---|---|
| 1 | Skill-bundled verification scripts | Reuse a tested SDK validator invocation and binary-safe package inventory/diff; emit hashes, changed-part subsets, namespace/section preservation, and validation results once. Keep report-specific expected values, arithmetic, and cardinalities explicit. Avoid recreating a general OPC validator. | **30–90 s** from reduced code composition/setup; the 21.20-second build is measured, but its full removal depends on distribution/cache strategy. |
| 2 | `SKILL.md` and agent-oriented references | Read applicable sections once; use bounded discovery summaries and one compact final evidence report. Explain paragraph loops enclosing complete tables, nested row scopes, summary ownership, and report-header versus Word-header distinctions. | **30–90 s**, primarily less reading/writing; no billed-token saving can be quantified. |
| 3 | `SKILL.md` | Gate optional visual work on a named unresolved layout risk, after structural verification. Reuse installed tools without renderer setup by default; select one relevant case rather than export every dataset. | **30–60 s** when visual work is unnecessary; the 35-second PDF-inspection interval supports this order of magnitude, not a universal saving. |
| 4 | Skill-bundled inspection script | Produce ordered blocks/stories, table counts/dimensions, formatting-equivalence groups, run boundaries, section/reference hazards, and selected text locally. Print a bounded summary and expand only chosen prototypes. | **15–40 s**; grouping 20 equivalent tables and suppressing repeated namespace markup is directly relevant. Business interpretation remains the agent's job. |
| 5 | Skill-bundled REST script | Discover/save metadata once, upload files, request multipart diagnostics, check `valid`, validate/extract attachments, and link each output to its actual dataset. Return concise results. Do not require Content-Length when valid HTTP framing supplies another transfer method. | **15–40 s** from avoiding ad hoc curl/MIME code; server time was already small. |
| 6 | MCP/REST or engine | No new editing, inspection, batch-generation, aggregation, or rendering endpoint is justified by this run. Exact engine/docs identity remains unknown; this reinforces the existing provenance requirement rather than proving a new scope need. | **No demonstrated server-side optimization.** Preserve current contracts. |

## 7. New lessons and next experiments

1. Sample 1 already established transport/validator reuse, bounded reads, explicit summaries, and visual limits. Sample 2 newly demonstrates **one prototype for many uniformly formatted nested records**, separate inner/outer empty-case obligations, and useful root-scope testing through changed currency.
2. DOM editing preserved a 23-part package without rebuilding headers/footers. One harmless declaration-newline change complements Sample 1's zero-changed-parts result: preservation checks must allow both, while still verifying content and properties.
3. Timing now separates fast exports from slower setup/post-conversion tooling and exposes substantial helper/reporting work after successful authoring. Neither trial supports a DSL or dedicated editor.
4. Next time, record monotonic stage durations, overlap/permission waits, and available token metrics. Use compact inspection and authorized reusable helpers while retaining semantic/preservation checks. **502 seconds or less** is an evaluation target, not a promise or a controlled comparison between documents.
5. Test one genuinely harder structure, such as split-run fields or heterogeneous row formatting, rather than just more records. Define hierarchy/empty-case expectations first; render only for an unresolved layout question. These recommendations authorize no implementation or architecture change.
