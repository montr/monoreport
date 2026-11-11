# Feature 021 — AI Template Authoring Experiments

![Asya and Arch exploring MonoReport templates](asya-and-arch.png)

*Guess the Template from the Report.* Five independent Codex experiments explored whether a completed DOCX or XLSX report can be converted into a reusable MonoReport template, with representative JSON data and verified output, before building a dedicated authoring Skill.

## Experiment results

| Experiment | Input size | Source structure | Repetition modeled | First verified output | Total time | Tokens processed¹ | Outcome |
|---|---:|---|---|---:|---:|---:|---|
| [1 — DOCX estimate](experiments/sample-1/findings.md) | 2,741 B | 3 material rows | 1 vertical level | Not recorded | 8:22 | 1,484,153 | Passed |
| [2 — DOCX orders](experiments/sample-2/findings.md) | 60,146 B | 20 orders, 152 items | 2 nested levels | Not recorded | 10:08 | 1,645,810 | Passed |
| [3 — DOCX project costs](experiments/sample-3/findings.md) | 29,139 B | 3 projects; grouped cost rows and merged cells | 3 nested levels in one table | 2:22 | 6:02 | 1,153,463 | Passed |
| [4 — XLSX employees](experiments/sample-4/findings.md) | 12,135 B | 10 departments, 100 employees | 2 vertical levels | 2:37 | 5:12 | 1,042,040 | Passed² |
| [5 — XLSX tender comparison](experiments/sample-5/findings.md) | 9,479 B | Lots, items, suppliers, and criteria | Vertical lots/items; horizontal suppliers; **criteria left fixed** | 4:29 | 6:47 | 1,040,720 | **Partial** |

¹ Totals are cumulative input (including cached input) plus output (including reasoning), measured from the Codex JSONL sessions **up to the experiment's completion**. They exclude separate retrospectives and permission-review sessions. Elapsed times are the experiment measurements, not controlled performance benchmarks. A token count is not a direct cost estimate.

² The source XLSX already contained two Open XML SDK DPI diagnostics; the template and generated outputs introduced no new diagnostics.

## Token usage

| Experiment | Cached input | New input | Output (including reasoning) | Total |
|---|---:|---:|---:|---:|
| DOCX 1 | 1,401,088 | 63,901 | 19,164 | 1,484,153 |
| DOCX 2 | 1,560,064 | 60,586 | 25,160 | 1,645,810 |
| DOCX 3 | 1,072,896 | 65,087 | 15,480 | 1,153,463 |
| XLSX 4 | 987,904 | 41,785 | 12,351 | 1,042,040 |
| XLSX 5 | 968,064 | 54,928 | 17,728 | 1,040,720 |

## Findings

1. **DOCX feasibility:** Three documents were converted without template repairs, including nested repetition inside a single table, merged heading cells, and split-run title text. Successful structural and package checks do not establish perfect visual fidelity in every Word version.
2. **XLSX feasibility:** The employee workbook demonstrated vertical grouped repetition with preserved worksheet geometry, merges, print settings, and source-validation baseline.
3. **Critical limitation:** Sample 5 modeled the number of suppliers as variable but treated the comparison criteria as fixed attributes. All implemented test datasets passed, yet the result did not support a second independent horizontal dimension. **Technical validity is not semantic completeness.**
4. **Formula boundary:** In Sample 5, an empty lot caused `template.xlsx_unsafe_formula` because its subtotal referenced removed item cells. This is a limitation of that authored template, not a demonstrated universal MonoReport restriction.
5. **Efficiency:** Much of the measured work occurred after first generation: ad hoc REST/MIME handling, verification scripts, validator setup, documentation reads, and reporting. Reusable helpers are promising, but their speedup has not yet been measured.
6. **Evidence boundaries:** Analyzer/Render success, OOXML checks, structural preservation, and application-level visual/recalculation checks are different forms of evidence. Do not infer one from another.

## Implications for the planned Skill

1. Provide portable agent-oriented authoring guidance for DOCX and XLSX, plus general formatter rules. Discover the *actual* formatter catalog and version-specific engine contracts through the connected MonoReport API/MCP rather than maintaining a hard-coded formatter list.
2. Identify **every candidate independent data dimension before authoring**. Resolve material ambiguity using a concise user description or contrasting completed reports. Verify each intended dimension independently, especially combined vertical and horizontal repetition.
3. Accept one or several finished reports, optional JSON, and an optional natural-language description of dynamic fields, hierarchy, invariants, calculations, and empty-collection behavior.
4. Offer **Fast** authoring with essential checks and **Full** verification with stronger semantic, OOXML, and preservation checks; Full should also be usable on an existing Fast result. Exact defaults and check policies remain to be approved.
5. Prioritize bounded document inspection, reliable REST multipart handling, and reusable structural verification before creating any dedicated OOXML editor or authoring DSL.
6. Keep future Playground authoring through an agent runtime and external model-provider APIs (including OpenCode Inference) possible through the same portable tools, without making that integration part of this Feature.

## Evidence and artifacts

| Sample | Input | Template | Representative JSON | Modified result | Report | Retrospective |
|---|---|---|---|---|---|---|
| DOCX 1 | [DOCX](experiments/sample-1/input/skill-sample-1.docx) | [DOCX](experiments/sample-1/output/skill-sample-1.template.docx) | [JSON](experiments/sample-1/output/skill-sample-1.json) | [DOCX](experiments/sample-1/output/skill-sample-1.result.docx) | [Findings](experiments/sample-1/findings.md) | [Review](experiments/sample-1/retrospective.md) |
| DOCX 2 | [DOCX](experiments/sample-2/input/skill-sample-2.docx) | [DOCX](experiments/sample-2/output/skill-sample-2.template.docx) | [JSON](experiments/sample-2/output/skill-sample-2.json) | [DOCX](experiments/sample-2/output/skill-sample-2.result.docx) | [Findings](experiments/sample-2/findings.md) | [Review](experiments/sample-2/retrospective.md) |
| DOCX 3 | [DOCX](experiments/sample-3/input/skill-sample-3.docx) | [DOCX](experiments/sample-3/output/skill-sample-3.template.docx) | [JSON](experiments/sample-3/output/skill-sample-3.json) | [DOCX](experiments/sample-3/output/skill-sample-3.result.docx) | [Findings](experiments/sample-3/findings.md) | [Review](experiments/sample-3/retrospective.md) |
| XLSX 4 | [XLSX](experiments/sample-4/input/skill-sample-4.xlsx) | [XLSX](experiments/sample-4/output/skill-sample-4.template.xlsx) | [JSON](experiments/sample-4/output/skill-sample-4.json) | [XLSX](experiments/sample-4/output/skill-sample-4.result.xlsx) | [Findings](experiments/sample-4/findings.md) | [Review](experiments/sample-4/retrospective.md) |
| XLSX 5 | [XLSX](experiments/sample-5/input/skill-sample-5.xlsx) | [XLSX](experiments/sample-5/output/skill-sample-5.template.xlsx) | [JSON](experiments/sample-5/output/skill-sample-5.json) | [XLSX](experiments/sample-5/output/skill-sample-5.result.xlsx) | [Findings](experiments/sample-5/findings.md) | [Review](experiments/sample-5/retrospective.md) |

## Feature documentation

1. [Feature specification](feature-spec.md)
2. [Architecture](architecture.md)

The experiments are research evidence, not acceptance of all Feature 021 architecture choices. The Hybrid MCP/REST details, context binding, initial tool scope, and Fast/Full policy still require Product Owner decisions.
