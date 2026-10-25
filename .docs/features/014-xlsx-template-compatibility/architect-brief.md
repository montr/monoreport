# Feature 014 — XLSX Template Compatibility and Core Improvements

## 1. Product goal

Users should be able to edit XLSX report templates in Microsoft Excel using ordinary spreadsheet editing and formatting operations, save them, and continue using them in MonoReport without unnecessary restrictions imposed by the XLSX engine.

Support routine changes to fonts, colors, borders, cell contents, formulas, and worksheet structure where compatible with correct report generation. Empty worksheets should not be rejected solely because they are empty. This is not a requirement to support every Excel or OOXML feature.

Preserve generated report correctness, formulas, formatting, and document integrity. Identify genuinely necessary template restrictions rather than assuming current validation behavior is a product requirement.

## 2. Known problems and investigation topics

1. **Excel resave compatibility (Feature 013, S3-01).** A working XLSX template can become invalid after opening and saving it in Excel, even unchanged. Observed constructs include `mc:AlternateContent` with `x15ac:absPath`, `xr:revisionPtr`, and a worksheet selection `A1:XFD1` beyond `PrintArea`. Investigate the underlying validation approach, not just exceptions for these examples.
2. **Worksheet/PrintArea validation.** Adding an empty worksheet caused `template.xlsx_invalid_print_area`. Determine when a finite PrintArea is actually necessary, including empty and non-template worksheets, and how PrintArea interacts with cropping and selections.
3. **Analyzer model.** Evaluate the purpose and consumers of `requirements` (aggregated input-data tree) and `occurrences` (individual expressions and source locations). Determine whether both representations are warranted.
4. **Generate JSON and levels (Feature 013, S3-02).** `Samples/xlsx-employees.template.xlsx` contains `$l(10).Employees` and `$l(40).Employees`, but generated sample JSON collapses these into one `Employees` item without `$level`. Investigate a correct representation together with the Analyzer model.
5. **Generate JSON ownership.** Assess whether sample JSON construction belongs in Web UI, API, or Core, considering actual consumers and avoiding unnecessary abstractions.

## 3. Research outcome

1. Inspect the current implementation, relevant architecture decisions, and Feature 012/013 documentation and review findings; verify reported behavior against the repository.
2. Distinguish actual defects, deliberate safety restrictions, and incidental implementation limitations.
3. Compare reasonable approaches and recommend the simplest coherent solution, with key risks and compatibility implications.
4. Identify decisions that require Product Owner approval before writing the Feature Spec.

This is discovery only. Do not implement changes or treat architectural hypotheses as approved requirements.

## 4. Relevant repository references

1. `.docs/features/013-xlsx-web-ui/feature-spec.md`
2. `.docs/features/013-xlsx-web-ui/implementation-plan.md`
3. `.docs/features/013-xlsx-web-ui/implementation-review.md` (especially S3-01 and S3-02)
4. `Samples/xlsx-employees.template.xlsx` and `Samples/xlsx-employees.json`
5. `MonoReport/Services/Xlsx/`
6. Analyzer, `TemplateRequirement`, `TemplateRequirementBuilder`, Web UI `sampleJson.ts`

Read project `AGENTS.md` and applicable Skills. Feature 014 does not automatically include the separate future-improvements backlog.
