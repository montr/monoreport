# Feature 014 — XLSX Template Compatibility

**Status:** Draft for Product Owner review, based on the approved scope and decisions. Implementation is not authorized by this document.

## 1. Goal and scope

1. Users can edit supported XLSX templates in Microsoft Excel using ordinary formatting and content changes, supported formulas and supported worksheet editing operations within the compatibility boundaries in section 2. Saving those edits must not cause unnecessary analysis or generation failures.
2. Include harmless Excel-authored metadata and worksheet selection normalization, optional PrintArea for worksheets without repetition controls, and rich-text substitution preserving ordinary character formatting.
3. Preserve generation correctness, native value types, formulas, formatting and workbook integrity through the existing Core, CLI and HTTP interfaces. Existing analysis must accept the same supported source constructs; its data model and Sample JSON behavior are outside this feature.
4. This scope supersedes the broader discovery brief and the relevant earlier restrictions in Feature 012. The [architecture research](architecture-research.md) and [rich-text research](rich-text-research.md) provide rationale; their additional proposals are not requirements unless included here.

## 2. Required behavior

### 2.1. Excel-authored metadata and worksheet views

1. Opening and saving an otherwise supported template in Excel must not make it unusable solely because Excel adds harmless authoring metadata. Support the observed save-path metadata inside AlternateContent and transient workbook revision pointers; omit that disposable metadata, including the embedded local save path, from generated reports.
2. Valid selections, active cells and viewport positions must not prevent analysis or generation because they extend outside PrintArea, lie wholly outside retained content, contain multiple ranges, or refer to cells that repeat. Preserve representable view state; otherwise normalize it to valid output view state without changing report content or formulas.
3. Retain ordinary workbook and worksheet settings where compatible with generation. Selection normalization does not authorize silently removing frozen/split panes, print titles or other settings with unproven transformations.
4. Unknown or meaningful alternate content, extensions and relationships must not be silently discarded or accepted solely as harmless metadata. Unsupported constructs that cannot be preserved safely receive actionable diagnostics identifying the actual workbook part and available worksheet/cell context. Malformed packages remain errors.

### 2.2. PrintArea and worksheet retention

1. A worksheet without repetition controls may omit PrintArea. Preserve its content and coordinates, including scalar substitutions, empty and formatting-only worksheets, ordinary reference sheets, authored row/column sizes, sheet name, order and visibility. Do not drop such sheets, infer a crop boundary or add a synthetic PrintArea.
2. A worksheet using repetition controls still requires one finite rectangular PrintArea. Missing or malformed boundaries and malformed/misplaced controls receive actionable errors; they must not cause the sheet to be treated as an ordinary uncontrolled sheet.
3. An explicitly configured PrintArea retains the existing physical crop behavior: remove external controls and disposable content, express the retained report relative to A1, and transform supported references and settings. Invalid, duplicate or disjoint explicit areas remain unsupported, including on sheets without repetition.
4. References and supported defined names must remain correct across all retained sheets, including references from an unchanged sheet into an expanded report and references to blank cells beyond the last stored cell. Omitting PrintArea must not truncate the worksheet's valid reference domain.

### 2.3. Rich-text substitution and value types

1. Preserve whole-cell formatting during substitution, including fonts, colors, fills, borders, alignment and number formats. Partial-character formatting alone must not cause analysis or generation to fail.
2. Support expressions and escaped braces spanning rich-text runs in both shared and inline strings. For textual results, preserve the character formatting of unchanged surrounding text. Inserted text inherits the formatting at the expression's opening brace; an escaped brace inherits the first brace's formatting. Different formatting within the replaced expression does not split the replacement into proportional styles.
3. Apply these rules consistently to static cells and repeated cells, including headers and footers. Preserve whitespace, line breaks, Unicode and SpreadsheetML-escaped text. Shared source strings must produce independent values and formatting for each output occurrence. Unchanged rich literals retain their formatting.
4. Keep existing scalar semantics: a whole-cell unformatted expression returns a native number, Boolean or string according to its JSON value; interpolation and explicit formatters return text. Native number, Boolean and empty results retain whole-cell formatting; character formatting stored only in the placeholder need not survive, and is not promoted into new cell styles. Existing precision, missing/null and invalid-data rules continue to apply.
5. Inserted text remains inert, including leading equals signs and expression-like text. Rich-text support does not introduce expressions inside Excel formulas or formatted runs in formula results.
6. Preserve phonetic annotations on unchanged literal text with valid positions after substitution. Remove annotations associated with replaced expression text; do not invent phonetic readings for inserted data.

### 2.4. Correctness and compatibility boundaries

1. Preserve supported formula semantics and references through repetition and crop, including cross-sheet references and supported names. Retain existing recalculation behavior; MonoReport does not evaluate formulas or populate calculated caches.
2. Preserve supported merges, cell types, styles, dimensions and print settings. Reject ambiguous or unsafe transformations with actionable diagnostics and no partial generated report. Analysis and generation must agree on data-independent source compatibility; successful analysis does not guarantee data-dependent generation success.
3. Keep caller-owned template bytes unchanged and produce a valid workbook that Excel can open without repair. Existing DOCX behavior and public transport contracts remain compatible.
4. This feature does not promise all Excel/OOXML capabilities, broader formula/reference syntax, general reconstruction of panes or print settings, byte-identical XML, or exact Excel pagination and browser-preview fidelity. Existing compatibility boundaries remain except where explicitly amended above.

## 3. Acceptance criteria

1. **Excel save compatibility:** an existing supported template, including the employees sample, remains analyzable and generates the expected report after an unchanged Excel save and representative ordinary cell-formatting/content edits. Verify the end-to-end workflow: edit/save template in Excel → Analyze → Generate → reopen generated workbook in Excel. Harmless save-path and revision-pointer metadata is absent from output; report data and formatting remain correct.
2. **Selection normalization:** full-row, multiple-range, outside-crop and repeated-cell selections do not block a valid report. Output selection, active-cell and viewport state is consistent and does not change report values, geometry or formulas.
3. **Optional PrintArea:** a workbook mixing a cropped repeating report with no-PrintArea empty, formatting-only, scalar and reference sheets preserves those sheets' identity, content and coordinates. Cross-sheet references and references to implicit blank cells remain correct; no synthetic print area or content is introduced.
4. **Boundary errors:** repetition without a finite PrintArea, invalid explicit areas, malformed controls and unsafe retained references still fail with useful source diagnostics and no partial report. Meaningful unsupported alternate content is not silently deleted.
5. **Rich-text fidelity:** expressions within and across differently styled runs substitute correctly, preserving surrounding formatting and expression-start styling. Shared/inline strings, escaped braces, whitespace, changed-length replacements and repeated uses retain correct text and independent formatting. Unaffected phonetic annotations remain correctly associated.
6. **Native values and formulas:** rich source formatting does not turn native numeric/Boolean values into text or weaken precision checks. String identifiers, formatted values, null/missing behavior and inert inserted text follow existing semantics. Supported dependent formulas retain their meaning and recalculation settings.
7. **Workbook integrity and regressions:** representative Excel-authored outputs reopen in Excel without repair and retain the required visual formatting. Package validity is accompanied by value, type, reference and formatting checks; existing XLSX generation and DOCX workflows remain compatible.

## 4. Exclusions

1. Analyzer model/schema changes, requirement/occurrence enrichment, vertical-level metadata, Generate JSON correctness improvements and Sample JSON ownership or new generation services belong to Feature 015. Feature 013 finding S3-02 is not resolved by Feature 014.
2. New UI elements, preview changes, endpoints, CLI commands, formula evaluation, general OOXML support, rich-text-to-native-cell style synthesis and phonetic-reading generation are outside scope.
