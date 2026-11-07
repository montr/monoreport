# XLSX Template Compatibility

MonoReport generates Excel workbooks from XLSX templates and JSON data. It supports selected spreadsheet structures and safe row/column expansion, and rejects unsupported features rather than risk corrupt or misleading output. For authoring examples, see the [XLSX Template Guide](xlsx-template-guide.md).

MonoReport checks the structures needed for safe template processing; it is not a general Word, Excel, or OOXML validator. A source document may be processed even if an Office application cannot open it, provided MonoReport can read the required content and no documented compatibility restriction prevents generation.

A feature can be **Supported**, **Preserved unchanged**, **Removed**, or **Rejected**, depending on whether a worksheet uses repetition or PrintArea cropping.

## Compatibility

| Excel feature | MonoReport behavior |
|---|---|
| Workbook formats | **Supported:** `.xlsx`. **Rejected:** `.xltx`, `.xlsm`, `.xltm`, `.xlsb`, and unsupported packages. |
| Ordinary worksheets | **Preserved unchanged** where possible without repetition or PrintArea: sheet order/visibility, original coordinates, sparse and formatting-only cells, and supported workbook content. |
| Scalar expressions | **Supported.** JSON paths, escapes, missing-value behavior, and registered formatters. Complete unformatted values retain supported native Excel types; interpolation and formatter output are text. Arrays and objects cannot be inserted as scalar cells. Inserted text is inert: leading `=` does not create a formula, and braces are not re-evaluated. Prohibited XML 1.0 characters in evaluated replacement text are removed with a warning. |
| Numbers and dates | **Supported:** finite, Excel-safe numeric values with at most 15 significant decimal digits; authored number formats and date systems (1900/1904). ISO date strings remain text. Unsafe numeric values are **rejected**. |
| PrintArea | **Supported.** A single finite rectangular area defines retained content, cropped and moved to A1. **Required** for repetition; explicit areas also crop nonrepeating sheets. Invalid, disjoint, duplicate, or ambiguous PrintArea definitions are **rejected**. |
| Vertical repetition | **Supported.** Nested group headers/footers, multirow blocks, and independent runs. Malformed markers, invalid hierarchy, and inconsistent level controls are **rejected**. |
| Horizontal repetition | **Supported.** Nested blocks, sibling regions, parent-only columns, and multirow headers. Only one outer horizontal region per worksheet is allowed; child regions must fit wholly within their parents. Invalid marker geometry is **rejected**. |
| Combined repetition | **Supported.** Hierarchical rows and dynamic columns may expand together. |
| Data hierarchy and `$path` | **Supported.** `$level` and input order determine the vertical hierarchy. For horizontal intersections, contiguous records with a shared `$path` must cover the complete positional combination of leaf columns. Reopened paths or incomplete combinations are **rejected**; incorrect value order cannot always be detected. |
| Cell formatting | **Preserved** where supported: cell styles, fonts, borders, number formats, rich-text formatting, and compatible text runs. |
| Row and column properties | **Preserved** where supported: row heights, column widths, hidden states, and zero-size rows/columns. |
| Merged cells | **Supported** when wholly retained or validly owned by expanded regions. One-cell merges become ordinary cells and empty merges disappear. Partial, overlapping, cross-owner, or data-losing merges are **rejected**. |
| Formulas | **Supported:** ordinary A1 cells/ranges, relative (`A1`), absolute (`$A$1`), mixed (`$A1`, `A$1`) references, supported same-workbook sheet references, and compatible shared formulas. References are relocated when unambiguous. Cross-sheet references require one unique target cell or a complete contiguous target range; records on different worksheets are not implicitly matched. MonoReport does not calculate formulas. |
| Unsupported formulas | **Rejected:** `INDIRECT`, `OFFSET`, structured, external, 3-D, whole-row/column, array/data-table, spill and implicit-intersection references; incomplete, disjoint, removed or ambiguous targets. |
| Defined names | **Supported:** safe, worksheet-qualified absolute cells/ranges in supported workbook or sheet scopes. **Rejected:** relative/mixed or unqualified, formula-valued, recursive, multi-area, unsupported function/macro, or ambiguous definitions. Worksheet-local names take precedence; names keep fixed global meaning across repetition, so named cells must remain unique and ranges complete and contiguous, even when unused. |
| Print settings | **Preserved** where supported: margins, orientation, scaling, and non-coordinate settings. |
| Manual page breaks | **Supported** on nonrepeating sheets when coordinates can be safely preserved or adjusted after cropping. **Rejected** on sheets with repetition. |
| Frozen and split panes | **Preserved** on nonrepeating sheets only when cropping leaves their positions valid. **Rejected** on repeating sheets and when cropping invalidates a pane. |
| Selections and active cells | **Preserved or adjusted** when safely representable. Removed or ambiguous selections reset to a valid output position; invalid or inconsistent coordinates are **rejected**. |
| Print Titles | **Rejected.** Unlike PrintArea, repeated print-title rows/columns are unsupported. |
| Excel tables (Insert → Table) | **Rejected.** Structured Excel Table objects are not ordinary formatted cell ranges; use ordinary cells instead. |
| AutoFilter | **Rejected.** |
| PivotTables | **Rejected.** |
| Images and drawings | **Rejected.** Embedded spreadsheet images and drawing objects are not currently preserved. |
| Charts | **Rejected.** |
| Conditional formatting | **Rejected.** |
| Data validation | **Rejected.** |
| Hyperlinks | **Rejected.** |
| Comments | **Rejected.** |
| Phonetic annotations | **Preserved or adjusted** for unaffected text where supported; annotations overlapping replacements are **removed**. |
| Other workbook metadata | **Preserved** when recognized and compatible. **Removed:** stale formula caches/calculation chains and recognized transient save-path or revision metadata. **Rejected:** unsupported OOXML extensions, relationships, and cell/row metadata (for example, unrecognized `AlternateContent` or custom extension records). Worksheet synchronization anchors must remain unchanged by cropping or expansion; otherwise generation is **rejected**. |

## Important Limitations

1. **PrintArea is a processing boundary:** control cells sit outside the retained region. A sheet without repetition and without PrintArea keeps its original coordinates.
2. **Horizontal binding is positional:** each logical `$path` must supply records in column order for the complete horizontal product, except the supported empty-product descriptor case.
3. **Formulas are not evaluated:** cached results are removed and recalculation is requested when opened in Excel.
4. **Generation is not Excel rendering:** visual fidelity, pagination, printing, application-specific resource limits, and calculated results must be checked in an actual spreadsheet editor.

## Data Binding

1. XLSX input data must be a JSON object. Ordinary and parent-only fields read the first record of a logical group; leaf fields read their corresponding horizontal tuple, and headers use their own horizontal collections.
2. Parent paths follow actual data ancestors, even where the template omits an intermediate level. Unavailable ancestors render empty with a warning by default and fail in strict missing-value mode.

## Validation

1. Analyzer checks data-independent syntax and structural compatibility without JSON data; successful analysis does not prove data-dependent row/column counts, ancestor availability, or formula transformations.
2. Generation rejects unsupported or ambiguous structures with diagnostics instead of silently producing invalid output.
