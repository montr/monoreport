# DOCX Template Compatibility

MonoReport supports DOCX-based document generation while preserving compatible Microsoft Word content. It is not a general-purpose Word editor. For template syntax and examples, see the [DOCX Template Guide](docx-template-guide.md).

MonoReport checks the structures needed for safe template processing; it is not a general Word, Excel, or OOXML validator. A source document may be processed even if an Office application cannot open it, provided MonoReport can read the required content and no documented compatibility restriction prevents generation.

A feature can be **Supported**, **Preserved unchanged**, **Removed**, or **Rejected**, depending on its location and interaction with template operations.

## Compatibility

| Word feature | MonoReport behavior |
|---|---|
| Document formats | **Supported:** DOCX and existing supported Word OOXML package variants (DOTX, DOCM, DOTM). This document describes DOCX template behavior; macro execution is not provided. |
| Text and formatting | **Supported.** Replace expressions while preserving supported character and paragraph formatting. Prohibited XML 1.0 characters in evaluated replacement text, including formatter output, are removed with a warning. |
| Tables | **Supported.** Fields and structural blocks work in tables. Paragraph-level blocks can contain complete tables, including vertical merges. Row-level `For`/`If` blocks **reject** vertical merges intersecting the operation. |
| Images | **Supported.** Existing images are preserved, including in repeated content and Raw blocks. |
| Lists and numbering | **Preserved unchanged.** Styles and numbering definitions remain; automatic restarts and recalculation are not guaranteed. |
| Comments | **Preserved unchanged** when unaffected, inside or spanning valid Raw, or fully enclosing `For`/`If` when endpoints and required attachments remain intact outside the operation and the resulting range remains valid. **Removed** when safely attached only to a deleted Raw delimiter. **Rejected** when local to `For`/`If` or left invalid by an operation. |
| Bookmarks | **Preserved unchanged** when unaffected, inside or spanning valid Raw, or fully enclosing `For`/`If` when endpoints remain intact outside the operation and the resulting range remains valid. **Removed** when safely attached only to a deleted Raw delimiter. **Rejected** when local to `For`/`If` or invalidated by an operation. |
| Permission ranges | **Supported** when endpoints remain valid, including ranges local to a repeat body/conditional branch and enclosing ranges. Repeated local ranges use distinct identifiers. Affected ranges with invalid, crossing, incomplete, misordered, or duplicate endpoints within their story and operation boundaries are **rejected**; unaffected ranges are preserved. |
| Internal hyperlinks | **Preserved unchanged** when outside `For`/`If` and their targets remain valid; unrelated pre-existing unresolved targets are not automatically repaired. **Rejected** within `For`/`If`, or when an affected target is removed, repeated, or ambiguous. |
| External hyperlinks | **Preserved unchanged**, including links and fragments. |
| Table of contents | **Preserved unchanged.** Word fields and cached contents are not recalculated; associated bookmarks/links remain subject to the rules above. |
| Word fields and cross-references | **Preserved unchanged.** Field instructions and cached results, including `REF`, `PAGEREF`, and field hyperlinks, are opaque to template expressions and are not evaluated by MonoReport. Hyperlink display text is also opaque. |
| Content controls | **Preserved unchanged.** Expressions inside content controls are not evaluated or assembled across their boundaries. |
| Tracked changes | **Rejected** when recognized unresolved revisions occur in inspected main content, headers, footers, footnotes, or endnotes, including Raw, inactive branches, and content controls; drawing/picture descendants are excluded from revision inspection. Accept or reject revisions in Word first. |

## Structural Operations

1. **`For` and `If`:** Place markers in dedicated sibling paragraphs or supported dedicated table rows. Paragraph blocks may enclose complete tables; row blocks remain within one table. Their entire marker/body footprint is subject to compatibility checks, including in inactive branches and zero-iteration loops.
2. **`Raw`:** Use exact `{Raw}` and `{/Raw}` markers in matching dedicated sibling paragraphs or rows. The body is literal, including expressions and braces, and Raw does not nest: another opening marker inside its body is literal, and the first eligible close ends the block. Raw may occur inside a `For`/`If` body but cannot cross its boundaries.
3. **Raw preservation:** Complete tables, supported merges, and multiple paragraphs within a cell are allowed. Delimiter removal must not damage a merge chain or section boundary, or leave surviving annotations or references invalid. Complete annotations attached only to a removed delimiter may be deleted safely; otherwise generation is rejected.
4. **Empty cells:** Structural operations may remove all authored content. MonoReport inserts a minimal empty paragraph where needed for valid Word table-cell structure, including cells ending in a nested table.

## Validation

1. Analyzer and generation apply the same data-independent document compatibility restrictions; unsupported structures produce actionable diagnostics rather than silently producing invalid output.
2. Preserved Word content is not interpreted, recalculated, or automatically repaired.
