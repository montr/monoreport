# DOCX Template Compatibility

> **Status:** Draft product policy, approved in principle. This document describes the intended compatibility contract for a future simplification feature. It does **not** describe all behavior of the current engine and does **not** change the acceptance criteria of Feature 010.

MonoReport generates printable documents from DOCX templates and data. It is not a general-purpose Word or OOXML editor. The goal is to support practical report and print-form templates while preserving unrelated Word content wherever safe.

## Compatibility rules

Each Word feature is classified by its behavior in context. A feature may be preserved in one situation and rejected in another.

| Word feature | MonoReport behavior |
| --- | --- |
| Plain text and character/paragraph formatting | **Supported.** Replace MonoReport expressions while preserving the surrounding Word formatting within the documented template syntax and rendering constraints. |
| Tables | **Supported.** Render supported expressions, loops, and conditions in table content; repeat supported table rows and other documented structural units. |
| Images | **Supported.** Preserve images in ordinary content and correctly handle images in supported repeated content. |
| Lists and numbering | **Preserved unchanged.** Retain Word list formatting and numbering definitions. Automatic numbering restarts or recalculation are not guaranteed. |
| Word comments | **Preserved unchanged.** Comments outside MonoReport structural operations, provided their ranges do not cross operation boundaries or deleted content.<br><br>**Rejected.** Comments inside `For` or `If` structural operations, or whose ranges intersect operation boundaries. MonoReport does not clone comments or their definitions. |
| Bookmarks | **Preserved unchanged.** Bookmarks outside MonoReport structural operations, with both endpoints in unaffected content.<br><br>**Rejected.** Bookmarks inside `For` or `If` structural operations, or crossing operation boundaries. Dynamic bookmark creation and renaming are not supported. |
| Internal hyperlinks | **Preserved unchanged.** Links to existing bookmarks when both the link and its target remain outside structural operations and their identity is unaffected.<br><br>**Rejected.** Links or targets affected by structural operations, including deleted, repeated, or ambiguous bookmark targets. MonoReport does not retarget links to cloned bookmarks. |
| External hyperlinks | **Preserved unchanged.** External hyperlink relationships, addresses, and fragments/anchors, including relationship-based links with `w:anchor`. They are not treated as internal bookmark references. |
| Automatic table of contents | **Preserved unchanged.** Word fields and their cached display content are not recalculated by MonoReport. Users update the table of contents in Word when necessary.<br><br>**Rejected.** Any associated bookmark or link structure that violates the structural-operation restrictions above. |
| Word cross-references and other fields | **Preserved unchanged.** Field instructions and cached results are not evaluated or refreshed by MonoReport. Fields are opaque to template expression substitution. This includes `REF`, `PAGEREF`, and field-based hyperlinks. |
| Content controls (`w:sdt`) | **Preserved unchanged.** Content-control subtrees, including bindings and contained text, are opaque: MonoReport does not interpret template expressions inside them. Content controls must not cause expressions outside them to be interpreted across their boundaries. |
| Tracked changes and revision markup | **Rejected.** Accept or reject all tracked changes in Word before using the template. Analyze reports a diagnostic identifying the unsupported revision markup. |

## Validation and diagnostics

1. **Validate before generating.** Analyze and Generate must apply the same compatibility rules. Unsupported structures must produce actionable diagnostics; Generate must not silently produce a corrupt DOCX.
2. **Explain the remedy.** Diagnostics should identify the affected Word feature and location, and tell the author what to change in Word (for example, remove a comment from a loop, move a bookmark outside a loop, or accept tracked changes).
3. **Preserve rather than interpret.** Features designated *Preserved unchanged* are not evaluated, updated, or automatically repaired by MonoReport.
4. **Keep structural boundaries explicit.** `For` and `If` denote MonoReport structural operations. A feature that is safe in ordinary content may be unsupported when its identity, range, or references would be copied or removed by such an operation.
5. **No implicit fallback.** Do not silently choose the first or last generated bookmark, share a repeated comment identity, or rewrite an external hyperlink fragment to make an unsupported template appear to work.

## Scope and adoption

1. This policy is a **target contract**, not a claim that the current implementation already behaves this way.
2. Feature 010 is accepted or rejected against its existing approved specification. Simplification of comment cloning, bookmark renaming, and internal-link remapping belongs to a **separate follow-up feature** after Feature 010 acceptance.
3. That feature should align Core, Analyzer diagnostics, tests, examples, and user-facing documentation with this policy, and remove obsolete code only after confirming safe rejection and preservation behavior.
4. Future changes to supported Word features must update this document and the corresponding compatibility tests together.
