# Feature 015 — Workspace UI Redesign

**Status:** Approved by Product Owner — 2026-10-28  
**Feature directory:** `.docs/features/015-workspace-ui-redesign/`  
**Reference artifacts:** `ui-mockup.html`, `implementation-guidelines.md`

## 1. Purpose

Redesign the MonoReport workspace as a calm, efficient place to edit a Dataset, view a DOCX/XLSX Template, generate and inspect a separate Result, and diagnose issues. Improve the existing UI through simpler composition and existing Ant Design capabilities rather than rebuilding the application or adding custom design infrastructure.

## 2. Scope and principles

1. Preserve existing document generation, analysis, example loading, Monaco editing, and DOCX/XLSX viewer integrations unless a specific requirement below demands a change.
2. Use native Ant Design components, icons, layout APIs, and theme tokens by default; keep any required custom CSS minimal. The neutral HTML mockup illustrates composition and behavior, not colors, pixels, custom controls, or production code.
3. Start with a focused audit and safe removal of demonstrably obsolete or redundant CSS. Do not remove CSS still required by the current interface merely because a later stage may replace it. Remove newly obsolete CSS alongside the corresponding replacement.
4. Keep the feature implementable and independently reviewable in small stages. Obtain Product Owner acceptance of each UI stage before continuing to the next; final stage ordering and dependencies belong in the Implementation Plan.
5. Preserve existing responsive behavior, keyboard accessibility, visible focus, and supported workflows. Desktop is the primary target; narrow/mobile screens are lower priority and do not warrant a separate redesigned UI.
6. Do not introduce Projects, accounts, cloud storage, a data source catalog, Office add-ins, or new business capabilities as part of this feature.

## 3. Workspace layout

1. The top application toolbar displays the MonoReport identity on the left without a redundant `Playground` subtitle or separator, and one right-aligned action group in this order: `Examples`, `Analyze` with dropdown arrow, `Generate`, application `…`.
2. The main workspace has two primary regions: an always-visible editable `Dataset` on the left and document viewer on the right. It does not permanently display three equal-width Dataset/Template/Result columns.
3. Use an Ant Design `Splitter` for adjusting region widths, with a visually narrow divider and a usable drag target. Avoid a custom splitter implementation where the library suffices.
4. The right region switches between `Template` and `Result` using standard Ant Design radio-button-style controls. Show the active Template filename beside the tabs; preserve viewer-specific controls for zoom, pages, sheets, and other viewing capabilities rather than duplicating them.
5. The document region collapses and restores through native Splitter controls without unloading the selected Template or Result.
6. Each main region has a compact header with the active resource filename and applicable overflow menu. Keep filenames and the Dataset dirty indicator grouped and safely truncated at narrow widths. In the document header, place the filename directly after the Template/Result tabs and place a stale Result's `Outdated` indicator at the far right immediately before the overflow menu. Identify Example versus Local files in a tooltip; use `colorTextTertiary` for Example filenames and `colorText` for Local filenames. Do not show the Example title or recent-file selectors.
7. The bottom diagnostics region is collapsible and contains `Problems`, `Analysis`, and `Output` tabs. The status bar remains visible when diagnostics are collapsed.
8. Use existing Ant Design layout and spacing conventions. No visual requirement depends on reproducing mockup colors, borders, or precise dimensions. The goal is reduced visual noise and more room for the editor and viewer.
9. Preserve existing narrow-screen behavior where practical, without prioritizing phone usability or inventing a separate mobile interface.

## 4. Examples

1. `Examples` is in the top-right action group immediately before Analyze, not beside the application logo.
2. For this feature, use a standard Ant Design dropdown menu grouped by document type (DOCX and XLSX). No sidebar/catalog explorer/search is required.
3. Selecting an example loads its existing template and JSON data using the current sample/example mechanism.
4. Loading an example must not silently discard unsaved Dataset changes. Use the same unsaved-change safeguard as when replacing the Dataset.
5. On first visit or after Clear Local Data, open the default built-in example. On subsequent visits, restore the last working context, including the last selected example or previously opened local files where access is available. Ask for renewed file permission when the browser requires it; do not silently substitute the default example when restoration needs user action.
6. Selecting an Example loads its Dataset and Template. Opening a Local Dataset or Template afterward replaces only that resource. Example filenames identify the bundled file, not the Example title, and are not treated as disk-backed files.

## 5. JSON data editor and files

1. Keep Monaco as the always-visible Dataset editor, including when its content is empty, with existing formatting and editing capabilities.
2. The Dataset header shows one active filename as non-interactive text beside `Dataset`, with an adjacent `…` menu. Keep the unsaved-changes indicator immediately beside and visually grouped with the filename, including when it truncates.
3. Keep one active Dataset. The menu exposes `Open…`, `Save`, `Save As…`, `Format`, `Word wrap`, and `Copy`; there is no Dataset `Close` action. Use applicable existing Monaco actions for formatting, wrapping, and copying.
4. Opening a JSON file from disk replaces the active Dataset while preserving unsaved-change protection and disk-file associations. Preserve supported bounded file-handle restoration without adding recent-file selectors.
5. Where supported, `Save` writes to the associated file and `Save As…` selects a new file name/location, subject to browser permissions. Without direct file saving, the existing Save actions use the JSON download fallback; do not add a separate Dataset Download action. Use `JSON file` when naming the format and `Dataset` for the workspace concept.
6. Keep a clear distinction between an Example Dataset, an in-memory unsaved Dataset, and a disk-backed Local file: available save behavior may differ. Replacing the Dataset must not silently discard edits.
7. Distinguish unsaved Dataset edits from an outdated document Result. Saving the Dataset alone does not make the generated Result current.
8. Do not implement a persistent JSON-content library, full file explorer, Projects, or cloud synchronization. Persist only existing supported file handles/context metadata and gracefully handle unavailable or revoked permissions.

## 6. Templates and document results

1. The document header includes the `Template` / `Result` switch, the active Template/Result filename immediately after the tabs, and a `…` menu. When Result is active and stale, show `Outdated` as a native Ant Design Tag with `color="warning"`, `variant="outlined"`, and an `ExclamationCircleOutlined` icon at the far right immediately before the menu; keep the filename in its stable position after the tabs. Identify the resource's Example or Local source in a tooltip; do not show the Example title.
2. Retain applicable `Open…`, `Download`, and `Close` actions for Template and Result. Remove unsupported Save/Save As actions. Avoid redundant context words in menu labels.
3. Uploading a template replaces the active Template directly; the user need not close it first. Preserve the active Dataset.
4. When no template is selected, show a clear empty state with an upload button and drag-and-drop using the standard Ant Design Upload functionality where appropriate.
5. Closing a template does not close or overwrite JSON. Switching templates preserves user JSON and marks any last result as potentially outdated.
6. Keep one active Template. Opening a Template from disk or selecting an Example replaces only the Template, preserving the active Dataset. Preserve supported bounded local-file handle restoration, request permission to reopen as needed, and handle browsers without file-handle persistence gracefully.
7. Keep the last successfully generated document displayed after JSON edits, template changes, or unsuccessful subsequent generation. Indicate clearly when it no longer reflects the current input. Do not destroy the last successful result merely because its inputs changed.
8. An unsuccessful generation must not replace the last successful document with an empty result. Expose the error through diagnostics and status.
9. A first-time session without a successful generation should show a meaningful `Not generated` or equivalent state rather than implying success.
10. Leave zoom/page/sheet UI to the native DOCX/XLSX viewers where supported. Add no duplicate toolbar solely for those functions.
11. Name a generated Result from the Template used for that generation: remove a trailing `.template` immediately before the file extension, or append `-result` before the extension when that suffix is absent. Keep that source association and filename after the active Template changes, and use the same derived name in Result and Download.
12. Restore main splitter sizes and collapsed states and the lower diagnostics panel height/collapsed state after reload. Expanding a collapsed panel restores its last user-selected size. Include these preferences in Clear Local Data and preserve responsive behavior.

## 7. Primary actions and Generate JSON

1. `Generate` stays a prominent primary action in the top toolbar. `Analyze` is next to it and is the primary portion of a standard Ant Design `Dropdown.Button` (or supported equivalent).
2. Clicking Analyze performs the existing template analysis and makes its structured outcome available in the `Analysis` tab.
3. The dropdown associated with Analyze contains `Generate JSON…`, since generating sample data is related to template analysis.
4. `Generate JSON…` opens an Ant Design modal with only two options: `Items per collection` (`InputNumber`) and values mode (`Random values` or `Empty values` via radio buttons).
5. **Stage 4 Product Owner decision (2026-10-30):** default to 10 items per collection; accept positive integers with no arbitrary count maximum. Apply the count recursively to arrays with known object fields, leaving unknown-shape arrays empty. Limit the complete dataset to 100,000 generated array elements across all nesting levels and sibling collections; reject excess before allocation or dataset creation. Empty mode uses empty strings/false; Random mode uses generic sample strings/random booleans without inferred number/date types. Do not invent per-collection configuration.
6. Generated JSON becomes the active in-memory Dataset after existing unsaved-change protection; it must not silently overwrite user-edited JSON. Do not show a Destination choice in the dialog.
7. Suggest a generated Dataset filename from the source Template basename: remove a trailing `.template` from the basename when present, then append `-generated.json`. Keep the existing internal resource IDs and Dataset creation semantics.
8. The command should use the existing analysis/generation pipeline where available, avoiding unnecessary duplicate analysis. Any limits in the current generator should be disclosed during planning rather than silently fabricating richer behavior.
9. Place an application-level `…` immediately after Generate. Include `API documentation` (using the existing application URL) and `Clear Local Data…`. Do not add speculative menu commands to fill space.

## 8. Diagnostics and status

1. The collapsible diagnostics region uses Ant Design Tabs: `Problems`, `Analysis`, `Output`.
2. `Problems` presents meaningful errors/warnings from operations; where possible distinguish their originating action. A failed Generate should be visible here without erasing the previous document.
3. `Analysis` shows the available structured results from the existing Analyzer. Do not invent semantic validation or field metadata beyond what the actual API supplies. Structure the presentation using suitable standard UI components.
4. `Output` is a chronological in-session operation history including at least action, start time, result (success/failure), and actual elapsed duration for Generate, Analyze, and Generate JSON. Show running state where practical.
5. The `Clear` action belongs inside the Output tab content, right-aligned above the event list, not in the shared tabs header. Clear only the Output event history.
6. Keep Output in memory for the active application session; do not add a logging backend or persistence subsystem. Do not log complete input JSON or document contents.
7. **Stage 4 Product Owner decision (2026-10-30), with Stage 5 presentation refinement:** the persistent bottom status bar shows the latest Output event with its action, outcome, start time, and measured duration. Clicking it opens Output; keep Problems count separately. Before any operation, show an honest empty-operation state. Output Clear preserves this latest status summary. When Result is the active tab and the retained result differs from current inputs, show compact `Outdated` as an outlined warning-colored Ant Design Tag with an `ExclamationCircleOutlined` icon at the far right immediately before the document overflow menu with a Tooltip explaining the changed Dataset or Template. Keep the filename immediately after the tabs in the same position as freshness changes. Never show the indicator in the status bar or while Template is active. Existing result freshness semantics remain unchanged.
8. Status and diagnostic visibility are independent: the status line is visible even when diagnostics are closed. Allow the user to reopen each diagnostic tab from the bottom area.
9. Keep the existing GitHub link in the status bar. Move the existing API link to the main `…` menu, preserving the real destination.
10. Status/error colors should follow Ant Design semantic tokens; the mockup's neutral palette must not become a custom theme mandate. Do not introduce additional red brand styling.

## 9. Clear Local Data and confirmations

1. `Clear Local Data…` resides in the main application `…` menu.
2. Before clearing, display an Ant Design standard confirmation briefly describing the local data removed, warning that unsaved workspace changes will be discarded, and assuring that files on disk will not be deleted. Never imply the command deletes the user's original JSON/DOCX/XLSX files on disk.
3. Confirm potentially destructive changes (closing/replacing unsaved JSON, replacing work by opening an example, and similar actions) through standard Ant Design `Modal.confirm` or `Popconfirm` when appropriate.
4. Reuse the application's existing local-data clearing behavior and verify its actual scope before implementation. Do not expand it to unrelated browser/site storage.
5. Do not force confirmations for routine non-destructive navigation or saving.

## 10. Keyboard commands and discoverability

1. Offer practical shortcuts for JSON save/open/save-as where browser support permits, formatting, word wrap, Generate, Analyze, and other repeatedly used operations only where a conflict-free binding is available.
2. `Cmd+S` / `Ctrl+S` saves the active Dataset and `Cmd+Shift+S` / `Ctrl+Shift+S` invokes Save As throughout the workspace, including when Monaco or a workspace menu has focus. Prevent the browser Save Page action for these commands. Let dialogs and ordinary non-Monaco editable controls handle their own input; preserve Monaco's editor commands and macOS versus Windows/Linux conventions.
3. Display only actually implemented, platform-correct shortcuts in the relevant dropdown menus/tooltips.
4. Prefer Monaco's command infrastructure and standard focus handling. A separate global command palette is not required for this feature.
5. Keep primary actions discoverable without requiring the user to memorize shortcuts.

## 11. Technical implementation guardrails

1. Prefer Ant Design `Flex`/`Space`/`Button`, `Dropdown`/`Menu`, `Dropdown.Button`, `Select`, `Radio.Group`/`Radio.Button`, `Splitter`, `Tabs`, `Modal`, `Form`, `InputNumber`, `Upload.Dragger`, `Empty`, and `@ant-design/icons` as appropriate. Use existing native viewer/editor UI where stronger.
2. Use standard Ant Design confirmation, feedback, accessibility, and focus behavior. Avoid homegrown menus, radio-switches, dialogs, tabs, file pickers, and draggable dividers.
3. Avoid custom CSS except where a concrete requirement cannot be fulfilled using framework tokens/APIs; prefer existing theme configuration. Document unavoidable exceptions concisely.
4. Keep the UI consistent with existing MonoReport and avoid broad refactoring or adding new dependencies unless approved.
5. Preserve existing working behavior during incremental migration. No unrelated application, API, Docker, or build changes are requested.

## 12. Acceptance criteria

1. The top-right action group appears in the agreed order without a redundant `Playground` subtitle; Examples is a grouped dropdown and Analyze has a `Generate JSON…` dropdown action.
2. JSON editor and document viewer are simultaneously usable and resizable, with a narrow standard splitter; Template/Result is a native radio-style switch.
3. Native Splitter controls collapse and restore the document region without losing its selected Template or Result.
4. One active Dataset is always editable in Monaco, even when empty. Its filename/source and dirty state are displayed together; Open, Save, Save As, Format, Word wrap, and Copy remain available without Dataset Close, and replacing it preserves unsaved-change protection. Workspace Cmd/Ctrl+S and Cmd/Ctrl+Shift+S perform Save and Save As even when focus is outside Monaco, while dialogs and ordinary editable controls retain their input behavior.
5. One active Template and its filename/source are shown beside the Template/Result tabs. Opening a Template replaces only that resource and does not wipe the Dataset; applicable Template/Result Download and Close actions retain distinct semantics, and an empty Template supports upload/drop.
6. The previous successful result stays visible and is correctly marked outdated after input changes or failed generation.
7. Generate JSON offers exactly the two agreed settings and opens a new Dataset without silently overwriting unsaved content. Its suggested filename derives from the source Template basename with a trailing `.template` removed when present and `-generated.json` appended.
8. Problems, Analysis, and Output work as separate bottom tabs; Output has its own Clear button and uses measured durations, not hardcoded mock timings.
9. Status shows the latest operation event, including executing/final outcome, action, start time and duration; when Result is active and stale, `Outdated` appears as an outlined warning-colored Ant Design Tag with an `ExclamationCircleOutlined` icon at the far right immediately before the document overflow menu with a Tooltip. The filename stays immediately after the Template/Result tabs as freshness changes. Problems count remains separate; clicking status opens Output. GitHub stays in status and API is under main `…`.
10. `Clear Local Data…` and any destructive file operations use standard confirmations and never delete source disk files unexpectedly.
11. Implemented keyboard shortcuts do not conflict with Monaco/browser behavior; menu hints are truthful.
12. The UI remains usable on supported desktop viewports and does not unnecessarily regress existing narrow-screen behavior; native controls and current viewer/editor behavior are preserved.
13. A focused comparison confirms that obsolete CSS was safely removed while still-needed CSS was preserved; no unnecessary custom UI control or styling infrastructure was introduced.
14. Product Owner can inspect and accept each UI migration milestone independently before the next begins.
15. Generated Result and Download filenames derive from the source Template and remain associated with that generation when the active Template changes.
16. Main-panel sizes/collapsed states and diagnostics-panel height/collapsed state restore after reload; expanding restores the last user-selected size. Clear Local Data removes these preferences, and the layout remains responsive.
17. The keyboard-shortcuts dialog uses only the `Workspace` and `Dataset editor` section headings, command names, and keyboard keycaps. It has no table headings, explanations, row separators, or redundant `(?)` suffix and remains compact and responsive.

## 13. Explicitly out of scope

1. Data Sources catalog, source configuration, and field insertion workflow.
2. Word/Excel Office add-ins and template authoring in the browser.
3. Projects, multi-template project workspaces, multi-user collaboration, and cloud file providers.
4. A permanent content library or unbounded file history, storing template binaries in IndexedDB/OPFS, or a new backend for sessions. Existing bounded handle/context restoration remains in scope; file selectors and multiple active workspace resources do not.
5. A new examples sidebar, custom command palette, new viewer engine, and custom design system.
6. Language switching; future placement can be reconsidered in the application `…` menu if needed.
7. Unrelated dependency audits, license reviews, bundle optimization, and telemetry work.

## 14. Approved product decisions and technical verification

1. **Restoration:** reopen the last working context on subsequent visits, including recent local JSON/templates where browser file permissions allow. On first visit or after clearing local state, open the default example. Store file handles only, not copies of local file contents; ask for permission when needed.
2. **Active files:** the workspace exposes one active Dataset and one active Template, with a separate retained Result. Existing bounded file-handle/context restoration may retain references for recovery, but the UI does not expose recent-file selectors or multiple simultaneously active resources.
3. **Saving:** where direct disk writes are supported, Save writes the associated file and Save As selects a destination. Otherwise retain Save and Save As as JSON download fallbacks without a separate Dataset Download command. Keep file picker descriptions as `JSON file`.
4. **Generate JSON:** the dialog has only collection item count and Random/Empty values; generated data opens as a new JSON dataset. The Stage 4 decision in Section 7, item 5 approves default 10, positive integer counts without an arbitrary maximum, and preflight rejection above 100,000 total generated array elements. Preserve successful matching-template analysis caching and all existing datasets.
5. **Analysis:** inspect the existing Analyzer response and display what it actually returns; do not invent fields or metadata.
6. **Responsive behavior:** desktop is the main use case; preserve existing narrow-screen support as reasonably possible, without adding major mobile-specific UI. Verify shortcuts against existing Monaco/browser bindings and platform conventions.
7. **Workspace simplification (2026-10-31):** keep one active Dataset, one active Template and a separate Result; retain filenames/source indicators and all supported file operations, with no file selectors or Dataset Close. Examples continue to load both resources and later Local opens replace only their matching resource.
8. **Result names and splitter state (2026-10-31):** generated result names derive from their source Template and remain associated after Template changes. Persist main and diagnostics Splitter sizes/collapse states, restore the last user-selected size on expansion, and include these preferences in Clear Local Data.
9. **Generated Dataset names and Save commands (2026-10-31):** derive suggested generated Dataset names from the source Template basename, removing a trailing `.template` and appending `-generated.json`, without changing internal IDs or creation semantics. Bind Cmd/Ctrl+S and Cmd/Ctrl+Shift+S to workspace Dataset Save and Save As while preserving dialog and ordinary editable-control handling and preventing the browser Save Page action.

## 15. Reference interpretation

1. `ui-mockup.html` is an illustrative interactive reference. It uses fake generated documents, mock file operations, sample analysis, and simulated timings; none of those internals are requirements.
2. `implementation-guidelines.md` identifies preferred standard Ant Design components and reinforces the project's existing conventions; it does not override this specification or AGENTS.md.
3. The approved Feature Spec is the behavioral source of truth. Implementation Plan stages, component placement details, and actual browser API limitations must be reconciled against the current code before implementation.
