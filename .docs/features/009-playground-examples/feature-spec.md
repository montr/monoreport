# Feature 009 — Playground Examples

Status: Completed

The Product Owner approved this specification and the Implementation Plan on 2026-10-16 and accepted the implemented Feature on 2026-10-16. Implementation used one agent without feature-orchestrator, independent review, browser automation, or Git operations.

## 1. Goal

1. Make MonoReport Playground usable immediately by loading a working template and its JSON data on startup.
2. Let users switch between bundled examples or upload and edit their own templates without losing unrelated JSON data.
3. Place document downloads in the corresponding Template and Result tabs, using format-neutral UI terminology.

## 2. Scope

1. Provide a bundled, static example catalog and local example assets, with one explicit default example.
2. Automatically load the default example's template and JSON on initial Playground entry and display its template preview. Do not generate a result automatically.
3. Keep the **Template**, **Result**, and **Errors** tab row separate. Add a contextual toolbar below it, with a compact **Examples** selector in the Template context, without introducing a new panel or separate demonstration mode.
4. Selecting an example loads its template and JSON into the existing Playground workflow.
5. Preserve the existing template-upload flow. Uploading a custom template replaces only the template, retaining the current JSON.
6. Provide **Download** in the **Template** tab for the current template and in the **Result** tab for the last successfully generated document. Relocate the existing result download action rather than duplicating it.
7. Retain the existing **Generate** and **Generate JSON** actions and their purposes.

## 3. Example catalog and assets

1. Store the catalog as a static JSON manifest bundled with the UI; no backend endpoint, database, or GitHub fetch is required.
2. The manifest defines `defaultExample` and an `examples` array. Each entry specifies a stable `id`, display `title`, document `format`, template asset path, and JSON-data asset path.
3. Bundle the referenced template and data files as static UI resources. Keep one authoritative copy of each sample asset wherever practical.
4. The initial catalog contains DOCX examples only. The `format` identifier anticipates future XLSX examples but does not authorize XLSX support, a generic document engine, or additional preview implementations.
5. The example selector shows human-readable titles in manifest order. The default example ID must identify an existing entry; invalid catalog entries or missing assets must fail visibly without corrupting the current editor state.

## 4. Template and JSON behavior

1. On successful initial example load, set the current template and JSON and show the template preview. The Result tab does not acquire a generated result until the user invokes Generate.
2. Selecting a different bundled example replaces both the current template and JSON, after confirmation when user changes would otherwise be discarded.
3. User changes include manual JSON edits and uploading a custom template. If the current state is unchanged from the loaded example, switching examples does not require confirmation.
4. When the user edits only the example JSON, retain the example selection and show an unobtrusive modified indicator next to its name.
5. When the user uploads a custom template, show **Custom** as the current selector state. **Custom** is a local state, not a manifest entry or selectable saved example; retain JSON exactly as it was.
6. **Generate JSON** continues to analyze the currently loaded template, including an uploaded custom template, and follows its existing replacement/confirmation behavior.
7. Asynchronous example loading must not overwrite newer user edits, uploads, or example selections. A canceled selection leaves both template and JSON unchanged. Failed loading must not leave a partially applied example.
8. Clearing the current template must not implicitly reload the default example.

## 5. UI and downloads

1. Keep the **Template**, **Result**, and **Errors** tabs in their own row and place a separate contextual toolbar immediately below it; preserve the current Playground panel layout.
2. In the Template context, place **Examples** on the left and **Upload** and **Download** on the right. Template Download retrieves the current template, whether it was bundled with an example or uploaded by the user.
3. In the Result context, place **Download** on the right. It downloads the last successfully generated output; disable or hide it when no output is available, consistent with existing UI conventions. The Errors context has no document actions.
4. Remove the result Download action from its previous location. Do not show duplicate result-download actions.
5. Use format-neutral labels: **Template**, **Result**, **Upload**, **Download**, **Examples**, and **Generate**. Preserve correct file extensions, media types, and download filenames for the actual document format.
6. Confirm destructive example switches in the same product interaction style as the rest of Playground. Do not silently discard user work.
7. In narrow layouts, keep the selector and tab actions usable without obscuring the preview or breaking the split-panel interaction.

## 6. Out of scope

1. XLSX rendering, generation, or upload support; additional document formats.
2. Example catalog management UI, remote example sources, personal example storage, search, tags, categories, and user accounts.
3. Persisting editor data or preferences across browser sessions.
4. Changing MonoReport Core, CLI, HTTP API, template syntax, or generation semantics.
5. General Playground redesign or global user-configurable editor/preview settings.

## 7. Acceptance criteria

1. A fresh Playground session loads the configured default DOCX example, populates JSON, renders the template preview, and does not call document generation automatically.
2. The Examples selector offers manifest-defined examples in order; choosing one loads the corresponding template and JSON in the existing UI.
3. Switching away from a modified example prompts for confirmation. Cancel preserves both current template and JSON; confirm replaces both.
4. Editing example JSON does not change the selected example identity, but visibly indicates modifications.
5. Uploading a custom template retains JSON byte-for-byte as editor text, changes the selector display to Custom, and permits Generate and Generate JSON with that template.
6. After custom upload, selecting an example requests confirmation before overwriting the template and JSON.
7. Rapid selection changes or edits during pending asset loads cannot cause an older request to overwrite newer state. A missing/malformed asset yields a controlled error and no partial replacement.
8. Template Download returns the currently loaded template with matching content and an appropriate filename, for both bundled examples and custom uploads.
9. Result Download is located in the Result tab, unavailable before a successful generation, and downloads the generated file unchanged; the old result download control is removed.
10. Template and Result labels are format-neutral and existing Generate/Generate JSON, preview switching, upload, editor, and layout behaviors continue to work.
11. Automated behavioral tests cover catalog loading, dirty-state confirmation, custom upload, stale-request protection, load failures, and both download flows. Browser smoke checks cover the main flow in Safari and Chrome as practical.

## 8. Approved decisions

1. Preserve the current behavior that clears Result after the template or JSON changes.
2. The initial catalog contains one default example named **Construction cost estimate**, using the authoritative `Samples/sample.template.docx` and `Samples/sample.json` files. The simulated examples in the HTML mockup are not production assets.
3. Continue to use `window.confirm` for destructive example switches.
4. Load the static registry and assets with ordinary browser `fetch`, load template and data in parallel, and apply them atomically. Use separate content revisions and asynchronous request identities; do not add custom caching or a generalized resource-loading framework.
5. Browser automation is excluded. The Product Owner will perform manual Chrome and Safari acceptance.

## 9. Delivery workflow

1. Obtain explicit Product Owner approval of this specification.
2. Use Codex `/plan` to inspect the current UI implementation, check actual sample assets, and propose a concise implementation plan and targeted verification before making code changes.
3. Obtain explicit Product Owner approval of the plan; implement with one agent and no mandatory orchestrator or independent-review gate.
4. Report tests, manual checks, limitations, and changes for Product Owner acceptance. Preserve Git index and do not stage, commit, or publish without separate authorization.

## 10. Approved cache revalidation correction — 2026-10-27

1. The Product Owner approved extending the examples caching fix to the manifest, DOCX/XLSX example templates and example JSON data. These resources can change across deployments while retaining their URLs; each fetch must use the standard no-cache request mode to revalidate cached responses.
2. Retain parallel template/data loading, atomic application of the selected pair and existing protection against obsolete asynchronous responses. XLSX support remains governed by Feature 013.
3. Keep resource URLs unchanged. Do not introduce cache-busting queries, custom caching infrastructure or deployment configuration changes.
4. Verify the request policy through the existing DOCX/XLSX loading tests and run focused App tests and TypeScript checking. Browser HTTP-cache behavior is supplied by the platform; component tests verify the request contract and loading behavior.
