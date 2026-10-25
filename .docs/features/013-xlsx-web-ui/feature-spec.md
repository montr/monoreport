# Feature 013 — XLSX Web UI Support

**Status:** Approved — 2026-10-25
**Feature directory:** `.docs/features/013-xlsx-web-ui/`  
**Architecture research:** [architecture-research.md](architecture-research.md)

## 1. Goal

Extend the existing MonoReport Web Playground to support XLSX templates and generated workbooks alongside DOCX, using the same core user workflow and existing document APIs. A user must be able to upload and inspect an XLSX template, enter JSON manually or generate a JSON starting point from template analysis, generate a workbook, inspect the result, and download the original generated XLSX file.

## 2. Existing capabilities and constraints

1. The Playground currently supports DOCX upload, template/result previews, a Monaco JSON editor, Generate JSON, Generate, diagnostics, examples, and download.
2. Existing `POST /api/documents/generate` accepts DOCX and XLSX templates. XLSX generation is already implemented in Core and the API.
3. Existing `POST /api/documents/analyze` supports XLSX through the shared Analyzer contract (`valid`, `errors`, `warnings`, `requirements`, `occurrences`). The XLSX Analyzer does not need to be implemented as part of this feature.
4. The current DOCX preview uses `docx-preview` 0.4.1 with a dedicated compatibility normalization step. Preserve this behavior.
5. The XLSX generator preserves formulas as formulas but clears cached formula results and requests recalculation. The browser preview does not calculate formulas.
6. Only workbooks within MonoReport's existing XLSX engine compatibility rules can be generated successfully. Browser preview support is not an expansion of those engine rules.

## 3. Required user behavior

### 3.1. Template selection and upload

1. The existing Upload action and drag-and-drop flow accept `.docx` and `.xlsx` files, with suitable picker filters and visible UI copy.
2. The selected file extension guides the initial template-preview format and component; it is not authoritative validation of document content. An `.xlsx` file must never be sent to DOCX preview or treated as a DOCX download.
3. Template changes clear or invalidate the previous generated result as appropriate and prevent outdated asynchronous operations from replacing newer content.
4. Unsupported or invalid documents receive clear feedback. The server remains authoritative for validating document content; a filename extension alone does not establish a valid template.

### 3.2. Template preview

1. The Template tab offers a read-only browser preview of the selected XLSX workbook without requiring generation.
2. Users can switch between visible worksheets in their authored order. Hidden worksheets do not need a navigation tab.
3. The viewer displays actual worksheet content and a practical subset of formatting: cell values/text, template expressions, number formats, fonts, fills, borders, alignment, merged cells, row/column sizes, and hidden rows/columns where supported.
4. XLSX preview uses a scrollable spreadsheet viewport suitable for the existing Splitter layout, rather than the DOCX paper-page layout. It remains usable in the current responsive UI.
5. Preview fidelity is approximate: exact Microsoft Excel visual or printed-page parity is not promised.
6. If preview fails or refuses a document for resource/safety reasons, the user receives a clear preview-specific message; the file can still be used for generation, subject to server validation.

### 3.3. JSON input and Generate JSON

1. The current Monaco JSON editor is reused without introducing an XLSX-specific data editor.
2. Users can enter, paste, and edit JSON manually for XLSX templates.
3. **Generate JSON is available for XLSX**, using the already implemented template analysis API and the same applicable user flow as DOCX. This confirmed Product Owner decision supersedes the earlier deferral recommendation in the architecture research (sections 6–9); the research is retained as historical context.
4. The generated JSON is a starting point derived from the Analyzer's requirements, not a guarantee that all data-dependent XLSX generation constraints will pass.
5. Analyzer errors and request failures are reported through existing UI mechanisms. Analysis must not be a mandatory preflight for manually entered JSON or block the normal Generate action.
6. A separate standalone XLSX analysis panel or Analyzer enhancement is not required.

### 3.4. Generation and result preview

1. Generate submits the selected DOCX/XLSX template and current JSON through the existing `POST /api/documents/generate` API; no new generation endpoint or engine behavior is required.
2. The successful generation response MIME type determines the result format and corresponding preview component, independently of the uploaded file extension. A successful XLSX generation produces a Result tab preview using the same XLSX viewer as the Template tab, including worksheet navigation.
3. The preview uses a separate view of the returned document bytes. Preview must not mutate the original downloaded workbook.
4. Rendering a result is independent of generation success: if the result cannot be previewed, the generated workbook remains available for download.
5. On a generation failure, show existing diagnostic information, including code/message and available field, data path, workbook part or source location. Do not present a preview error as a generation error.
6. Preserve existing cancellation and stale-result protection when switching templates or generating multiple times.

### 3.5. Formula preview limitation

1. Do not add formula evaluation or workbook recalculation to the browser or backend as part of this feature.
2. Display cached formula values only when present in the workbook. Formula cells without a cached result **may appear blank** in the preview, including newly generated MonoReport XLSX files.
3. A missing cached value must not be misrepresented as zero, as a successful calculation, or as a workbook generation error.
4. The downloaded workbook retains its original formulas and recalculation settings. Opening it in an application capable of recalculation is outside the browser-preview responsibility.

### 3.6. Download and existing UI parity

1. Download returns the original generated document bytes without exporting or serializing from the viewer. Sanitize any response-provided filename and require an extension consistent with the response MIME type; otherwise use a safe, format-matched `.docx` or `.xlsx` filename fallback.
2. Preserve the current DOCX upload, Generate JSON, generation, preview, error handling, examples, and download behavior.
3. Retain the existing Template / Result / Errors structure, Monaco editor, upload and drop interactions, and general layout. Do not create a separate XLSX Playground.
4. If an XLSX example is added to the existing Examples catalog, it must follow the same loading and revision guards; **bundling additional XLSX examples is optional**, not a prerequisite for completing the primary XLSX workflow.

## 4. Agreed architectural boundaries

1. Add an XLSX-only browser preview based on the architecture research's selected `@silurus/ooxml` **0.89.4** candidate; keep the existing DOCX viewer unchanged. Any later viewer consolidation is a separate decision and feature.
2. Prefer the XLSX-specific package entry point, lazy loading, and locally served JavaScript/WASM/worker assets to avoid increasing the initial DOCX page payload unnecessarily.
3. Keep frontend integration narrow: explicit document-format handling and selection of the corresponding viewer are sufficient. Do not introduce a generic document plugin framework.
4. Reuse existing document upload, analysis, generation, diagnostics and download APIs. No new backend endpoint, formula engine, server document converter, or document persistence is in scope.
5. Follow project dependency rules: permissible open-source licenses, required distribution notices, security review, and pinned/controlled dependency upgrades. Do not adopt AGPL/GPL or license-restricted alternatives without a separate decision.
6. Work in the already created feature branch. Do not introduce feature flags solely to hide this work or change the existing DOCX viewer.

## 5. Error handling, safety, and nonfunctional requirements

1. Uploaded and generated XLSX files are untrusted inputs to the new XLSX viewer. The XLSX viewer must not execute workbook macros/formulas, activate external links, request external fonts/resources, or render workbook strings as untrusted application HTML. This requirement does not change the existing DOCX preview behavior.
2. Prevent obsolete preview sessions, workers, or callbacks from changing the UI after replacement, unmounting, or format switches. Preserve session invalidation, viewer destruction, and resource cleanup. **Accepted limitation — Product Owner decision, 2026-10-25:** In pinned `@silurus/ooxml` 0.89.4, a pending initial-load parsing worker may survive viewer destruction until loading succeeds, fails, or a worker request times out; completed workbook resources are disposed through existing cleanup. Retain the configured 30-second worker-request timeout, which is not a total-load or post-disposal deadline. Rapid switching can temporarily overlap obsolete parsing work. No permanent leak or demonstrated user-facing failure was established, and a maintained dependency patch is not justified at this time. Do not patch or replace the dependency for this limitation; preserve stale-operation protections.
3. Treat preview memory/resource limits as independent of the existing API request-size limit. Failure to preview an oversized or unsupported workbook must be recoverable and must not prevent download of an already generated result.
4. Ensure XLSX viewer JavaScript, WASM, and worker assets load from the deployed application origin, including production builds; do not depend on third-party hosted services at runtime.
5. Keep preview read-only. No spreadsheet editing, formula recalculation, chart/advanced-feature fidelity promise, or Excel print-layout parity is required.
6. Do not introduce unrelated refactoring, production configuration changes, feature flags, or new automated end-to-end testing infrastructure solely for this feature.

## 6. Acceptance scenarios

1. **DOCX regression:** existing DOCX upload, Generate JSON, preview, generation, diagnostics, examples, and download still work as before.
2. **XLSX template:** selecting or dropping a valid MonoReport XLSX template shows its worksheet contents in Template preview, including visible sheet tabs and representative formatting.
3. **Manual JSON:** with an XLSX template, entering valid JSON and clicking Generate returns a valid XLSX workbook visible in Result and downloadable as XLSX.
4. **Generated JSON:** with an XLSX template, Generate JSON calls existing analysis and populates the JSON editor with the supported starting structure; users can edit it and generate normally.
5. **Multiple worksheets:** a workbook with two visible sheets can be navigated in Template and Result; the selected sheet's contents change correctly.
6. **Formula cells:** a fresh generated workbook with missing formula caches shows blank formula results in preview without converting formulas to zeros or treating generation as failed. The downloaded workbook still contains its formulas.
7. **Recoverable preview failure:** an intentionally unrenderable/over-limit result shows a preview error while its already generated XLSX remains downloadable.
8. **Generation diagnostics:** a bad template or data error appears in Errors with available context; XLSX analysis is not required before Generate.
9. **Format switching:** changing between XLSX and DOCX while previews/requests are in flight does not leak the old preview/result into the newly selected file.
10. **Integration smoke check:** the production build serves the XLSX runtime assets, and a representative generated XLSX download opens as a valid workbook in a desktop spreadsheet application.

## 7. Explicitly out of scope

1. Replacing or consolidating the DOCX viewer.
2. Enhancing the XLSX Analyzer or adding a separate analysis UI.
3. Server-side PDF/HTML conversion, Office-hosted document services, and a custom spreadsheet renderer.
4. Editing XLSX in the browser, evaluating formulas, populating formula caches, or modifying XLSX generation semantics.
5. Supporting workbook features rejected by the current MonoReport XLSX engine, or guaranteeing exact Excel layout fidelity.
6. Presenting successful-generation warnings through the optional multipart response protocol.
7. Extensive benchmark suites, multi-viewer browser test matrices, and a new generic document-rendering framework.
8. Redesigning the existing Errors tab or adding a separate analysis view. If richer analysis or diagnostics presentation becomes necessary in the future, evaluate the UI options at that time (e.g., a new tab, extending or renaming Errors, or another approach). No future design is selected or preferred by this Feature Spec.

## 8. Verification and handoff

1. Developer prepares an Implementation Plan against this approved Feature Spec and the linked architecture research; no implementation starts until Product Owner approval.
2. Include focused frontend tests for format detection/dispatch, JSON analysis action, generation response handling, diagnostics, preview isolation, and stale operations. Reuse existing test patterns.
3. Perform a bounded real-browser acceptance pass on representative XLSX and DOCX samples, including a multi-sheet workbook and a generated formula workbook. Check production asset loading and a manual Excel-compatible download/open flow.
4. If the selected library fails acceptance or licensing checks, return to Product Owner with evidence and a narrower alternative; do not silently substitute another viewer or expand feature scope.
