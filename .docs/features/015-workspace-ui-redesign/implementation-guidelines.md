# MonoReport UI — Ant Design Implementation Guidelines

## Principles

1. Use standard Ant Design components and existing MonoReport editor/viewer integrations by default. The HTML mockup is a layout and behavior reference, not implementation code or a pixel-perfect styling contract.
2. Use Ant Design theme tokens and component APIs before introducing custom CSS. Avoid custom controls, menus, dialogs, tabs, radio-like switches, and splitters.
3. Keep the existing Monaco editor and DOCX/XLSX viewers. Do not reimplement features already provided by them, including document zoom when available.
4. Use minimal CSS only where a required layout or behavior cannot be achieved with the existing library APIs. Document the reason briefly.
5. Preserve keyboard accessibility, focus, and responsive behavior using standard controls.

## Component mapping

1. Main toolbar and pane headers: `Flex` / `Space`, `Button`, and appropriate layout primitives. Keep the controls in one compact toolbar; do not create bespoke toolbar widgets.
2. Examples: `Dropdown` with grouped `Menu` items (DOCX / XLSX), without a sidebar for now.
3. Analyze / Generate JSON: `Dropdown.Button` if it fits the required primary-action-plus-menu behavior; main action runs Analyze, arrow menu opens `Generate JSON…`.
4. Generate: standard primary `Button`. Application overflow commands: `Dropdown` + native `Menu` item layout, including API and `Clear Local Data…`; do not add placeholder icons or uneven item extras for decoration.
5. Dataset and Template headers show one active filename as non-interactive text beside the heading/tabs; use `Flex` with native `Typography.Text` ellipsis and keep the dirty indicator in that same truncation-safe group. In the document header, keep the filename immediately after the Template/Result tabs and place stale Result freshness at the far right before the overflow menu. Do not add file selectors. Their overflow menus use `Dropdown` + `Menu`; labels omit redundant context words. Example filenames use `colorTextTertiary`, Local filenames use `colorText`.
6. Template / Result: `Radio.Group` with button-style `Radio.Button` options, or the existing equivalent standard Ant Design radio control. Do not implement custom toggle buttons.
7. Main split panels and expandable diagnostics layout: standard Ant Design `Splitter` / `Splitter.Panel`, including a narrow divider via supported API where possible. Use controlled sizes and native resize/collapse callbacks to persist split sizes and collapsed states; expanding restores the last non-collapsed user size. Keep persisted preferences in Clear Local Data. Avoid an in-house drag implementation.
8. Diagnostics: `Tabs` for `Problems`, `Analysis`, `Output`; their content is independent of the tab strip. Place `Clear` inside the Output tab content, aligned to the right above events.
9. Generate JSON settings: `Modal` with `Form`, `InputNumber` for items per collection, and `Radio.Group` for `Random values` / `Empty values`. Do not include a destination selector. Create a new dataset by default to avoid overwriting active data.
10. Confirmations: use `Modal.confirm` for potentially destructive operations and `Popconfirm` only for simple contextual confirmations. Preserve unsaved-change protection and clarify that clearing local application data does not delete files on disk.
11. Empty states: standard `Empty` and `Upload.Dragger` where appropriate; upload and close actions should be available without requiring an extra closing step before replacement.
12. Status and feedback: `Typography`, `Tag` / `Badge`, and `message` / `notification` where appropriate. Following the Stage 4 Product Owner decision, status shows the latest Output event (action, outcome, start time, duration) and opens Output when clicked; keep Problems count separate. Show compact `Outdated` using the native Ant Design `Tag` with warning color, outlined variant, and `ExclamationCircleOutlined` icon at the far right immediately before the document overflow menu only for a stale active Result, and explain it with a Tooltip. Keep the filename directly after the tabs.
13. Icons: `@ant-design/icons`; do not embed custom icon assets for standard actions.
14. Shortcuts: reuse Monaco commands where available. Handle Dataset Save (`Cmd+S` / `Ctrl+S`) and Save As (`Cmd+Shift+S` / `Ctrl+Shift+S`) throughout the workspace, preventing the browser Save Page action while leaving dialogs and ordinary non-Monaco editable controls alone. Keep Monaco commands available and display actual platform-appropriate shortcuts in menu entries and tooltips.
15. The Keyboard shortcuts Modal uses native `Flex` layout for section headings and command/keycap rows. Keep only the `Workspace` and `Dataset editor` sections, omit descriptions and table chrome, and preserve accessible shortcut names and native Modal focus/Escape behavior.

## Behavior boundaries

1. Keep the last successfully generated document visible when JSON changes or a later generation fails; indicate that the result is outdated.
2. Output records operations, timestamps, outcomes, and durations within the current session; `Clear` clears Output only.
3. File operations should use supported browser APIs with a clear fallback; an HTML mockup does not implement disk access.
4. One active Dataset, one active Template and a separately retained Result are visible workspace concepts; do not expose recent-file selectors. No Projects, new persistent file history, Office Add-in, or Data Sources implementation in this UI feature.
5. Review and safely remove demonstrably unnecessary CSS before layout changes. Remove CSS made obsolete by a component only when that component is replaced; preserve still-used styles until then.
