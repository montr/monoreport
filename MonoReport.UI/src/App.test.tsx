import { fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { ConfigProvider } from "antd";
import { LocalFiles, type LocalFileHandle } from "./services/localFiles";
import { WorkspacePersistence, workspaceStorage, type WorkspaceContext } from "./services/workspacePersistence";
import { testIndexedDb } from "./test/indexedDb";
import examplesRegistry from "../public/examples.json";

vi.mock("./components/MonacoJsonEditor", async () => {
  const { useImperativeHandle } = await import("react");
  return { MonacoJsonEditor: ({ value, onChange, ref, wordWrap, onWordWrapChange }: { value: string; onChange(value: string): void; ref: import("react").Ref<import("./components/MonacoJsonEditor").JsonEditorCommands>; wordWrap: boolean; onWordWrapChange(enabled: boolean): void }) => {
    useImperativeHandle(ref, () => ({
      async format() { onChange(JSON.stringify(JSON.parse(value), null, 2)); },
      async copy() { await navigator.clipboard.writeText(value); },
      toggleWrap() { onWordWrapChange(!wordWrap); }
    }));
    return <div className="monaco-editor"><textarea aria-label="Dataset" data-word-wrap={wordWrap} value={value} onChange={event => onChange(event.target.value)} /></div>;
  } };
});

const previewSources: Array<Blob | File> = [];
vi.mock("./components/DocxPreview", () => ({
  DocxPreview: ({ source, label }: { source: Blob | File; label: string }) => {
    previewSources.push(source);
    return <div>{label} preview</div>;
  }
}));

const xlsxSources: Array<Blob | File> = [];
const previewTestState = vi.hoisted(() => ({ failResultPreview: false }));
vi.mock("./components/XlsxPreview", () => ({
  XlsxPreview: ({ source, label }: { source: Blob | File; label: string }) => {
    xlsxSources.push(source);
    return <div>{previewTestState.failResultPreview && label === "Result" ? "Result XLSX preview failed" : `${label} XLSX preview`}</div>;
  }
}));

const mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const template = () => new File([new Uint8Array([80, 75, 3, 4])], "sample.template.docx", { type: mime });
const manifest = {
  defaultExample: "construction-cost-estimate",
  examples: [{ id: "construction-cost-estimate", title: "Construction cost estimate", format: "docx", templatePath: "/examples/sample.template.docx", dataPath: "/examples/sample.json" }]
};
const sampleData = '{"Project":{"Name":"Riverside"}}';
const xlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const success = (bytes = new Uint8Array([80, 75, 7, 8]), mediaType = mime, filename = mediaType === xlsxMime ? "generated.xlsx" : "generated.docx") => {
  const blob = new Blob([bytes], { type: mediaType });
  return { ok: true, status: 200, blob: async () => blob, headers: new Headers({ "Content-Type": mediaType, "Content-Disposition": `attachment; filename=${filename}` }) } as Response;
};
const analysis = (requirements: unknown[] = [{ name: "Company", type: "object", fields: [{ name: "Name", type: "value" }] }], valid = true) =>
  new Response(JSON.stringify({ valid, errors: valid ? [] : [{ code: "template.invalid_structure", message: "Invalid block" }], warnings: [], requirements, occurrences: [] }), { status: 200, headers: { "Content-Type": "application/json" } });

function mockPlaygroundFetch(api: (input: RequestInfo | URL, init?: RequestInit) => Promise<Response> = async () => { throw new Error("Unexpected API request"); }) {
  const apiMock = vi.fn(api);
  const fetchMock = vi.fn((input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    if (url === "/examples.json") return Promise.resolve(new Response(JSON.stringify(manifest), { status: 200, headers: { "Content-Type": "application/json" } }));
    if (url === "/examples/sample.template.docx") return Promise.resolve(new Response(new Uint8Array([80, 75, 3, 4]), { status: 200, headers: { "Content-Type": mime } }));
    if (url === "/examples/sample.json") return Promise.resolve(new Response(sampleData, { status: 200, headers: { "Content-Type": "application/json" } }));
    return apiMock(input, init);
  });
  vi.stubGlobal("fetch", fetchMock);
  return { apiMock, fetchMock };
}

async function selectTemplate(user: ReturnType<typeof userEvent.setup>, file = template()) {
  await user.upload(document.querySelector('input[hidden][type="file"]') as HTMLInputElement, file);
}

function selectedExample() {
  return document.querySelector(".source .context-name");
}

async function chooseExample(user: ReturnType<typeof userEvent.setup>, title: string) {
  await user.click(screen.getByRole("button", { name: "Examples" }));
  await user.click(await screen.findByRole("menuitem", { name: title }));
}

async function openMenu(user: ReturnType<typeof userEvent.setup>, context: "Dataset" | "Document") {
  await user.click(screen.getByRole("button", { name: `${context} menu` }));
}

async function dataCommand(user: ReturnType<typeof userEvent.setup>, name: string | RegExp) {
  await openMenu(user, "Dataset");
  await user.click(await screen.findByRole(name === "Word wrap" ? "menuitemcheckbox" : "menuitem", { name: name === "Format" ? /^Format/ : name }));
}

async function generateJson(user: ReturnType<typeof userEvent.setup>, count = 10) {
  await user.click(screen.getByRole("button", { name: "Analyze actions" }));
  await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
  const dialog = await screen.findByRole("dialog");
  if (count !== 10) fireEvent.change(within(dialog).getByRole("spinbutton", { name: "Items per collection" }), { target: { value: String(count) } });
  const dirty = Boolean(screen.queryByLabelText("Dataset modified"));
  await user.click(within(dialog).getByRole("button", { name: "OK" }));
  if (dirty) {
    const confirmation = await screen.findByRole("dialog", { name: "Generate dataset?" });
    await user.click(within(confirmation).getByRole("button", { name: "Replace" }));
  }
}

function expectResultOutdated(outdated = true) {
  const indicator = screen.queryByText("Outdated", { exact: true });
  const resultActive = (screen.getByRole("radio", { name: "Result" }) as HTMLInputElement).checked;
  const header = screen.getByLabelText("Preview").querySelector(".pane-header")!;
  const contextName = header.querySelector(".context-name")!;
  const resultTab = header.querySelector(".ant-radio-group")!;
  const menuButton = within(header as HTMLElement).getByRole("button", { name: "Document menu" });
  expect(contextName.previousElementSibling).toBe(resultTab);
  if (outdated && resultActive) {
    expect(indicator).toBeInTheDocument();
    const tag = indicator!.closest(".ant-tag")!;
    expect(tag).toHaveClass("ant-tag-warning", "ant-tag-outlined");
    expect(tag.querySelector(".anticon-exclamation-circle")).toBeInTheDocument();
    expect(contextName.lastElementChild).toHaveTextContent(/\.(docx|xlsx)$/);
    expect(menuButton.previousElementSibling).toContainElement(indicator!);
  } else expect(indicator).not.toBeInTheDocument();
  expect(screen.getByRole("status")).not.toHaveTextContent("Outdated");
}

async function diagnosticsTab(user: ReturnType<typeof userEvent.setup>, name: string) {
  await user.click(screen.queryByRole("button", { name }) ?? screen.getByRole("tab", { name }));
}

async function documentCommand(user: ReturnType<typeof userEvent.setup>, name: string) {
  await openMenu(user, "Document");
  await user.click(await screen.findByRole("menuitem", { name }));
}

async function confirmReplacement(user: ReturnType<typeof userEvent.setup>, accept = true) {
  const dialog = await screen.findByRole("dialog");
  await user.click(within(dialog).getByRole("button", { name: accept ? "Replace" : "Cancel" }));
  await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
}

function localHandle(name: string, text: string) {
  const file = new File([text], name);
  const writable = { write: vi.fn(async (_text: string) => undefined), close: vi.fn(async (): Promise<void> => undefined), abort: vi.fn(async () => undefined) };
  const handle: LocalFileHandle = { name, getFile: vi.fn(async () => file), isSameEntry: vi.fn(async () => false), createWritable: vi.fn(async () => writable) };
  return { handle, file, writable };
}

function nativeFiles(...selections: LocalFileHandle[][]) {
  vi.stubGlobal("isSecureContext", true);
  const picker = vi.fn(async () => selections.shift() ?? []);
  vi.stubGlobal("showOpenFilePicker", picker);
  return picker;
}

async function rememberWorkspace(context: WorkspaceContext, db = testIndexedDb()) {
  vi.stubGlobal("indexedDB", db.factory);
  await new WorkspacePersistence(localStorage, db.factory, new LocalFiles({})).save(context);
  return db;
}

async function clearWorkspaceDialog(user: ReturnType<typeof userEvent.setup>) {
  await user.click(screen.getByRole("button", { name: "Application menu" }));
  await user.click(await screen.findByRole("menuitem", { name: "Clear Local Data…" }));
  return screen.findByRole("dialog", { name: "Clear Local Data?" });
}

describe("MonoReport playground", () => {
  afterEach(() => { vi.unstubAllEnvs(); });

  beforeEach(() => {
    // Native portal Escape handling needs distinct IDs; the library's test mode forces one shared ID.
    vi.stubEnv("NODE_ENV", "development");
    localStorage.clear();
    vi.stubGlobal("indexedDB", undefined);
    vi.stubGlobal("isSecureContext", false);
    vi.stubGlobal("showOpenFilePicker", undefined);
    vi.stubGlobal("showSaveFilePicker", undefined);
    class TestResizeObserver {
      observe() {}
      unobserve() {}
      disconnect() {}
    }
    vi.stubGlobal("ResizeObserver", TestResizeObserver);
    previewSources.length = 0;
    xlsxSources.length = 0;
    previewTestState.failResultPreview = false;
    mockPlaygroundFetch();
  });

  it("replaces the active dataset only after confirming unsaved edits and generates from its replacement", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => success());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.upload(screen.getByLabelText("Upload dataset"), new File(['{"one":1}'], "one.json"));
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"one":1}'));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"dirty":true}' } });
    const replacement = new File(['{"two":2}'], "two.json");
    await user.upload(screen.getByLabelText("Upload dataset"), replacement);
    await confirmReplacement(user, false);
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"dirty":true}');
    expect(selectedExample()).toHaveTextContent("one.json");
    expect(screen.getByLabelText("Dataset modified")).toHaveClass("ant-badge");
    await user.upload(screen.getByLabelText("Upload dataset"), replacement);
    await confirmReplacement(user);
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"two":2}'));
    expect(selectedExample()).toHaveTextContent("two.json");
    expect(screen.queryByLabelText("Dataset modified")).not.toBeInTheDocument();
    expect(screen.queryByRole("combobox")).not.toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    expect(((apiMock.mock.calls[0][1] as RequestInit).body as FormData).get("data")).toBe('{"two":2}');
    await waitFor(() => expect(JSON.parse(localStorage.getItem(workspaceStorage.context)!).datasets).toHaveLength(1));
  });

  it.each([false, true])("shows noninteractive filenames and source tooltips through Example-to-local transitions (native: %s)", async native => {
    const user = userEvent.setup();
    const datasetName = `${"long-dataset-name-".repeat(12)}.json`;
    const templateName = `${"long-template-name-".repeat(12)}.xlsx`;
    const localData = localHandle(datasetName, '{"local":true}');
    const localTemplate = localHandle(templateName, "workbook");
    if (native) nativeFiles([localData.handle], [localTemplate.handle]);
    render(<ConfigProvider theme={{ token: { motion: false, colorText: "#123456", colorTextSecondary: "#789abc", colorTextTertiary: "#456def" } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const datasetHeader = screen.getByLabelText("Input dataset").querySelector(".pane-header")!;
    const documentHeader = screen.getByLabelText("Preview").querySelector(".pane-header")!;
    const dataName = within(datasetHeader as HTMLElement).getByText("sample.json");
    const docName = within(documentHeader as HTMLElement).getByText("sample.template.docx");
    for (const name of [dataName, docName]) {
      expect(name.tagName).toBe("SPAN");
      expect(name.closest("button, [role=combobox], a")).toBeNull();
      expect(name).toHaveStyle({ color: "#456def", minWidth: 0, flex: "0 1 auto" });
      expect(name).toHaveClass("ant-typography-ellipsis");
    }
    expect(datasetHeader).not.toHaveTextContent("Construction cost estimate");
    expect(documentHeader).not.toHaveTextContent("Construction cost estimate");
    await user.hover(dataName);
    expect(await screen.findByRole("tooltip")).toHaveTextContent("Example file");
    await user.unhover(dataName);
    if (native) await dataCommand(user, /^Open…/);
    else await user.upload(screen.getByLabelText("Upload dataset"), localData.file);
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"local":true}'));
    const localDataName = within(datasetHeader as HTMLElement).getByText(datasetName);
    expect(localDataName).toHaveStyle({ color: "#123456" });
    expect(localDataName).toHaveAttribute("title", datasetName);
    expect(docName).toBeInTheDocument();
    await user.hover(localDataName);
    await waitFor(() => expect(screen.getByRole("tooltip")).toHaveTextContent("Local file"));
    await user.unhover(localDataName);
    if (native) await documentCommand(user, "Open…");
    else await selectTemplate(user, localTemplate.file);
    await screen.findByText("Template XLSX preview");
    const localTemplateName = within(documentHeader as HTMLElement).getByText(templateName);
    expect(localTemplateName).toHaveStyle({ color: "#123456" });
    expect(localTemplateName).toHaveAttribute("title", templateName);
    expect(localTemplateName).toHaveClass("ant-typography-ellipsis");
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"local":true}');
    await user.hover(localTemplateName);
    await waitFor(() => expect(screen.getByRole("tooltip")).toHaveTextContent("Local file"));
    await user.unhover(localTemplateName);
    await chooseExample(user, "Construction cost estimate");
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData));
    expect(within(documentHeader as HTMLElement).getByText("sample.template.docx")).toHaveStyle({ color: "#456def" });
    // A local template replaces only the template of the newly loaded Example.
    await selectTemplate(user, localTemplate.file);
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(within(datasetHeader as HTMLElement).getByText("sample.json")).toHaveStyle({ color: "#456def" });
    expect(within(documentHeader as HTMLElement).getByText(templateName)).toHaveStyle({ color: "#123456" });
    expect(screen.queryByRole("combobox")).not.toBeInTheDocument();
  });

  it("groups the dirty indicator with the truncated Dataset filename, away from the menu", async () => {
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const name = screen.getByText("sample.json");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "dirty" } });
    const dot = screen.getByLabelText("Dataset modified");
    expect(dot.parentElement).toHaveClass("context-name");
    expect(dot.parentElement).toContainElement(name);
    expect(dot.parentElement).not.toContainElement(screen.getByRole("button", { name: "Dataset menu" }));
    expect(dot).toHaveClass("ant-badge-status");
    expect(dot).toBeVisible();
    expect(dot.closest(".pane-header")?.lastElementChild).not.toBe(dot);
    expect(name).toHaveClass("ant-typography-ellipsis");
  });

  it("keeps the Dataset editor visible when empty and offers file operations without Close or Clear", async () => {
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const editor = screen.getByLabelText("Dataset");
    fireEvent.change(editor, { target: { value: "" } });
    expect(editor).toBeVisible();
    await openMenu(user, "Dataset");
    expect(screen.queryByRole("menuitem", { name: /Close|Clear/ })).not.toBeInTheDocument();
    expect(screen.getByRole("menuitem", { name: /^Save(?! As)/ })).toBeInTheDocument();
    expect(screen.getByRole("menuitem", { name: /^Save As/ })).toBeInTheDocument();
    expect(screen.queryByRole("menuitem", { name: "Download" })).not.toBeInTheDocument();
    let filename = "";
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) { filename = this.download; });
    await user.click(screen.getByRole("menuitem", { name: /^Save(?! As)/ }));
    expect(await (vi.mocked(URL.createObjectURL).mock.calls.at(-1)![0] as Blob).text()).toBe("");
    expect(filename).toBe("sample.json");
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    expect(screen.queryByText("No dataset open")).not.toBeInTheDocument();
  });

  it("preserves the disk association after cancelling replacement and ignores replacement approval after a newer edit", async () => {
    const user = userEvent.setup();
    const original = localHandle("original.json", '{"original":true}');
    const replacement = localHandle("replacement.json", '{"replacement":true}');
    nativeFiles([original.handle], [replacement.handle], [replacement.handle]);
    const savePicker = vi.fn();
    vi.stubGlobal("showSaveFilePicker", savePicker);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, /^Open…/);
    await waitFor(() => expect(selectedExample()).toHaveTextContent("original.json"));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"dirty":1}' } });
    await dataCommand(user, /^Open…/);
    await confirmReplacement(user, false);
    expect(selectedExample()).toHaveTextContent("original.json");
    await dataCommand(user, /^Save(?! As)/);
    await waitFor(() => expect(original.writable.write).toHaveBeenCalledWith('{"dirty":1}'));
    expect(savePicker).not.toHaveBeenCalled();
    expect(replacement.handle.createWritable).not.toHaveBeenCalled();
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"dirty":2}' } });
    await dataCommand(user, /^Open…/);
    await screen.findByRole("dialog", { name: "Open dataset?" });
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"newer":3}' } });
    await confirmReplacement(user);
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"newer":3}');
    expect(selectedExample()).toHaveTextContent("original.json");
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
  });

  it("opens native JSON handles once by identity and keeps unsaved text when reopening the same file", async () => {
    const user = userEvent.setup();
    const one = localHandle("one.json", '{"original":1}');
    const alias = localHandle("one.json", '{"changedOnDisk":1}');
    one.handle.isSameEntry = vi.fn(async other => other === alias.handle);
    const picker = nativeFiles([one.handle], [alias.handle]);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, /^Open…/);
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"original":1}'));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"unsaved":2}' } });
    await dataCommand(user, /^Open…/);
    await waitFor(() => expect(picker).toHaveBeenCalledTimes(2));
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"unsaved":2}');
    expect(selectedExample()).toHaveTextContent("one.json");
    expect(screen.queryByRole("combobox")).not.toBeInTheDocument();
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
  });

  it("leaves file and text drops over the active editor to Monaco without opening a dataset", async () => {
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const editor = screen.getByLabelText("Dataset");
    const file = new File(['{"dropped":true}'], "native.json");
    const read = vi.spyOn(file, "text");
    const targetDrop = vi.fn(), targetDrag = vi.fn();
    editor.addEventListener("drop", targetDrop);
    editor.addEventListener("dragover", targetDrag);
    for (const types of [["Files"], ["text/plain"]]) {
      for (const type of ["dragover", "drop"]) {
        const event = new Event(type, { bubbles: true, cancelable: true });
        Object.assign(event, { dataTransfer: { files: types.includes("Files") ? [file] : [], types } });
        fireEvent(editor, event);
        expect(event.defaultPrevented).toBe(false);
      }
    }
    expect(targetDrop).toHaveBeenCalledTimes(2);
    expect(targetDrag).toHaveBeenCalledTimes(2);
    expect(read).not.toHaveBeenCalled();
    expect(editor).toHaveValue(sampleData);
    expect(screen.queryByRole("combobox")).not.toBeInTheDocument();
    expect(screen.queryByTitle("native.json")).not.toBeInTheDocument();
  });

  it("restores an empty workspace with a visible editable Dataset instead of an empty-state drop area", async () => {
    await rememberWorkspace({ datasets: [], templates: [] });
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await waitFor(() => expect(screen.getByRole("button", { name: "Examples" })).toBeEnabled());
    const editor = screen.getByLabelText("Dataset");
    await waitFor(() => expect(editor).toHaveValue(""));
    expect(editor).toBeVisible();
    fireEvent.change(editor, { target: { value: '{"typed":true}' } });
    expect(editor).toHaveValue('{"typed":true}');
    expect(selectedExample()).toHaveTextContent("Untitled.json");
    expect(screen.queryByText("No dataset open")).not.toBeInTheDocument();
  });

  it("saves clean and dirty JSON to the native file association without changing result freshness", async () => {
    const user = userEvent.setup();
    const originalText = '{ "disk": 1 }\r\n';
    const original = localHandle("original.json", originalText);
    nativeFiles([original.handle]);
    const savePicker = vi.fn();
    vi.stubGlobal("showSaveFilePicker", savePicker);
    mockPlaygroundFetch(async () => success());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, /^Open…/);
    await waitFor(() => expect(selectedExample()).toHaveTextContent("original.json"));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    await openMenu(user, "Dataset");
    expect(screen.getByRole("menuitem", { name: /^Save(?! As)/ })).toBeInTheDocument();
    expect(screen.getByRole("menuitem", { name: /^Save As/ })).toBeInTheDocument();
    expect(screen.queryByRole("menuitem", { name: "Download" })).not.toBeInTheDocument();
    await user.click(screen.getByRole("menuitem", { name: /^Save(?! As)/ }));
    await waitFor(() => expect(original.writable.write).toHaveBeenCalledWith(originalText));
    expect(screen.queryByLabelText("Dataset modified")).not.toBeInTheDocument();
    expectResultOutdated(false);
    expect(screen.getByRole("status")).toHaveTextContent("Generate · Success");
    const dirtyText = '{ "edited": 2 }\n';
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: dirtyText } });
    await dataCommand(user, /^Save(?! As)/);
    await waitFor(() => expect(original.writable.write).toHaveBeenLastCalledWith(dirtyText));
    expect(selectedExample()).toHaveTextContent(/^original.json$/);
    expectResultOutdated();
    expect(savePicker).not.toHaveBeenCalled();
    await waitFor(() => expect(original.writable.write).toHaveBeenCalledTimes(2));
    expect(screen.queryByLabelText("Dataset modified")).not.toBeInTheDocument();
    expect(savePicker).not.toHaveBeenCalled();
    expectResultOutdated();
  });

  it("uses Save and Save As browser downloads when native file saving is unavailable", async () => {
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"dirty":true}' } });
    const click = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
    await openMenu(user, "Dataset");
    expect(screen.queryByRole("menuitem", { name: "Download" })).not.toBeInTheDocument();
    await user.click(screen.getByRole("menuitem", { name: /^Save(?! As)/ }));
    let blob = vi.mocked(URL.createObjectURL).mock.calls.at(-1)![0] as Blob;
    expect(await blob.text()).toBe('{"dirty":true}');
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    await dataCommand(user, /^Save As/);
    blob = vi.mocked(URL.createObjectURL).mock.calls.at(-1)![0] as Blob;
    expect(await blob.text()).toBe('{"dirty":true}');
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    expect(click).toHaveBeenCalledTimes(2);
    await openMenu(user, "Dataset");
    expect(screen.getByRole("menuitem", { name: /^Save(?! As)/ })).toBeInTheDocument();
  });

  it("requires Overwrite after external timestamp changes, preserves Cancel, and remembers each successful save timestamp", async () => {
    const user = userEvent.setup();
    const original = localHandle("original.json", '{"disk":1}');
    const opened = new File(['{"disk":1}'], "original.json", { lastModified: 100 });
    const changed = new File(['{"external":2}'], "original.json", { lastModified: 200 });
    const saved = new File(['{"edited":3}'], "original.json", { lastModified: 300 });
    vi.mocked(original.handle.getFile).mockResolvedValue(opened);
    nativeFiles([original.handle]);
    vi.stubGlobal("showSaveFilePicker", vi.fn());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, /^Open…/);
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"disk":1}'));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"edited":3}' } });
    vi.mocked(original.handle.getFile).mockResolvedValue(changed);
    await dataCommand(user, /^Save(?! As)/);
    const cancel = await screen.findByRole("dialog");
    expect(cancel).toHaveTextContent("File changed on disk");
    expect(original.handle.createWritable).not.toHaveBeenCalled();
    await user.click(within(cancel).getByRole("button", { name: "Cancel" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(original.writable.write).not.toHaveBeenCalled();
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"edited":3}');
    await dataCommand(user, /^Save(?! As)/);
    const overwrite = await screen.findByRole("dialog");
    vi.mocked(original.handle.getFile).mockResolvedValue(saved);
    await user.click(within(overwrite).getByRole("button", { name: "Overwrite" }));
    await waitFor(() => expect(screen.queryByLabelText("Dataset modified")).not.toBeInTheDocument());
    expect(original.writable.write).toHaveBeenCalledExactlyOnceWith('{"edited":3}');
    await dataCommand(user, /^Save(?! As)/);
    await waitFor(() => expect(original.writable.close).toHaveBeenCalledTimes(2));
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    vi.mocked(original.handle.getFile).mockResolvedValue(changed);
    await dataCommand(user, /^Save(?! As)/);
    expect(await screen.findByRole("dialog")).toHaveTextContent("File changed on disk");
    expect(original.writable.close).toHaveBeenCalledTimes(2);
  });

  it("blocks Save when disk timestamp access fails and protects the current file selected again through Save As", async () => {
    const user = userEvent.setup();
    const original = localHandle("original.json", "{}");
    const alias = localHandle("original.json", "{}");
    original.handle.isSameEntry = vi.fn(async other => other === alias.handle);
    nativeFiles([original.handle]);
    vi.stubGlobal("showSaveFilePicker", vi.fn(async () => alias.handle));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, /^Open…/);
    await waitFor(() => expect(selectedExample()).toHaveTextContent("original.json"));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"keep":1}' } });
    vi.mocked(original.handle.getFile).mockRejectedValueOnce(new DOMException("Access revoked", "NotAllowedError"));
    await dataCommand(user, /^Save(?! As)/);
    await screen.findByText("Access revoked");
    expect(original.handle.createWritable).not.toHaveBeenCalled();
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"keep":1}');
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    vi.mocked(alias.handle.getFile).mockResolvedValue(new File(["{}"], "original.json", { lastModified: original.file.lastModified + 1 }));
    await dataCommand(user, /^Save As…/);
    const dialog = await screen.findByRole("dialog");
    expect(dialog).toHaveTextContent("File changed on disk");
    expect(alias.handle.createWritable).not.toHaveBeenCalled();
    await user.click(within(dialog).getByRole("button", { name: "Cancel" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    vi.mocked(original.handle.getFile).mockRejectedValueOnce(new DOMException("File missing", "NotFoundError"));
    await dataCommand(user, /^Save(?! As)/);
    await screen.findByText("File missing");
    expect(original.handle.createWritable).not.toHaveBeenCalled();
  });

  it("saves the exact snapshot to the original handle, keeps edits during save dirty, and Save As changes the destination", async () => {
    const user = userEvent.setup();
    const original = localHandle("original.json", '{ "value": 1 }\r\n');
    const destination = localHandle("copy.json", "{}");
    nativeFiles([original.handle]);
    const savePicker = vi.fn(async () => destination.handle);
    vi.stubGlobal("showSaveFilePicker", savePicker);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, /^Open…/);
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{ "value": 1 }\n'));
    await dataCommand(user, /^Save(?! As)/);
    await waitFor(() => expect(original.writable.close).toHaveBeenCalledOnce());
    expect(original.writable.write).toHaveBeenCalledWith('{ "value": 1 }\r\n');
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{ "value": 2 }\n' } });
    let finish!: () => void;
    original.writable.close.mockImplementationOnce(() => new Promise<void>(resolve => { finish = resolve; }));
    await dataCommand(user, /^Save(?! As)/);
    await waitFor(() => expect(original.writable.write).toHaveBeenCalledWith('{ "value": 2 }\n'));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"value":3}' } });
    finish();
    await screen.findByText("Saved.");
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    expect(savePicker).not.toHaveBeenCalled();
    await dataCommand(user, /^Save As…/);
    await waitFor(() => expect(destination.writable.close).toHaveBeenCalledOnce());
    expect(destination.writable.write).toHaveBeenCalledWith('{"value":3}');
    expect(screen.queryByLabelText("Dataset modified")).not.toBeInTheDocument();
    expect(selectedExample()).toHaveTextContent("copy.json");
    await dataCommand(user, /^Save(?! As)/);
    await waitFor(() => expect(destination.writable.close).toHaveBeenCalledTimes(2));
    expect(savePicker).toHaveBeenCalledOnce();
    expect(original.writable.write).toHaveBeenCalledTimes(2);
  });

  it("handles workspace Save shortcuts from Format and Monaco while preserving dialog controls", async () => {
    const user = userEvent.setup();
    const original = localHandle("original.json", "{}");
    const destination = localHandle("copy.json", "{}");
    nativeFiles([original.handle]);
    const savePicker = vi.fn(async () => destination.handle);
    vi.stubGlobal("showSaveFilePicker", savePicker);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, /^Open…/);
    await waitFor(() => expect(selectedExample()).toHaveTextContent("original.json"));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"edited":true}' } });

    await openMenu(user, "Dataset");
    const format = await screen.findByRole("menuitem", { name: /^Format/ });
    const save = new KeyboardEvent("keydown", { key: "s", metaKey: true, bubbles: true, cancelable: true });
    format.dispatchEvent(save);
    expect(save.defaultPrevented).toBe(true);
    await waitFor(() => expect(original.writable.write).toHaveBeenCalledWith('{"edited":true}'));
    expect(savePicker).not.toHaveBeenCalled();

    const editor = screen.getByLabelText("Dataset");
    const saveAs = new KeyboardEvent("keydown", { key: "s", ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true });
    editor.dispatchEvent(saveAs);
    expect(saveAs.defaultPrevented).toBe(true);
    await waitFor(() => expect(destination.writable.write).toHaveBeenCalledWith('{"edited":true}'));
    expect(savePicker).toHaveBeenCalledOnce();

    const input = document.createElement("input");
    document.body.append(input);
    const inputSave = new KeyboardEvent("keydown", { key: "s", ctrlKey: true, bubbles: true, cancelable: true });
    input.dispatchEvent(inputSave);
    expect(inputSave.defaultPrevented).toBe(false);
    expect(savePicker).toHaveBeenCalledOnce();
    input.remove();

    await user.click(screen.getByRole("button", { name: "Analyze actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const spinbutton = await screen.findByRole("spinbutton", { name: "Items per collection" });
    const dialogSave = new KeyboardEvent("keydown", { key: "s", ctrlKey: true, bubbles: true, cancelable: true });
    spinbutton.dispatchEvent(dialogSave);
    expect(dialogSave.defaultPrevented).toBe(false);
    expect(savePicker).toHaveBeenCalledOnce();
  });

  it("acquires a destination for unbacked data and preserves work on cancelled, denied, and failed saves", async () => {
    const user = userEvent.setup();
    const destination = localHandle("saved.json", "{}");
    nativeFiles();
    const savePicker = vi.fn().mockRejectedValueOnce(new DOMException("Cancelled", "AbortError")).mockResolvedValue(destination.handle);
    vi.stubGlobal("showSaveFilePicker", savePicker);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{ "keep": true }' } });
    await dataCommand(user, /^Save(?! As)/);
    await waitFor(() => expect(savePicker).toHaveBeenCalledOnce());
    expect(savePicker).toHaveBeenCalledWith(expect.objectContaining({ suggestedName: "sample.json" }));
    expect(screen.queryByText("Dataset could not be saved")).not.toBeInTheDocument();
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    destination.handle.createWritable = vi.fn().mockRejectedValueOnce(new DOMException("Access revoked", "NotAllowedError")).mockResolvedValue(destination.writable);
    await dataCommand(user, /^Save(?! As)/);
    await screen.findByText("Dataset could not be saved");
    expect(selectedExample()).toHaveTextContent("sample.json");
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    destination.writable.write.mockRejectedValueOnce(new Error("Disk full"));
    await dataCommand(user, /^Save(?! As)/);
    await screen.findByText("Disk full");
    expect(destination.writable.abort).toHaveBeenCalledOnce();
    expect(screen.getByLabelText("Dataset")).toHaveValue('{ "keep": true }');
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    await dataCommand(user, /^Save(?! As)/);
    await waitFor(() => expect(selectedExample()).toHaveTextContent("saved.json"));
    expect(screen.queryByLabelText("Dataset modified")).not.toBeInTheDocument();
    await chooseExample(user, "Construction cost estimate");
    await waitFor(() => expect(selectedExample()).toHaveTextContent("sample.json"));
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
  });

  it("keeps active data through cancelled, unreadable, and malformed JSON opens", async () => {
    const user = userEvent.setup();
    const malformed = localHandle("bad.json", "{");
    const missing = localHandle("missing.json", "{}");
    missing.handle.getFile = vi.fn().mockRejectedValue(new DOMException("File missing", "NotFoundError"));
    const picker = nativeFiles();
    picker.mockRejectedValueOnce(new DOMException("Cancelled", "AbortError")).mockResolvedValueOnce([missing.handle]).mockResolvedValueOnce([malformed.handle]);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "keep edits" } });
    await dataCommand(user, /^Open…/);
    expect(screen.queryByText("Dataset could not be opened")).not.toBeInTheDocument();
    await dataCommand(user, /^Open…/);
    await screen.findByText(/File missing/);
    await dataCommand(user, /^Open…/);
    await screen.findByText("Dataset could not be opened");
    expect(screen.getByLabelText("Dataset")).toHaveValue("keep edits");
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
  });

  it("switches reusable native templates and closes one without touching JSON or retained result bytes", async () => {
    const user = userEvent.setup();
    const one = localHandle("one.docx", "one"), two = localHandle("two.xlsx", "two");
    const picker = nativeFiles([one.handle], [two.handle], [one.handle], [two.handle]);
    const response = success();
    const blob = await response.blob();
    mockPlaygroundFetch(async () => response);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await documentCommand(user, "Open…");
    await waitFor(() => expect(previewSources.at(-1)).toBe(one.file));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
    await documentCommand(user, "Open…");
    await waitFor(() => expect(xlsxSources.at(-1)).toBe(two.file));
    expectResultOutdated();
    await documentCommand(user, "Open…");
    await waitFor(() => expect(picker).toHaveBeenCalledTimes(3));
    await waitFor(() => expect(previewSources.at(-1)).toBe(one.file));
    expectResultOutdated(false);
    expect(screen.getByRole("status")).toHaveTextContent("Generate · Success");
    await documentCommand(user, "Open…");
    expect(xlsxSources.at(-1)).toBe(two.file);
    await documentCommand(user, "Close");
    expect(screen.getByText("Select a DOCX or XLSX template")).toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(previewSources.at(-1)).toBe(blob);
    let downloadedFilename = "";
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) { downloadedFilename = this.download; });
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(blob);
    expect(downloadedFilename).toBe("one-result.docx");
  });

  it("activates Template when native Open finishes while Result is selected and when selecting an existing template", async () => {
    const user = userEvent.setup();
    const opened = localHandle("local.xlsx", "template bytes");
    let finish!: (file: File) => void;
    vi.mocked(opened.handle.getFile).mockImplementationOnce(() => new Promise<File>(resolve => { finish = resolve; }));
    nativeFiles([opened.handle]);
    const response = success();
    const blob = await response.blob();
    const { apiMock } = mockPlaygroundFetch(async () => response);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
    await documentCommand(user, "Open…");
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    finish(opened.file);
    await screen.findByText("Template XLSX preview");
    expect(screen.getByRole("radio", { name: "Template" })).toBeChecked();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(previewSources.at(-1)).toBe(blob);

    await selectTemplate(user, new File(["fallback"], "uploaded.docx"));
    expect(screen.getByRole("radio", { name: "Template" })).toBeChecked();
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(previewSources.at(-1)).toBe(blob);
    expect(apiMock).toHaveBeenCalledOnce();
  });

  it("keeps a replacement dataset dirty when a save of the previous dataset finishes", async () => {
    const user = userEvent.setup();
    const one = localHandle("one.json", '{"one":1}'), two = localHandle("two.json", '{"two":2}');
    nativeFiles([one.handle], [two.handle]);
    vi.stubGlobal("showSaveFilePicker", vi.fn());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, /^Open…/);
    await waitFor(() => expect(selectedExample()).toHaveTextContent("one.json"));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"saved":1}' } });
    let finish!: () => void;
    one.writable.close.mockImplementationOnce(() => new Promise<void>(resolve => { finish = resolve; }));
    await dataCommand(user, /^Save(?! As)/);
    await waitFor(() => expect(one.writable.write).toHaveBeenCalledWith('{"saved":1}'));
    await dataCommand(user, /^Open…/);
    await confirmReplacement(user);
    await waitFor(() => expect(selectedExample()).toHaveTextContent("two.json"));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"unsaved":2}' } });
    finish();
    await screen.findByText("Saved.");
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"unsaved":2}');
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    expect(selectedExample()).toHaveTextContent("two.json");
  });

  it("confirms Example replacement of a dirty local Dataset and replaces both resources", async () => {
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.upload(screen.getByLabelText("Upload dataset"), new File(['{"local":true}'], "local.json"));
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"local":true}'));
    await selectTemplate(user, new File(["local"], "local.docx"));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "replace active edits" } });
    await chooseExample(user, "Construction cost estimate");
    await confirmReplacement(user, false);
    expect(screen.getByLabelText("Dataset")).toHaveValue("replace active edits");
    expect(screen.getByText("local.docx")).toBeInTheDocument();
    await chooseExample(user, "Construction cost estimate");
    await confirmReplacement(user);
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData));
    expect(selectedExample()).toHaveTextContent("sample.json");
    expect(screen.getByText("sample.template.docx")).toBeInTheDocument();
    expect(screen.queryByLabelText("Dataset modified")).not.toBeInTheDocument();
  });

  it("loads the default example without generation and exposes native workspace controls and links", async () => {
    const user = userEvent.setup();
    const { fetchMock } = mockPlaygroundFetch();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    expect(await screen.findByText("Template preview")).toBeInTheDocument();
    expect(fetchMock).toHaveBeenCalledWith("/examples.json", { cache: "no-cache" });
    expect(fetchMock).toHaveBeenCalledWith("/examples/sample.template.docx", { cache: "no-cache" });
    expect(fetchMock).toHaveBeenCalledWith("/examples/sample.json", { cache: "no-cache" });
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(selectedExample()).toHaveTextContent("sample.json");
    expect(screen.queryByRole("button", { name: "Form" })).not.toBeInTheDocument();
    expect(screen.getByRole("radio", { name: "Template" })).toBeChecked();
    expect(screen.getByRole("radio", { name: "Result" })).not.toBeChecked();
    expect(screen.getByRole("status")).toHaveTextContent("No operations yet");
    expect(screen.getByRole("link", { name: "MonoReport" })).toHaveAttribute("href", "https://github.com/montr/monoreport");
    expect(screen.getByText("dev")).toBeInTheDocument();
    expect(screen.queryByRole("link", { name: "API documentation" })).not.toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Application menu" }));
    const docs = await screen.findByRole("link", { name: "API documentation" });
    expect(docs).toHaveAttribute("href", "/docs/");
    expect(docs).toHaveAttribute("target", "_blank");
    expect(docs).toHaveAttribute("rel", "noopener noreferrer");
    const applicationMenu = screen.getByRole("menu");
    expect(within(applicationMenu).getAllByRole("menuitem").map(item => item.textContent)).toEqual(["API documentation", "Keyboard shortcuts", "Clear Local Data…"]);
    expect(applicationMenu.querySelectorAll(".ant-dropdown-menu-item-extra")).toHaveLength(0);
    expect(document.querySelector(".ant-splitter-horizontal")).not.toBeNull();
  });

  it("keeps the filename after the tabs and places active Result freshness before the menu", async () => {
    const user = userEvent.setup();
    mockPlaygroundFetch(async () => success());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");

    const header = document.querySelector(".preview .pane-header")!;
    const radioGroup = header.querySelector(".ant-radio-group")!;
    const filenameGroup = header.querySelector(".context-name")!;
    const filename = within(filenameGroup as HTMLElement).getByText("sample.docx");
    const menu = screen.getByRole("button", { name: "Document menu" });
    expect(filenameGroup.previousElementSibling).toBe(radioGroup);
    expect(filenameGroup.lastElementChild).toBe(filename);
    expect(screen.queryByText("Outdated", { exact: true })).not.toBeInTheDocument();

    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "changed" } });
    const outdated = screen.getByText("Outdated", { exact: true });
    expect(filenameGroup.previousElementSibling).toBe(radioGroup);
    expect(filenameGroup.lastElementChild).toBe(filename);
    expect(menu.previousElementSibling).toContainElement(outdated);
    await user.hover(outdated);
    expect(await screen.findByRole("tooltip")).toHaveTextContent("The Dataset or Template changed since this Result was generated.");
    await user.unhover(outdated);

    await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
    expect(screen.queryByText("Outdated", { exact: true })).not.toBeInTheDocument();
    expect(filenameGroup.previousElementSibling).toBe(radioGroup);
    expect(filenameGroup.lastElementChild).toHaveTextContent("sample.template.docx");
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(screen.getByText("Outdated", { exact: true })).toBeInTheDocument();
    expect(filenameGroup.previousElementSibling).toBe(radioGroup);
    expect(filenameGroup.lastElementChild).toBe(filename);
  });

  it("keeps template download disabled without a template and opens the upload picker", async () => {
    const user = userEvent.setup();
    vi.stubGlobal("fetch", vi.fn(async () => new Response("Unavailable", { status: 503 })));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    expect(await screen.findByText("Examples unavailable")).toBeInTheDocument();
    await openMenu(user, "Document");
    expect(screen.getByRole("menuitem", { name: "Download" })).toHaveAttribute("aria-disabled", "true");
    const input = document.querySelector('input[hidden][type="file"]') as HTMLInputElement;
    const click = vi.spyOn(input, "click");
    await user.click(screen.getByRole("menuitem", { name: "Upload…" }));
    expect(click).toHaveBeenCalledOnce();
    await selectTemplate(user);
    await openMenu(user, "Document");
    expect(screen.getByRole("menuitem", { name: "Download" })).not.toHaveAttribute("aria-disabled", "true");
  });

  it("opens a template from the native empty-state drop area without uploading it to the server or replacing JSON", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await documentCommand(user, "Close");
    const book = new File([new Uint8Array([80, 75])], "book.xlsx", { type: xlsxMime });
    fireEvent.dragEnter(screen.getByText("Select a DOCX or XLSX template"), { dataTransfer: { files: [book], types: ["Files"] } });
    expect(document.querySelector(".canvas")).not.toHaveClass("drop-target");
    fireEvent.drop(screen.getByText("Select a DOCX or XLSX template"), { dataTransfer: { files: [book], types: ["Files"] } });
    expect(await screen.findByText("Template XLSX preview")).toBeInTheDocument();
    expect(xlsxSources.at(-1)).toBe(book);
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(apiMock).not.toHaveBeenCalled();
  });

  it("loads an XLSX registry example with format-correct file metadata and preview dispatch", async () => {
    const xlsxExample = examplesRegistry.examples.find(example => example.id === "xlsx-tender")!;
    const otherXlsxExample = examplesRegistry.examples.find(example => example.id === "xlsx-employees")!;
    expect(xlsxExample.format).toBe("xlsx");
    const registry = { defaultExample: xlsxExample.id, examples: [xlsxExample, otherXlsxExample] };
    const fetchMock = vi.fn((input: RequestInfo | URL, _init?: RequestInit) => {
      const url = String(input);
      if (url === "/examples.json") return Promise.resolve(new Response(JSON.stringify(registry), { status: 200 }));
      if (url === xlsxExample.templatePath) return Promise.resolve(new Response(new Uint8Array([80, 75, 3, 4]), { status: 200, headers: { "Content-Type": "application/octet-stream" } }));
      if (url === xlsxExample.dataPath) return Promise.resolve(new Response('{"Title":"Office equipment"}', { status: 200 }));
      return Promise.reject(new Error(`Unexpected request: ${url}`));
    });
    vi.stubGlobal("fetch", fetchMock);

    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);

    expect(await screen.findByText("Template XLSX preview")).toBeInTheDocument();
    expect(selectedExample()).toHaveTextContent(xlsxExample.dataPath.split("/").at(-1)!);
    expect(screen.getByRole("button", { name: "Examples" })).toBeEnabled();
    const user = userEvent.setup();
    await user.click(screen.getByRole("button", { name: "Examples" }));
    expect(await screen.findByRole("menuitem", { name: xlsxExample.title })).toBeInTheDocument();
    expect(screen.getByText("XLSX", { selector: ".ant-dropdown-menu-item-group-title" })).toBeInTheDocument();
    expect(document.querySelector(".ant-dropdown .format-icon-xlsx")).toHaveStyle({ color: "rgb(33, 115, 70)" });
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"Title":"Office equipment"}');
    expect(xlsxSources.at(-1)).toBeInstanceOf(File);
    expect(xlsxSources.at(-1)).toMatchObject({ name: "xlsx-tender.template.xlsx", type: xlsxMime });
    expect(fetchMock).toHaveBeenCalledWith(xlsxExample.templatePath, { cache: "no-cache" });
    expect(fetchMock).toHaveBeenCalledWith(xlsxExample.dataPath, { cache: "no-cache" });
    expect(fetchMock).not.toHaveBeenCalledWith("/api/documents/analyze", expect.anything());
  });

  it("ignores a pending XLSX example after a custom template is selected", async () => {
    const xlsxExample = { id: "xlsx-tender", title: "XLSX tender", format: "xlsx", templatePath: "/examples/xlsx-tender.template.xlsx", dataPath: "/examples/xlsx-tender.json" };
    const registry = { defaultExample: manifest.defaultExample, examples: [...manifest.examples, xlsxExample] };
    let resolveXlsxTemplate!: (response: Response) => void;
    let resolveXlsxData!: (response: Response) => void;
    vi.stubGlobal("fetch", vi.fn((input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/examples.json") return Promise.resolve(new Response(JSON.stringify(registry), { status: 200 }));
      if (url === "/examples/sample.template.docx") return Promise.resolve(new Response(new Uint8Array([80, 75, 3, 4]), { status: 200, headers: { "Content-Type": mime } }));
      if (url === "/examples/sample.json") return Promise.resolve(new Response(sampleData, { status: 200 }));
      if (url === xlsxExample.templatePath) return new Promise<Response>(resolve => { resolveXlsxTemplate = resolve; });
      if (url === xlsxExample.dataPath) return new Promise<Response>(resolve => { resolveXlsxData = resolve; });
      return Promise.reject(new Error(`Unexpected request: ${url}`));
    }));
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await chooseExample(user, "XLSX tender");
    await waitFor(() => expect(resolveXlsxTemplate).toBeTypeOf("function"));

    const custom = new File([new Uint8Array([9])], "custom.xlsx", { type: xlsxMime });
    await selectTemplate(user, custom);
    resolveXlsxTemplate(new Response(new Uint8Array([80, 75, 3, 4]), { status: 200, headers: { "Content-Type": xlsxMime } }));
    resolveXlsxData(new Response('{"stale":true}', { status: 200 }));

    await waitFor(() => expect(screen.getByText("custom.xlsx")).toBeInTheDocument());
    expect(selectedExample()).toHaveTextContent(/sample.json|Untitled.json/);
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(xlsxSources.at(-1)).toBe(custom);
  });

  it("does not let a pending DOCX example replace a custom XLSX selection", async () => {
    let resolveTemplate!: (response: Response) => void;
    let resolveData!: (response: Response) => void;
    vi.stubGlobal("fetch", vi.fn((input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/examples.json") return Promise.resolve(new Response(JSON.stringify(manifest), { status: 200 }));
      if (url === "/examples/sample.template.docx") return new Promise<Response>(resolve => { resolveTemplate = resolve; });
      if (url === "/examples/sample.json") return new Promise<Response>(resolve => { resolveData = resolve; });
      return Promise.reject(new Error(`Unexpected request: ${url}`));
    }));
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await waitFor(() => expect(resolveTemplate).toBeTypeOf("function"));
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "custom.xlsx", { type: xlsxMime }));
    resolveTemplate(new Response(new Uint8Array([80, 75, 3, 4]), { status: 200, headers: { "Content-Type": mime } }));
    resolveData(new Response(sampleData, { status: 200 }));

    await waitFor(() => expect(screen.getByText("custom.xlsx")).toBeInTheDocument());
    expect(screen.getByLabelText("Dataset")).toHaveValue("{}");
    expect(screen.getByText("Template XLSX preview")).toBeInTheDocument();
  });

  it("formats and copies valid JSON while invalid JSON remains unchanged and never requests", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    const editor = screen.getByLabelText("Dataset");
    const clipboard = vi.spyOn(navigator.clipboard, "writeText");
    fireEvent.change(editor, { target: { value: '{"a":1}' } });
    await dataCommand(user, "Format");
    expect(editor).toHaveValue('{\n  "a": 1\n}');
    await openMenu(user, "Dataset");
    const uncheckedWrap = screen.getByRole("menuitemcheckbox", { name: "Word wrap", checked: false });
    expect(uncheckedWrap.querySelector('[data-icon="check"]')).not.toBeInTheDocument();
    await user.click(uncheckedWrap);
    await openMenu(user, "Dataset");
    const checkedWrap = screen.getByRole("menuitemcheckbox", { name: "Word wrap", checked: true });
    expect(checkedWrap.querySelector('.ant-dropdown-menu-item-icon [data-icon="check"]')).toBeInTheDocument();
    expect(checkedWrap.firstElementChild).toHaveClass("ant-dropdown-menu-item-icon");
    expect(checkedWrap.firstElementChild).toHaveAttribute("aria-hidden", "true");
    await user.click(checkedWrap);
    await openMenu(user, "Dataset");
    expect(screen.getByRole("menuitemcheckbox", { name: "Word wrap", checked: false }).querySelector('[data-icon="check"]')).not.toBeInTheDocument();
    const formatItem = screen.getByRole("menuitem", { name: /^Format/ });
    expect(formatItem.querySelector(".ant-dropdown-menu-item-extra")).toHaveTextContent(/Ctrl\+Shift\+I|⇧⌥F|Shift\+Alt\+F/);
    expect(formatItem.querySelector(".ant-dropdown-menu-item-label")).toHaveTextContent(/^Format$/);
    const uploadItem = screen.getByRole("menuitem", { name: /^Upload…/ });
    expect(uploadItem.querySelector(".ant-dropdown-menu-item-extra")).toHaveTextContent(/Ctrl\+O|⌘ O/);
    await user.keyboard("{Escape}");
    await dataCommand(user, "Copy");
    expect(clipboard).toHaveBeenCalledWith('{\n  "a": 1\n}');
    expect(await screen.findByText("Copied to clipboard.")).toBeInTheDocument();
    fireEvent.change(editor, { target: { value: "{" } });
    await dataCommand(user, "Format");
    expect(editor).toHaveValue("{");
    expect(screen.getByText("Invalid JSON")).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(apiMock).not.toHaveBeenCalled();
  });

  it("restores Word wrap across application remounts and persists both states without changing unrelated storage", async () => {
    const user = userEvent.setup();
    const preference = "monoreport.playground.editor.wordWrap";
    localStorage.setItem("unrelated", "keep");
    const mount = () => render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    let view = mount();
    await screen.findByText("Template preview");
    expect(screen.getByLabelText("Dataset")).toHaveAttribute("data-word-wrap", "false");
    await dataCommand(user, "Word wrap");
    expect(localStorage.getItem(preference)).toBe("true");
    view.unmount();
    view = mount();
    await screen.findByText("Template preview");
    expect(screen.getByLabelText("Dataset")).toHaveAttribute("data-word-wrap", "true");
    await openMenu(user, "Dataset");
    await user.click(screen.getByRole("menuitemcheckbox", { name: "Word wrap", checked: true }));
    expect(localStorage.getItem(preference)).toBe("false");
    view.unmount();
    mount();
    await screen.findByText("Template preview");
    expect(screen.getByLabelText("Dataset")).toHaveAttribute("data-word-wrap", "false");
    await openMenu(user, "Dataset");
    expect(screen.getByRole("menuitemcheckbox", { name: "Word wrap", checked: false })).toBeInTheDocument();
    expect(localStorage.getItem("unrelated")).toBe("keep");
    expect(localStorage.getItem(preference)).toBe("false");
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
  });

  it("keeps Word wrap usable if its preference cannot be read or written", async () => {
    const user = userEvent.setup();
    const read = Storage.prototype.getItem, write = Storage.prototype.setItem;
    vi.spyOn(Storage.prototype, "getItem").mockImplementation(function (this: Storage, key) {
      if (key === "monoreport.playground.editor.wordWrap") throw new DOMException("Storage blocked", "SecurityError");
      return read.call(this, key);
    });
    vi.spyOn(Storage.prototype, "setItem").mockImplementation(function (this: Storage, key, value) {
      if (key === "monoreport.playground.editor.wordWrap") throw new DOMException("Storage blocked", "QuotaExceededError");
      write.call(this, key, value);
    });
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, "Word wrap");
    await openMenu(user, "Dataset");
    expect(screen.getByRole("menuitemcheckbox", { name: "Word wrap", checked: true })).toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveAttribute("data-word-wrap", "true");
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
  });

  it.each([false, true])("groups context commands with native dividers and aligned text without decorative document icons (native files: %s)", async native => {
    const user = userEvent.setup();
    if (native) {
      nativeFiles();
      vi.stubGlobal("showSaveFilePicker", vi.fn());
    }
    mockPlaygroundFetch(async () => success());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await openMenu(user, "Dataset");
    const dataMenu = screen.getByRole("menu");
    expect(within(dataMenu).getAllByRole("separator")).toHaveLength(1);
    expect(within(dataMenu).getByRole("menuitem", { name: /^Format/ }).previousElementSibling).toHaveAttribute("role", "separator");
    expect(within(dataMenu).queryByRole("menuitem", { name: /Generate Dataset/ })).not.toBeInTheDocument();
    expect(within(dataMenu).queryByRole("menuitem", { name: "Close" })).not.toBeInTheDocument();
    expect(within(dataMenu).getByRole("menuitem", { name: /^Save(?! As)/ })).toBeInTheDocument();
    expect(within(dataMenu).getByRole("menuitem", { name: /^Save As/ })).toBeInTheDocument();
    expect(within(dataMenu).queryByRole("menuitem", { name: "Download" })).not.toBeInTheDocument();
    const dataItems = [...within(dataMenu).getAllByRole("menuitem"), within(dataMenu).getByRole("menuitemcheckbox")];
    dataItems.forEach(item => expect(item.firstElementChild).toHaveClass("ant-dropdown-menu-item-icon"));
    await user.keyboard("{Escape}");
    await openMenu(user, "Document");
    let menu = screen.getByRole("menu");
    expect(within(menu).getAllByRole("menuitem").map(item => item.textContent)).toEqual([native ? "Open…" : "Upload…", "Download", "Close"]);
    expect(within(menu).getAllByRole("separator")).toHaveLength(1);
    expect(within(menu).getByRole("menuitem", { name: "Close" }).previousElementSibling).toHaveAttribute("role", "separator");
    expect(menu.querySelector("[data-icon]")).not.toBeInTheDocument();
    within(menu).getAllByRole("menuitem").forEach(item => expect(item.firstElementChild).toHaveClass("ant-dropdown-menu-item-icon"));
    await user.keyboard("{Escape}");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    await openMenu(user, "Document");
    menu = screen.getByRole("menu");
    expect(within(menu).getAllByRole("menuitem").map(item => item.textContent)).toEqual(["Download", "Close"]);
    expect(within(menu).getAllByRole("separator")).toHaveLength(1);
    expect(within(menu).getByRole("menuitem", { name: "Close" }).previousElementSibling).toHaveAttribute("role", "separator");
    expect(menu.querySelector("[data-icon]")).not.toBeInTheDocument();
    within(menu).getAllByRole("menuitem").forEach(item => expect(item.firstElementChild).toHaveClass("ant-dropdown-menu-item-icon"));
  });

  it.each([true, false])("closes only the generated result, preserving dirty datasets and templates (active template: %s)", async hasTemplate => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => success());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const originalTemplate = previewSources.at(-1);
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"keep":true}' } });
    if (!hasTemplate) {
      await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
      await documentCommand(user, "Close");
      await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    }
    await documentCommand(user, "Close");
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"keep":true}');
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent("Generate · Success");
    expectResultOutdated(false);
    expect(screen.queryByText("Result preview")).not.toBeInTheDocument();
    expect(screen.queryByText("sample.docx")).not.toBeInTheDocument();
    expect(screen.getByRole("radio", { name: hasTemplate ? "Template" : "Result" })).toBeChecked();
    if (hasTemplate) {
      expect(previewSources.at(-1)).toBe(originalTemplate);
      await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    }
    expect(screen.getByText("Not generated", { selector: ".ant-empty-description" })).toBeInTheDocument();
    await openMenu(user, "Document");
    expect(screen.getByRole("menuitem", { name: "Download" })).toHaveAttribute("aria-disabled", "true");
    expect(screen.getByRole("menuitem", { name: "Close" })).toHaveAttribute("aria-disabled", "true");
    await user.keyboard("{Escape}");
    if (hasTemplate) {
      await user.click(screen.getByRole("button", { name: "Generate document" }));
      await screen.findByText("Result preview");
      expect(apiMock).toHaveBeenCalledTimes(2);
      expect(((apiMock.mock.calls[1][1] as RequestInit).body as FormData).get("data")).toBe('{"keep":true}');
      expectResultOutdated(false);
    expect(screen.getByRole("status")).toHaveTextContent("Generate · Success");
    } else {
      await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
      expect(screen.getByText("Select a DOCX or XLSX template")).toBeInTheDocument();
      expect(apiMock).toHaveBeenCalledOnce();
    }
  });

  it("shows native error feedback when Copy fails and preserves data", async () => {
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    vi.spyOn(navigator.clipboard, "writeText").mockRejectedValueOnce(new Error("Clipboard denied"));
    await dataCommand(user, "Copy");
    expect(await screen.findByText("The JSON could not be copied to the clipboard.")).toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.queryByText("Copied to clipboard.")).not.toBeInTheDocument();
  });

  it.each(["Construction cost estimate", "Second"])("activates Template when choosing %s without generating or losing the retained result", async title => {
    const user = userEvent.setup();
    const { apiMock, fetchMock } = mockPlaygroundFetch(async () => success());
    fetchMock.mockResolvedValueOnce(new Response(JSON.stringify({ ...manifest, examples: [...manifest.examples, { ...manifest.examples[0], id: "second", title: "Second" }] }), { status: 200 }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    const result = previewSources.at(-1);
    await chooseExample(user, title);
    expect(await screen.findByText("Template preview")).toBeInTheDocument();
    expect(screen.getByRole("radio", { name: "Template" })).toBeChecked();
    expect(selectedExample()).toHaveTextContent("sample.json");
    expect(screen.getByRole("button", { name: "Examples" })).toHaveTextContent("Examples");
    expect(apiMock).toHaveBeenCalledTimes(1);
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(screen.getByText("Result preview")).toBeInTheDocument();
    expect(previewSources.at(-1)).toBe(result);
    await openMenu(user, "Document");
    expect(screen.getByRole("menuitem", { name: "Download" })).toBeInTheDocument();
    expect(screen.queryByRole("menuitem", { name: "Hide panel" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Show document" })).not.toBeInTheDocument();
  });

  it("uses picker and drop through one replacement path, preserves JSON, and rejects unsupported files without replacement", async () => {
    const user = userEvent.setup();
    mockPlaygroundFetch(async () => success());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    const editor = screen.getByLabelText("Dataset");
    fireEvent.change(editor, { target: { value: '{"kept":true}' } });
    await selectTemplate(user);
    expect(screen.getByText("sample.template.docx")).toBeInTheDocument();
    const replacement = new File([new Uint8Array([1])], "second.docx", { type: mime });
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [replacement], types: ["Files"] } });
    expect(screen.getByText("second.docx")).toBeInTheDocument();
    expect(editor).toHaveValue('{"kept":true}');
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    expect(screen.getByRole("radio", { name: "Result" })).toBeChecked();
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [new File([new Uint8Array([1])], "bad.pdf")], types: ["Files"] } });
    expect(screen.getByText("Unsupported template format")).toBeInTheDocument();
    expect(screen.getByText("second-result.docx")).toBeInTheDocument();
    await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
    expect(screen.getByText("second.docx")).toBeInTheDocument();
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(screen.getByRole("radio", { name: "Result" })).toBeChecked();
    const third = new File([new Uint8Array([2])], "third.docx", { type: mime });
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [third], types: ["Files"] } });
    expect(screen.getByText("third.docx")).toBeInTheDocument();
    expect(screen.getByRole("radio", { name: "Template" })).toBeChecked();
    expect(screen.getByText("Template preview")).toBeInTheDocument();
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(screen.getByText("Result preview")).toBeInTheDocument();
    expectResultOutdated();
    expect(editor).toHaveValue('{"kept":true}');
  });

  it("selects XLSX by extension for picker/drop, preserves JSON and results, and switches back to DOCX", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => success());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
    const picker = document.querySelector('input[hidden][type="file"]')!;
    expect(picker).toHaveAttribute("accept", expect.stringContaining(".xlsx"));
    const book = new File([new Uint8Array([1])], "tender.XLSX", { type: mime });
    await selectTemplate(user, book);
    expect(screen.getByText("Template XLSX preview")).toBeInTheDocument();
    expect(xlsxSources.at(-1)).toBe(book);
    expect(previewSources.at(-1)).not.toBe(book);
    expectResultOutdated();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    const dropped = new File([new Uint8Array([2])], "second.xlsx");
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [dropped], types: ["Files"] } });
    expect(xlsxSources.at(-1)).toBe(dropped);
    expect(screen.getByText("second.xlsx")).toBeInTheDocument();
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [new File([], "bad.xls")], types: ["Files"] } });
    expect(screen.getByText("Select a .docx or .xlsx file.")).toBeInTheDocument();
    expect(screen.getByText("second.xlsx")).toBeInTheDocument();
    await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
    await selectTemplate(user);
    expect(screen.getByText("Template preview")).toBeInTheDocument();
    expect(screen.queryByText("Template XLSX preview")).not.toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(apiMock).toHaveBeenCalledOnce();
  });

  it("sends generation only on click, prevents duplicates, previews real response bytes, and downloads them", async () => {
    const user = userEvent.setup();
    let resolve!: (response: Response) => void;
    const responseBlob = new Blob([new Uint8Array([80, 75, 7, 8])], { type: mime });
    const { apiMock } = mockPlaygroundFetch(() => new Promise<Response>(done => { resolve = done; }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    expect(apiMock).not.toHaveBeenCalled();
    const generate = screen.getByRole("button", { name: "Generate document" });
    await user.click(generate);
    expect(generate).toBeDisabled();
    await user.click(generate);
    expect(apiMock).toHaveBeenCalledTimes(1);
    resolve({ ok: true, status: 200, blob: async () => responseBlob, headers: new Headers({ "Content-Type": mime, "Content-Disposition": "attachment; filename=generated.docx" }) } as Response);
    await screen.findByText("Result preview");
    expect(screen.getByText("sample.docx")).toBeInTheDocument();
    await user.hover(screen.getByText("sample.docx"));
    expect(await screen.findByRole("tooltip")).toHaveTextContent("Example file");
    await user.unhover(screen.getByText("sample.docx"));
    expect(screen.queryByText("sample.template.docx")).not.toBeInTheDocument();
    expect(previewSources.at(-1)).toBe(responseBlob);
    let downloadName = "";
    const click = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) { downloadName = this.download; });
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(responseBlob);
    expect(click).toHaveBeenCalled();
    expect(downloadName).toBe("sample.docx");
  });

  it("manually generates XLSX without analysis, previews the response format, and keeps the original download after preview failure", async () => {
    const user = userEvent.setup();
    previewTestState.failResultPreview = true;
    const resultBlob = new Blob([new Uint8Array([80, 75, 7, 8, 9])], { type: xlsxMime });
    const { apiMock } = mockPlaygroundFetch(async () => ({ ok: true, status: 200, blob: async () => resultBlob, headers: new Headers({ "Content-Type": `${xlsxMime}; charset=binary`, "Content-Disposition": "attachment; filename=result.xlsx" }) } as Response));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    const sourceFile = new File([new Uint8Array([1, 2])], "manual.xlsx", { type: xlsxMime });
    await selectTemplate(user, sourceFile);
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"manual":true}' } });
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(await screen.findByText("Result XLSX preview failed")).toBeInTheDocument();
    expect(screen.getByRole("radio", { name: "Result" })).toBeChecked();
    expect(xlsxSources.at(-1)).toBe(resultBlob);
    expect(previewSources).not.toContain(resultBlob);
    expect(apiMock).toHaveBeenCalledOnce();
    expect(String((apiMock.mock.calls[0] as [string])[0])).toBe("/api/documents/generate");
    const requestBody = (apiMock.mock.calls[0] as [string, RequestInit])[1].body as FormData;
    expect(requestBody.get("template")).toBe(sourceFile);
    expect(requestBody.get("data")).toBe('{"manual":true}');

    let downloadedFilename = "";
    const click = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) { downloadedFilename = this.download; });
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(resultBlob);
    expect(click).toHaveBeenCalled();
    expect(downloadedFilename).toBe("manual-result.xlsx");
  });

  it("renders real API diagnostics and multiple issues without mock content", async () => {
    const user = userEvent.setup();
    mockPlaygroundFetch(async () => new Response(JSON.stringify({ errors: [{ code: "first", message: "First issue", field: "Amount", dataPath: "Items[3].Amount", part: "/xl/worksheets/sheet2.xml", location: "C14" }, { code: "second", detail: "Second issue" }] }), { status: 422, headers: { "Content-Type": "application/problem+json" } }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime }));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(await screen.findByRole("tab", { name: "Problems 2" })).toBeInTheDocument();
    expect(screen.queryByText("2 problems")).not.toBeInTheDocument();
    expect(screen.getByText(/First issue/)).toBeInTheDocument();
    expect(screen.getByText(/Second issue/)).toBeInTheDocument();
    expect(screen.getByText(/Items\[3\]\.Amount/)).toBeInTheDocument();
    expect(screen.getByText(/\/xl\/worksheets\/sheet2\.xml · C14/)).toBeInTheDocument();
    expect(screen.queryByText("Unmatched collection region")).not.toBeInTheDocument();
  });

  it("retains exact successful bytes after edits, invalid JSON, and API failure while reporting freshness", async () => {
    const user = userEvent.setup();
    const response = success();
    const blob = await response.blob();
    const responses = [response, new Response(JSON.stringify({ title: "Failed", detail: "Incompatible", code: "generation.invalid_data" }), { status: 422, headers: { "Content-Type": "application/problem+json" } })];
    const { apiMock } = mockPlaygroundFetch(async () => responses.shift()!);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "{}" } });
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    expectResultOutdated(false);
    expect(screen.getByRole("status")).toHaveTextContent("Generate · Success");
    await dataCommand(user, "Format");
    expectResultOutdated(false);
    expect(screen.getByRole("status")).toHaveTextContent("Generate · Success");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "{" } });
    expectResultOutdated();
    expect(screen.getByText("Result preview")).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expectResultOutdated();
    expect(screen.getByRole("status")).toHaveTextContent("Generate · Failure");
    expect(screen.getByText("Generate · Invalid JSON")).toBeInTheDocument();
    expect(apiMock).toHaveBeenCalledTimes(1);
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "{}" } });
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(await screen.findByText(/generation.invalid_data/)).toBeInTheDocument();
    expectResultOutdated(false);
    expect(screen.getByRole("status")).toHaveTextContent("Generate · Failure");
    expect(screen.getByRole("radio", { name: "Result" })).toBeChecked();
    expect(previewSources.at(-1)).toBe(blob);
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(blob);
  });

  it("retains XLSX result format and download bytes after switching and closing the template", async () => {
    const user = userEvent.setup();
    const response = success(undefined, xlsxMime);
    const blob = await response.blob();
    mockPlaygroundFetch(async () => response);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user, new File([new Uint8Array([1])], "book.xlsx", { type: xlsxMime }));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result XLSX preview");
    await selectTemplate(user);
    expect(screen.getByRole("radio", { name: "Template" })).toBeChecked();
    expect(screen.getByText("Template preview")).toBeInTheDocument();
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(screen.getByText("book-result.xlsx")).toBeInTheDocument();
    const resultName = screen.getByText("book-result.xlsx");
    await user.hover(resultName);
    expect(await screen.findByRole("tooltip")).toHaveTextContent("Local file");
    await user.unhover(resultName);
    expect(screen.getByText("Result XLSX preview")).toBeInTheDocument();
    expectResultOutdated();
    await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
    await documentCommand(user, "Close");
    expect(screen.getByText("Select a DOCX or XLSX template")).toBeInTheDocument();
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(screen.getByText("Result XLSX preview")).toBeInTheDocument();
    expect(xlsxSources.at(-1)).toBe(blob);
    let downloadedFilename = "";
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) { downloadedFilename = this.download; });
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(blob);
    expect(downloadedFilename).toBe("book-result.xlsx");
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
  });

  it("keeps status and retained result reachable when Problems is collapsed and reopened", async () => {
    const user = userEvent.setup();
    const responses = [success(), new Response("Unavailable", { status: 503 })];
    mockPlaygroundFetch(async () => responses.shift()!);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    const preview = await screen.findByText("Result preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByRole("button", { name: "Hide diagnostics" });
    expectResultOutdated(false);
    expect(screen.getByRole("status")).toHaveTextContent("Generate · Failure");
    await user.click(screen.getByRole("button", { name: "Hide diagnostics" }));
    expect(localStorage.getItem(workspaceStorage.diagnosticsCollapsed)).toBe("true");
    const toggle = screen.getByRole("button", { name: "Problems 1" });
    expect(toggle.querySelector(".ant-badge-count")).toHaveTextContent("1");
    expect(toggle).toHaveAttribute("aria-expanded", "false");
    expect(preview).toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent("Generate · Failure");
    await user.click(toggle);
    expect(localStorage.getItem(workspaceStorage.diagnosticsCollapsed)).toBe("false");
    expect(toggle).not.toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Problems 1" }).querySelector(".ant-badge-count")).toHaveTextContent("1");
  });

  it("restores the diagnostics panel open state from its saved preference", async () => {
    localStorage.setItem(workspaceStorage.diagnosticsCollapsed, "false");
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    expect(screen.getByRole("button", { name: "Hide diagnostics" })).toBeInTheDocument();
  });

  it("ignores an obsolete success after JSON changes during generation", async () => {
    const user = userEvent.setup();
    let resolve!: (response: Response) => void;
    mockPlaygroundFetch(() => new Promise<Response>(done => { resolve = done; }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await selectTemplate(user);
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await user.type(screen.getByLabelText("Dataset"), " ");
    resolve(success());
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Generate · Cancelled"));
  });

  it("keeps the newer XLSX generation when an older request resolves afterward", async () => {
    const user = userEvent.setup();
    const older = success(new Uint8Array([1]), xlsxMime, "older.xlsx");
    const newer = success(new Uint8Array([2]), xlsxMime, "newer.xlsx");
    const resolvers: Array<(response: Response) => void> = [];
    const { apiMock } = mockPlaygroundFetch(() => new Promise<Response>(resolve => resolvers.push(resolve)));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime }));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"revision":2}' } });
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(apiMock).toHaveBeenCalledTimes(2);

    resolvers[1](newer);
    expect(await screen.findByText("Result XLSX preview")).toBeInTheDocument();
    resolvers[0](older);
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Generate · Success"));
    expect(xlsxSources.at(-1)).toBe(await newer.blob());
    expect(screen.getByRole("radio", { name: "Result" })).toBeChecked();
  });

  it("ignores a pending DOCX generation after switching to an XLSX template", async () => {
    const user = userEvent.setup();
    let resolveGeneration!: (response: Response) => void;
    mockPlaygroundFetch(() => new Promise<Response>(resolve => { resolveGeneration = resolve; }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user);
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    const xlsx = new File([new Uint8Array([80, 75])], "replacement.xlsx", { type: xlsxMime });
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [xlsx], types: ["Files"] } });
    resolveGeneration(success());

    await waitFor(() => expect(screen.getByText("replacement.xlsx")).toBeInTheDocument());
    expect(screen.getByText("Template XLSX preview")).toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent("Generate · Cancelled");
  });

  it("ignores an analysis response after the template is cleared", async () => {
    const user = userEvent.setup();
    let resolveAnalysis!: (response: Response) => void;
    mockPlaygroundFetch(() => new Promise<Response>(resolve => { resolveAnalysis = resolve; }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime }));
    const editor = screen.getByLabelText("Dataset");
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    await documentCommand(user, "Close");
    resolveAnalysis(analysis([{ name: "Stale", type: "value" }]));

    await waitFor(() => expect(screen.getByText("Select a DOCX or XLSX template")).toBeInTheDocument());
    expect(editor).toHaveValue(sampleData);
    expect(screen.getByRole("button", { name: "Analyze template" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Analyze actions" })).toBeDisabled();
    expect(screen.getByRole("status")).toHaveTextContent("Analyze · Cancelled");
    expect(screen.getByText("No analysis. Analyze a template to inspect its requirements.")).toBeInTheDocument();
  });

  it("analyzes only on explicit action, adds a dataset, and retains an outdated Result", async () => {
    const user = userEvent.setup();
    const responses = [success(), analysis([
      { name: "Company", type: "object", fields: [{ name: "Name", type: "value" }] },
      { name: "Orders", type: "array", fields: [{ name: "Number", type: "value" }] }
    ])];
    const { apiMock } = mockPlaygroundFetch(async () => responses.shift()!);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await selectTemplate(user);
    expect(apiMock).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");

    await generateJson(user, 1);
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();

    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue(JSON.stringify({ Company: { Name: "" }, Orders: [{ Number: "" }] }, null, 2)));
    expectResultOutdated();
    expect(screen.getByRole("status")).toHaveTextContent("Generate Dataset · Success");
    expect(screen.getByText("Result preview")).toBeInTheDocument();
    const [, analysisInit] = apiMock.mock.calls[1] as [string, RequestInit];
    expect(((analysisInit.body as FormData).get("template") as File).name).toBe("sample.template.docx");
  });

  it("generates an editable JSON scaffold from the original XLSX and nested requirements", async () => {
    const user = userEvent.setup();
    const xlsx = new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime });
    const payload = analysis([
      { name: "Report", type: "object", fields: [
        { name: "Title", type: "value" },
        { name: "Active", type: "boolean" },
        { name: "Rows", type: "array", fields: [{ name: "Amount", type: "value" }] }
      ] },
      { name: "Footer", type: "value" }
    ]);
    const { apiMock } = mockPlaygroundFetch(async () => payload);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user, xlsx);
    const editor = screen.getByLabelText("Dataset");
    fireEvent.change(editor, { target: { value: "" } });
    await generateJson(user, 1);

    await waitFor(() => expect(editor).toHaveValue(JSON.stringify({ Report: { Title: "", Active: false, Rows: [{ Amount: "" }] }, Footer: "" }, null, 2)));
    expect(apiMock).toHaveBeenCalledOnce();
    const [url, init] = apiMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe("/api/documents/analyze");
    expect((init.body as FormData).get("template")).toBe(xlsx);
    expect((init.body as FormData).has("data")).toBe(false);
  });

  it("protects dirty Dataset content when generated data would replace it", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => analysis());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const editor = screen.getByLabelText("Dataset");
    fireEvent.change(editor, { target: { value: "keep me" } });
    await user.click(screen.getByRole("button", { name: "Analyze actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const settings = await screen.findByRole("dialog", { name: "Generate Dataset" });
    await user.click(within(settings).getByRole("button", { name: "OK" }));
    await user.click(within(await screen.findByRole("dialog", { name: "Generate dataset?" })).getByRole("button", { name: "Cancel" }));
    expect(editor).toHaveValue("keep me");
    expect(apiMock).not.toHaveBeenCalled();
    await user.click(within(settings).getByRole("button", { name: "OK" }));
    await user.click(within(await screen.findByRole("dialog", { name: "Generate dataset?" })).getByRole("button", { name: "Replace" }));
    await waitFor(() => expect(editor).toHaveValue(JSON.stringify({ Company: { Name: "" } }, null, 2)));
    expect(screen.queryByRole("combobox")).not.toBeInTheDocument();
  });

  it("never applies stale analysis after JSON or template changes", async () => {
    const user = userEvent.setup();
    const resolvers: Array<(response: Response) => void> = [];
    mockPlaygroundFetch(() => new Promise<Response>(resolve => resolvers.push(resolve)));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await selectTemplate(user);
    const editor = screen.getByLabelText("Dataset");
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    fireEvent.change(editor, { target: { value: "edited during analysis" } });
    resolvers[0](analysis());
    await waitFor(() => expect(editor).toHaveValue("edited during analysis"));

    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    const replacement = new File([new Uint8Array([1])], "replacement.docx", { type: mime });
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [replacement], types: ["Files"] } });
    resolvers[1](analysis([{ name: "Stale", type: "value" }]));
    await waitFor(() => expect(screen.getByText("replacement.docx")).toBeInTheDocument());
    expect(editor).toHaveValue("edited during analysis");
  });

  it("ignores an in-flight XLSX analysis after switching to a DOCX template", async () => {
    const user = userEvent.setup();
    let resolveAnalysis!: (response: Response) => void;
    mockPlaygroundFetch(() => new Promise<Response>(resolve => { resolveAnalysis = resolve; }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime }));
    const editor = screen.getByLabelText("Dataset");
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    const docx = template();
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [docx], types: ["Files"] } });
    resolveAnalysis(analysis([{ name: "Stale XLSX", type: "value" }]));

    await waitFor(() => expect(screen.getByText(docx.name)).toBeInTheDocument());
    expect(editor).toHaveValue(sampleData);
    expect(screen.getByText("Template preview")).toBeInTheDocument();
  });

  it("ignores an in-flight DOCX analysis after switching to an XLSX template", async () => {
    const user = userEvent.setup();
    let resolveAnalysis!: (response: Response) => void;
    mockPlaygroundFetch(() => new Promise<Response>(resolve => { resolveAnalysis = resolve; }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user);
    const editor = screen.getByLabelText("Dataset");
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    const xlsx = new File([new Uint8Array([80, 75])], "replacement.xlsx", { type: xlsxMime });
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [xlsx], types: ["Files"] } });
    resolveAnalysis(analysis([{ name: "Stale DOCX", type: "value" }]));

    await waitFor(() => expect(screen.getByText(xlsx.name)).toBeInTheDocument());
    expect(editor).toHaveValue(sampleData);
    expect(screen.getByText("Template XLSX preview")).toBeInTheDocument();
  });

  it("shows analysis failures without replacing JSON", async () => {
    const user = userEvent.setup();
    mockPlaygroundFetch(async () => new Response(JSON.stringify({ valid: false,
      errors: [{ code: "template.unmatched_brace", field: "{Name", part: "/xl/worksheets/sheet1.xml", location: "B12", dataPath: "Invoice.Name",
        message: "Template braces are not balanced or escaped correctly. Token: «{Name». Nearby source text: «Invoice reference: {Name»." }],
      warnings: [], requirements: [], occurrences: [] }), { status: 200 }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime }));
    await generateJson(user, 1);
    expect(await screen.findByText("Generate Dataset · Error · template.unmatched_brace", { selector: "strong" })).toBeInTheDocument();
    expect(screen.getByText(/Invoice reference: \{Name/, { selector: ".issue p" })).toHaveTextContent("Invoice.Name · /xl/worksheets/sheet1.xml · B12");
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
  });

  it("applies example assets atomically and ignores them after JSON editing during the load", async () => {
    let resolveTemplate!: (response: Response) => void;
    let resolveData!: (response: Response) => void;
    vi.stubGlobal("fetch", vi.fn((input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/examples.json") return Promise.resolve(new Response(JSON.stringify(manifest), { status: 200 }));
      if (url.endsWith(".docx")) return new Promise<Response>(resolve => { resolveTemplate = resolve; });
      return new Promise<Response>(resolve => { resolveData = resolve; });
    }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await waitFor(() => expect(resolveTemplate).toBeTypeOf("function"));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "user edit" } });
    resolveTemplate(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
    resolveData(new Response('{"stale":true}', { status: 200 }));
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue("user edit"));
    expect(screen.getByText("Select a DOCX or XLSX template")).toBeInTheDocument();
  });

  it("ignores a pending example after a custom upload", async () => {
    let resolveTemplate!: (response: Response) => void;
    let resolveData!: (response: Response) => void;
    vi.stubGlobal("fetch", vi.fn((input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/examples.json") return Promise.resolve(new Response(JSON.stringify(manifest), { status: 200 }));
      if (url.endsWith(".docx")) return new Promise<Response>(resolve => { resolveTemplate = resolve; });
      return new Promise<Response>(resolve => { resolveData = resolve; });
    }));
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await waitFor(() => expect(resolveTemplate).toBeTypeOf("function"));
    await selectTemplate(user, new File([new Uint8Array([9])], "custom.docx", { type: mime }));
    resolveTemplate(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
    resolveData(new Response('{"stale":true}', { status: 200 }));
    await waitFor(() => expect(screen.getByText("custom.docx")).toBeInTheDocument());
    expect(selectedExample()).toHaveTextContent(/sample.json|Untitled.json/);
    expect(screen.getByLabelText("Dataset")).toHaveValue("{}");
  });

  it("marks edits, preserves Custom JSON, confirms restoration, and downloads exact template bytes", async () => {
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "byte-for-byte JSON" } });
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    const custom = new File([new Uint8Array([4, 5, 6])], "custom.docx", { type: mime });
    await selectTemplate(user, custom);
    expect(screen.getByLabelText("Dataset")).toHaveValue("byte-for-byte JSON");
    expect(selectedExample()).toHaveTextContent(/sample.json|Untitled.json/);
    const click = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(custom);
    expect(click).toHaveBeenCalled();
    await chooseExample(user, "Construction cost estimate");
    await confirmReplacement(user, false);
    expect(selectedExample()).toHaveTextContent(/sample.json|Untitled.json/);
    expect(screen.getByLabelText("Dataset")).toHaveValue("byte-for-byte JSON");
    await chooseExample(user, "Construction cost estimate");
    await confirmReplacement(user);
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData));
  });

  it("prevents a Generate response from applying after an example commits", async () => {
    let templateLoads = 0;
    let dataLoads = 0;
    let resolveTemplate!: (response: Response) => void;
    let resolveData!: (response: Response) => void;
    let resolveGenerate!: (response: Response) => void;
    vi.stubGlobal("fetch", vi.fn((input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/examples.json") return Promise.resolve(new Response(JSON.stringify(manifest), { status: 200 }));
      if (url.endsWith(".docx")) {
        templateLoads += 1;
        return templateLoads === 1 ? Promise.resolve(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } })) : new Promise<Response>(resolve => { resolveTemplate = resolve; });
      }
      if (url.endsWith(".json")) {
        dataLoads += 1;
        return dataLoads === 1 ? Promise.resolve(new Response(sampleData, { status: 200 })) : new Promise<Response>(resolve => { resolveData = resolve; });
      }
      return new Promise<Response>(resolve => { resolveGenerate = resolve; });
    }));
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user, new File([new Uint8Array([9])], "custom.docx", { type: mime }));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"edited":true}' } });
    await chooseExample(user, "Construction cost estimate");
    await confirmReplacement(user);
    await waitFor(() => expect(resolveTemplate).toBeTypeOf("function"));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    resolveTemplate(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
    resolveData(new Response(sampleData, { status: 200 }));
    await waitFor(() => expect(screen.getByText("sample.template.docx")).toBeInTheDocument());
    resolveGenerate(success());
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Generate · Cancelled"));
  });

  it("does not apply an older example after a rapid selection returns to the current example", async () => {
    const twoExamples = {
      ...manifest,
      examples: [...manifest.examples, { id: "second", title: "Second", format: "docx", templatePath: "/examples/second.template.docx", dataPath: "/examples/second.json" }]
    };
    let resolveTemplate!: (response: Response) => void;
    let resolveData!: (response: Response) => void;
    vi.stubGlobal("fetch", vi.fn((input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/examples.json") return Promise.resolve(new Response(JSON.stringify(twoExamples), { status: 200 }));
      if (url === "/examples/sample.template.docx") return Promise.resolve(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
      if (url === "/examples/sample.json") return Promise.resolve(new Response(sampleData, { status: 200 }));
      if (url.endsWith("second.template.docx")) return new Promise<Response>(resolve => { resolveTemplate = resolve; });
      return new Promise<Response>(resolve => { resolveData = resolve; });
    }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const user = userEvent.setup();
    await chooseExample(user, "Second");
    await waitFor(() => expect(resolveTemplate).toBeTypeOf("function"));
    await chooseExample(user, "Construction cost estimate");
    resolveTemplate(new Response(new Uint8Array([1, 2]), { status: 200, headers: { "Content-Type": mime } }));
    resolveData(new Response('{"second":true}', { status: 200 }));
    await waitFor(() => expect(selectedExample()).toHaveTextContent("sample.json"));
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.getByText("sample.template.docx")).toBeInTheDocument();
  });

  it("shows an active malformed example failure without partially replacing content", async () => {
    vi.stubGlobal("fetch", vi.fn((input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/examples.json") return Promise.resolve(new Response(JSON.stringify(manifest), { status: 200 }));
      if (url.endsWith(".docx")) return Promise.resolve(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
      return Promise.resolve(new Response("{", { status: 200 }));
    }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    expect(await screen.findByText("Example could not be loaded")).toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue("{}");
    expect(screen.getByText("Select a DOCX or XLSX template")).toBeInTheDocument();
  });
  it("places Analyze and Generate Dataset in the toolbar and offers only the two approved settings", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    expect(within(screen.getByRole("banner")).getByText("monoreport", { exact: true })).toBeInTheDocument();
    expect(within(screen.getByRole("banner")).queryByText("Playground")).not.toBeInTheDocument();
    const actions = screen.getByLabelText("Workspace actions");
    expect(within(actions).getAllByRole("button").map(button => button.getAttribute("aria-label"))).toEqual(["Examples", "Analyze template", "Analyze actions", "Generate document", "Application menu"]);
    expect(actions).toHaveClass("ant-flex", "ant-flex-wrap-wrap");
    expect(actions).toHaveStyle({ maxWidth: "100%" });
    const primaryGroup = screen.getByRole("button", { name: "Generate document" }).parentElement!;
    expect(primaryGroup).toHaveClass("ant-flex", "ant-flex-align-center");
    expect(primaryGroup).toContainElement(screen.getByRole("button", { name: "Application menu" }));
    await openMenu(user, "Dataset");
    expect(screen.queryByRole("menuitem", { name: /Generate Dataset/ })).not.toBeInTheDocument();
    await user.keyboard("{Escape}");
    await user.click(screen.getByRole("button", { name: "Analyze actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog");
    const count = within(dialog).getByRole("spinbutton", { name: "Items per collection" });
    expect(count.closest("form")).toHaveClass("ant-form-horizontal");
    expect(within(dialog).getByRole("radio", { name: "Random" })).toBeInTheDocument();
    expect(within(dialog).getByText("Values", { exact: true })).toBeInTheDocument();
    expect(dialog.querySelector(".ant-modal-footer")?.querySelectorAll("button")).toHaveLength(1);
    expect(dialog.querySelector(".ant-modal-footer")).toContainElement(within(dialog).getByRole("button", { name: "OK" }));
    expect(within(dialog).queryByRole("button", { name: "Cancel" })).not.toBeInTheDocument();
    expect(dialog).not.toHaveTextContent("Creates a starting point");
    expect(count).toHaveValue("10");
    expect(count).not.toHaveAttribute("aria-valuemax");
    expect(within(dialog).getByRole("radio", { name: "Empty" })).toBeChecked();
    expect(within(dialog).getAllByRole("radio")).toHaveLength(2);
    expect(within(dialog).queryByText("Destination")).not.toBeInTheDocument();
    await user.click(within(dialog).getByRole("button", { name: "Close" }));
    expect(apiMock).not.toHaveBeenCalled();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.getByRole("status")).toHaveTextContent("No operations yet");
  });

  it("does not start generation when settings are cancelled during form validation", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => analysis());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Analyze actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog");
    fireEvent.click(within(dialog).getByRole("button", { name: "OK" }));
    fireEvent.click(within(dialog).getByRole("button", { name: "Close" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).not.toHaveBeenCalled();
    expect(screen.getByRole("status")).toHaveTextContent("No operations yet");
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
  });

  it("rejects invalid settings without a request or dataset and records validation failures", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Analyze actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog");
    const count = within(dialog).getByRole("spinbutton", { name: "Items per collection" });
    for (const value of ["0", "-2", "1.5", ""]) {
      fireEvent.change(count, { target: { value } });
      await user.click(within(dialog).getByRole("button", { name: "OK" }));
      expect(await within(dialog).findByText("Enter a positive integer.")).toBeInTheDocument();
      expect(count).toHaveValue(value);
    }
    expect(apiMock).not.toHaveBeenCalled();
    await user.click(within(dialog).getByRole("button", { name: "Close" }));
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.getByRole("status")).toHaveTextContent("Generate Dataset · Failure");
    await diagnosticsTab(user, "Output");
    const rows = within(screen.getByRole("tabpanel", { name: "Output" })).getAllByRole("listitem");
    expect(rows).toHaveLength(4);
    rows.forEach(row => {
      expect(row).toHaveTextContent("Generate Dataset");
      expect(within(row).getByRole("img", { name: "Failure" })).toBeInTheDocument();
    });
  });

  it("renders actual nested analysis, occurrences and warnings, then reuses it across JSON edits and dataset creation", async () => {
    const user = userEvent.setup();
    const payload = {
      valid: true, errors: [],
      warnings: [{ code: "template.notice", message: "A supplied template warning", field: "Rows", part: "/xl/worksheets/sheet1.xml", location: "C4", formatter: "format", dataPath: "$.Rows" }],
      requirements: [{ name: "Rows", type: "array", fields: [{ name: "Name", type: "value" }, { name: "Enabled", type: "boolean" }, { name: "Tags", type: "array" }] }],
      occurrences: [{ kind: "for", expression: "{For:Rows}", path: "$.Rows", part: "/xl/worksheets/sheet1.xml", location: "C4" }]
    };
    const { apiMock } = mockPlaygroundFetch(async () => new Response(JSON.stringify(payload), { status: 200 }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    const panel = await screen.findByRole("tabpanel", { name: "Analysis" });
    expect(await within(panel).findByText("Valid", { exact: true })).toBeInTheDocument();
    for (const text of ["Rows · array", "Name · value", "Enabled · boolean", "Tags · array", "{For:Rows}", "A supplied template warning", "format"]) {
      expect(within(panel).getByText(text, { exact: true })).toBeInTheDocument();
    }
    expect(within(panel).getAllByText("C4")).toHaveLength(2);
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.queryByText("Result preview")).not.toBeInTheDocument();
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"keep":"dirty"}' } });
    await user.click(screen.getByRole("button", { name: "Analyze actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog");
    fireEvent.change(within(dialog).getByRole("spinbutton", { name: "Items per collection" }), { target: { value: "101" } });
    await user.click(within(dialog).getByRole("radio", { name: "Random" }));
    await user.click(within(dialog).getByRole("button", { name: "OK" }));
    await user.click(within(await screen.findByRole("dialog", { name: "Generate dataset?" })).getByRole("button", { name: "Replace" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    const data = JSON.parse((screen.getByLabelText("Dataset") as HTMLTextAreaElement).value);
    expect(data.Rows).toHaveLength(101);
    data.Rows.forEach((row: { Name: string; Enabled: boolean; Tags: unknown[] }) => {
      expect(row.Name).toMatch(/^Sample \d+$/);
      expect(typeof row.Enabled).toBe("boolean");
      expect(row.Tags).toEqual([]);
    });
    expect(apiMock).toHaveBeenCalledOnce();
    await generateJson(user, 2);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).toHaveBeenCalledOnce();
    expect(JSON.parse((screen.getByLabelText("Dataset") as HTMLTextAreaElement).value).Rows).toHaveLength(2);
    expect(selectedExample()).toHaveTextContent("sample-generated.json");
    await diagnosticsTab(user, "Problems 1");
    expect(await screen.findByText("Generate Dataset · Warning · template.notice", { exact: true })).toBeInTheDocument();
  });

  it("preserves structured invalid analysis and its warning/occurrence details, and never caches it for Generate Dataset", async () => {
    const user = userEvent.setup();
    const payload = {
      valid: false,
      errors: [{ code: "template.bad", message: "Bad template block", field: "{Name", dataPath: "$.Name", part: "/word/document.xml", location: "paragraph 8" }],
      warnings: [{ code: "template.warning", message: "Additional warning" }],
      requirements: [{ name: "Name", type: "value" }],
      occurrences: [{ kind: "field", expression: "{Name", path: "$.Name", part: "/word/document.xml", location: "paragraph 8" }]
    };
    let first = true;
    const { apiMock } = mockPlaygroundFetch(async () => {
      if (first) { first = false; return analysis(); }
      return new Response(JSON.stringify(payload), { status: 200 });
    });
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    expect(await screen.findByText("Company · object")).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    const panel = await screen.findByRole("tabpanel", { name: "Analysis" });
    expect(await within(panel).findByText("Invalid", { exact: true })).toBeInTheDocument();
    expect(within(panel).getByText("Name · value")).toBeInTheDocument();
    expect(within(panel).getByText("Additional warning")).toBeInTheDocument();
    expect(within(panel).getAllByText("paragraph 8")).toHaveLength(2);
    expect(screen.getByRole("status")).toHaveTextContent("Analyze · Failure");
    await generateJson(user);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).toHaveBeenCalledTimes(3);
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.getByRole("status")).toHaveTextContent("Generate Dataset · Failure");
    expect(await screen.findByText("Generate Dataset · Error · template.bad")).toBeInTheDocument();
    expect(screen.getByText("Generate Dataset · Warning · template.warning")).toBeInTheDocument();
  });

  it("clears analysis when the template changes and performs a fresh analysis for new data", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => analysis([{ name: "Rows", type: "array", fields: [{ name: "Name", type: "value" }] }]));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    expect(await screen.findByText("Rows · array")).toBeInTheDocument();
    const replacement = new File([new Uint8Array([80, 75])], "new.xlsx", { type: xlsxMime });
    await selectTemplate(user, replacement);
    expect(screen.queryByText("Rows · array")).not.toBeInTheDocument();
    expect(screen.getByText("No analysis. Analyze a template to inspect its requirements.")).toBeInTheDocument();
    await generateJson(user);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(JSON.parse((screen.getByLabelText("Dataset") as HTMLTextAreaElement).value).Rows).toEqual(Array.from({ length: 10 }, () => ({ Name: "" })));
    expect(selectedExample()).toHaveTextContent("new-generated.json");
    expect(apiMock).toHaveBeenCalledTimes(2);
    expect(((apiMock.mock.calls[1][1] as RequestInit).body as FormData).get("template")).toBe(replacement);
  });

  it("rejects aggregate overflow atomically, retains cached analysis, and allows a smaller complete dataset", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => analysis([{ name: "Rows", type: "array", fields: [{ name: "Items", type: "array", fields: [{ name: "Name", type: "value" }] }] }]));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"keep":true}' } });
    await generateJson(user, 316);
    expect(await screen.findByText(/would exceed 100,000 array elements in total/)).toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"keep":true}');
    expect(selectedExample()).toHaveTextContent("sample.json");
    expect(screen.getByRole("status")).toHaveTextContent("Generate Dataset · Failure");
    expect(selectedExample()).not.toHaveTextContent(/^Generated/);
    await generateJson(user, 2);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).toHaveBeenCalledOnce();
    expect(JSON.parse((screen.getByLabelText("Dataset") as HTMLTextAreaElement).value)).toEqual({ Rows: [{ Items: [{ Name: "" }, { Name: "" }] }, { Items: [{ Name: "" }, { Name: "" }] }] });
    expect(selectedExample()).toHaveTextContent("sample-generated.json");
  });

  it("records chronological measured events, includes internal analysis in Generate Dataset, and clears only Output", async () => {
    const user = userEvent.setup();
    let now = 1000;
    vi.spyOn(performance, "now").mockImplementation(() => now);
    const resolvers: Array<(response: Response) => void> = [];
    const { apiMock } = mockPlaygroundFetch(() => new Promise<Response>(resolve => resolvers.push(resolve)));
    const view = render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await generateJson(user, 1);
    expect(screen.getByRole("status")).toHaveTextContent("Generate Dataset · Running");
    now = 1125;
    resolvers[0](analysis());
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(screen.getByRole("status")).toHaveTextContent(/Generate Dataset · Success · .* · 125 ms/);
    now = 2000;
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    now = 2290;
    resolvers[1](analysis());
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent(/Analyze · Success · .* · 290 ms/));
    now = 3000;
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    now = 3450;
    resolvers[2](success());
    await screen.findByText("Result preview");
    expect(screen.getByRole("status")).toHaveTextContent(/Generate · Success · .* · 450 ms/);
    expect(apiMock.mock.calls.map(([url]) => String(url))).toEqual(["/api/documents/analyze", "/api/documents/analyze", "/api/documents/generate"]);
    await user.click(screen.getByRole("status"));
    let panel = screen.getByRole("tabpanel", { name: "Output" });
    const rows = within(panel).getAllByRole("listitem");
    expect(rows).toHaveLength(3);
    ["Generate Dataset", "Analyze", "Generate"].forEach((action, index) => {
      expect(within(rows[index]).getByText(action, { exact: true })).toBeInTheDocument();
      expect(within(rows[index]).getByRole("img", { name: "Success" })).toBeInTheDocument();
      expect(rows[index]).toHaveTextContent(["125 ms", "290 ms", "450 ms"][index]);
      expect(rows[index].querySelector("time")).toHaveAttribute("datetime", expect.stringMatching(/^\d{4}-\d\d-\d\dT/));
    });
    expect(within(panel).queryByRole("table")).not.toBeInTheDocument();
    expect(panel).not.toHaveTextContent("Company");
    expect(panel).not.toHaveTextContent("Riverside");
    expect(panel).not.toContainElement(screen.getByRole("button", { name: "Clear" }));
    expect(screen.getByRole("toolbar", { name: "Diagnostics actions" })).toContainElement(screen.getByRole("button", { name: "Clear" }));
    const savedStatus = screen.getByRole("status").textContent;
    const savedText = (screen.getByLabelText("Dataset") as HTMLTextAreaElement).value;
    const savedResult = previewSources.at(-1);
    await user.click(screen.getByRole("button", { name: "Clear" }));
    expect(within(panel).queryAllByRole("listitem")).toHaveLength(0);
    expect(screen.getByRole("status")).toHaveTextContent(savedStatus!);
    expect(screen.getByLabelText("Dataset")).toHaveValue(savedText);
    expect(previewSources.at(-1)).toBe(savedResult);
    await diagnosticsTab(user, "Analysis");
    expect(screen.getByText("Company · object")).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Hide diagnostics" }));
    expect(screen.getByRole("status")).toHaveTextContent(savedStatus!);
    await user.click(screen.getByRole("status"));
    panel = screen.getByRole("tabpanel", { name: "Output" });
    expect(within(panel).queryAllByRole("listitem")).toHaveLength(0);
    view.unmount();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    expect(screen.getByRole("status")).toHaveTextContent("No operations yet");
    await diagnosticsTab(user, "Output");
    expect(within(screen.getByRole("tabpanel", { name: "Output" })).queryAllByRole("listitem")).toHaveLength(0);
  });

  it("finishes cancelled work even when the transport ignores abort, and Clear preserves diagnostics and running status", async () => {
    const user = userEvent.setup();
    const resolvers: Array<(response: Response) => void> = [];
    mockPlaygroundFetch(() => new Promise<Response>(resolve => resolvers.push(resolve)));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await generateJson(user);
    const dialog = screen.getByRole("dialog");
    await user.click(within(dialog).getByRole("button", { name: "Close" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(screen.getByRole("status")).toHaveTextContent("Generate Dataset · Cancelled");
    resolvers[0](analysis([{ name: "Obsolete", type: "value" }]));
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    resolvers[1](analysis([], false));
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Analyze · Failure"));
    await diagnosticsTab(user, "Output");
    expect(within(screen.getByRole("tabpanel", { name: "Output" })).getAllByRole("listitem")).toHaveLength(2);
    await user.click(screen.getByRole("button", { name: "Clear" }));
    await diagnosticsTab(user, "Problems 1");
    expect(screen.getByText("Analyze · Error · template.invalid_structure")).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    await diagnosticsTab(user, "Output");
    await user.click(screen.getByRole("button", { name: "Clear" }));
    expect(screen.getByRole("status")).toHaveTextContent("Analyze · Running");
    resolvers[2](analysis());
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Analyze · Success"));
    expect(within(screen.getByRole("tabpanel", { name: "Output" })).queryAllByRole("listitem")).toHaveLength(0);
  });

  it("keeps tabs, Clear-first actions and status outside the active content scroll host, including empty tabs", async () => {
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const footer = screen.getByRole("contentinfo");
    const problems = within(footer).getByRole("button", { name: "Problems" });
    expect(problems.querySelector(".ant-badge")).not.toBeInTheDocument();
    expect(within(footer).getByRole("button", { name: "Analysis" })).toBeInTheDocument();
    expect(within(footer).getByRole("button", { name: "Output" })).toBeInTheDocument();
    await user.click(problems);
    for (const name of ["Problems", "Analysis", "Output"]) expect(within(footer).queryByRole("button", { name })).not.toBeInTheDocument();
    expect(within(footer).getByRole("status")).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Problems" }).querySelector(".ant-badge")).not.toBeInTheDocument();
    const panel = screen.getByRole("tabpanel", { name: "Problems" });
    expect(panel).toHaveStyle({ height: "100%", overflow: "auto" });
    const tabs = screen.getByRole("tablist");
    expect(panel).not.toContainElement(tabs);
    expect(tabs.closest(".ant-tabs")).toHaveStyle({ height: "100%" });
    const toolbar = screen.getByRole("toolbar", { name: "Diagnostics actions" });
    expect(panel).not.toContainElement(toolbar);
    expect(toolbar.parentElement).toContainElement(tabs);
    expect(toolbar.parentElement).toHaveStyle({ flex: "0 0 auto" });
    expect(toolbar).toHaveStyle({ minHeight: "24px", display: "flex" });
    expect(toolbar.compareDocumentPosition(panel) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(tabs.compareDocumentPosition(toolbar) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(within(toolbar).getAllByRole("button").map(button => button.textContent)).toEqual(["Clear", "Copy"]);
    expect(panel).not.toContainElement(footer);
    expect(tabs.closest(".ant-tabs")).not.toContainElement(footer);
    const copy = within(toolbar).getByRole("button", { name: "Copy" });
    expect(copy).toBeDisabled();
    const problemsClear = within(toolbar).getByRole("button", { name: "Clear" });
    expect(problemsClear).toBeDisabled();
    expect(problemsClear.parentElement).toBe(copy.parentElement);
    expect(problemsClear).toHaveClass("ant-btn-sm");
    expect(copy).toHaveClass("ant-btn-sm");
    expect(copy.parentElement).toHaveClass("ant-flex-justify-flex-start");
    await diagnosticsTab(user, "Output");
    const outputPanel = screen.getByRole("tabpanel", { name: "Output" });
    expect(outputPanel).toHaveStyle({ height: "100%", overflow: "auto" });
    expect(screen.getByRole("toolbar", { name: "Diagnostics actions" })).toBe(toolbar);
    expect(outputPanel).not.toContainElement(toolbar);
    expect(within(toolbar).getAllByRole("button").map(button => button.textContent)).toEqual(["Clear"]);
    const clear = within(toolbar).getByRole("button", { name: "Clear" });
    expect(clear).toBeDisabled();
    expect(clear).toHaveClass("ant-btn-sm");
    expect(clear.parentElement).toHaveClass("ant-flex-justify-flex-start");
    expect(screen.queryByRole("button", { name: "Copy" })).not.toBeInTheDocument();
    await diagnosticsTab(user, "Analysis");
    expect(screen.getByRole("toolbar", { name: "Diagnostics actions" })).toBe(toolbar);
    expect(toolbar).toHaveStyle({ minHeight: "24px", display: "flex" });
    expect(within(toolbar).queryAllByRole("button")).toHaveLength(0);
    expect(screen.getByRole("tabpanel", { name: "Analysis" })).not.toContainElement(toolbar);
    await user.click(screen.getByRole("button", { name: "Hide diagnostics" }));
    for (const name of ["Problems", "Analysis", "Output"]) expect(within(footer).getByRole("button", { name })).toBeInTheDocument();
  });

  it("clears only current Problems while retaining analysis, cached warnings, output, data and successful result bytes", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch();
    const response = success();
    const resultBlob = await response.blob();
    const warning = { code: "template.notice", message: "Review this field", part: "/word/document.xml", location: "paragraph 2" };
    apiMock.mockResolvedValueOnce(response).mockResolvedValueOnce(new Response(JSON.stringify({ valid: true, errors: [], warnings: [warning], requirements: [], occurrences: [] })));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    await screen.findByText("Valid", { exact: true });
    await diagnosticsTab(user, "Problems 1");
    const panel = screen.getByRole("tabpanel", { name: "Problems 1" });
    const actions = screen.getByRole("toolbar", { name: "Diagnostics actions" });
    const buttons = within(actions).getAllByRole("button");
    expect(buttons.map(button => button.textContent)).toEqual(["Clear", "Copy"]);
    buttons.forEach(button => { expect(button).toBeEnabled(); expect(button).toHaveClass("ant-btn-sm"); });
    await user.click(buttons[0]);
    expect(within(actions).getAllByRole("button")).toEqual(buttons);
    buttons.forEach(button => expect(button).toBeDisabled());
    expect(await within(panel).findByText("No problems")).toBeInTheDocument();
    expect(within(screen.getByRole("toolbar", { name: "Diagnostics actions" })).getByRole("button", { name: "Copy" })).toBeDisabled();
    expect(within(screen.getByRole("toolbar", { name: "Diagnostics actions" })).getByRole("button", { name: "Clear" })).toBeDisabled();
    expect(screen.getByRole("tab", { name: "Problems" }).querySelector(".ant-badge")).not.toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.getByText("Result preview")).toBeInTheDocument();
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(resultBlob);
    await diagnosticsTab(user, "Analysis");
    expect(screen.getByText("Valid", { exact: true })).toBeInTheDocument();
    expect(screen.getByText("Review this field")).toBeInTheDocument();
    await diagnosticsTab(user, "Output");
    expect(within(screen.getByRole("tabpanel", { name: "Output" })).getAllByRole("listitem")).toHaveLength(2);
    expect(screen.getByRole("status")).toHaveTextContent("Analyze · Success");
    // A new operation publishes its diagnostics normally, even when it reuses cached analysis.
    await generateJson(user);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).toHaveBeenCalledTimes(2);
    await diagnosticsTab(user, "Problems 1");
    expect(screen.getByText("Generate Dataset · Warning · template.notice")).toBeInTheDocument();
    expect(within(screen.getByRole("tabpanel", { name: "Problems 1" })).getByText(/Review this field/)).toBeInTheDocument();
  });

  it("copies every current diagnostic with source metadata and retains diagnostics when clipboard access fails", async () => {
    const user = userEvent.setup();
    const payload = {
      valid: false,
      errors: [{ code: "template.bad", message: "Bad expression", field: "Amount", dataPath: "$.Rows.Amount", part: "/xl/worksheets/sheet1.xml", location: "C4", formatter: "currency" }],
      warnings: [{ code: "template.notice", message: "Review this field", part: "/xl/worksheets/sheet2.xml", location: "D5" }],
      requirements: [], occurrences: []
    };
    mockPlaygroundFetch(async () => new Response(JSON.stringify(payload), { status: 200 }));
    const clipboard = vi.spyOn(navigator.clipboard, "writeText").mockResolvedValue(undefined);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    await screen.findByText("Invalid", { exact: true });
    await diagnosticsTab(user, "Problems 2");
    const panel = screen.getByRole("tabpanel", { name: "Problems 2" });
    await user.click(within(screen.getByRole("toolbar", { name: "Diagnostics actions" })).getByRole("button", { name: "Copy" }));
    expect(await screen.findByText("Diagnostics copied to clipboard.")).toBeInTheDocument();
    expect(JSON.parse(clipboard.mock.calls[0][0])).toEqual([
      { title: "template.bad", code: "template.bad", detail: "Bad expression", field: "Amount", dataPath: "$.Rows.Amount", part: "/xl/worksheets/sheet1.xml", location: "C4", formatter: "currency", action: "Analyze", severity: "Error" },
      { title: "template.notice", code: "template.notice", detail: "Review this field", part: "/xl/worksheets/sheet2.xml", location: "D5", action: "Analyze", severity: "Warning" }
    ]);
    clipboard.mockRejectedValueOnce(new DOMException("Denied", "NotAllowedError"));
    await user.click(within(screen.getByRole("toolbar", { name: "Diagnostics actions" })).getByRole("button", { name: "Copy" }));
    expect(await screen.findByText("The diagnostics could not be copied to the clipboard.")).toBeInTheDocument();
    expect(within(panel).getByText("Analyze · Error · template.bad")).toBeInTheDocument();
    expect(within(panel).getByText("Analyze · Warning · template.notice")).toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    await diagnosticsTab(user, "Output");
    expect(within(screen.getByRole("tabpanel", { name: "Output" })).getAllByRole("listitem")).toHaveLength(1);
  });

  it("uses the same native outcome icons in status and compact event rows, with final durations only", async () => {
    const user = userEvent.setup();
    let resolve!: (response: Response) => void;
    let signal: AbortSignal | undefined;
    let call = 0;
    const good = success();
    const resultBlob = await good.blob();
    mockPlaygroundFetch(async (_input, init) => {
      if (call++ === 0) return good;
      if (call === 2) return analysis([], false);
      signal = init?.signal as AbortSignal;
      return new Promise<Response>(done => { resolve = done; });
    });
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    expect(within(screen.getByRole("status")).getByRole("img", { name: "Success" }).querySelector("svg")).toHaveAttribute("data-icon", "check-circle");
    await user.click(screen.getByRole("button", { name: "Analyze template" }));
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Analyze · Failure"));
    expect(within(screen.getByRole("status")).getByRole("img", { name: "Failure" }).querySelector("svg")).toHaveAttribute("data-icon", "close-circle");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    const status = screen.getByRole("status");
    expect(within(status).getByRole("img", { name: "Running" })).toHaveClass("anticon-spin");
    expect(within(status).getByRole("img", { name: "Running" }).querySelector("svg")).toHaveAttribute("data-icon", "loading");
    expect(status).not.toHaveTextContent(/ms|%/);
    await user.click(status);
    const outputPanel = screen.getByRole("tabpanel", { name: "Output" });
    const runningRow = within(outputPanel).getAllByRole("listitem").at(-1)!;
    expect(within(runningRow).getByRole("img", { name: "Running" }).querySelector("svg")).toHaveAttribute("data-icon", "loading");
    expect(runningRow).not.toHaveTextContent(/ms|%/);
    expect(within(outputPanel).queryByRole("table")).not.toBeInTheDocument();
    expect(within(outputPanel).queryByRole("progressbar")).not.toBeInTheDocument();
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: sampleData + " " } });
    expect(signal?.aborted).toBe(true);
    expect(within(status).getByRole("img", { name: "Cancelled" }).querySelector("svg")).toHaveAttribute("data-icon", "minus-circle");
    expect(status).toHaveTextContent(/\d+ ms/);
    const rows = within(outputPanel).getAllByRole("listitem");
    ["Success", "Failure", "Cancelled"].forEach((outcome, index) => expect(within(rows[index]).getByRole("img", { name: outcome }).querySelector("svg")).toHaveAttribute("data-icon", ["check-circle", "close-circle", "minus-circle"][index]));
    resolve(success());
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Generate · Cancelled"));
    expect(previewSources.at(-1)).toBe(resultBlob);
    expect(within(outputPanel).getAllByRole("listitem")).toHaveLength(3);
  });

  it("cancels Generate Dataset through native Escape while pending and ignores the obsolete analysis", async () => {
    const user = userEvent.setup();
    let resolve!: (response: Response) => void;
    let signal: AbortSignal | undefined;
    const { apiMock } = mockPlaygroundFetch((_input, init) => {
      signal = init?.signal as AbortSignal;
      return new Promise<Response>(done => { resolve = done; });
    });
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await generateJson(user);
    expect(apiMock).toHaveBeenCalledOnce();
    await user.keyboard("{Escape}");
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(signal?.aborted).toBe(true);
    expect(screen.getByRole("status")).toHaveTextContent("Generate Dataset · Cancelled");
    resolve(analysis([{ name: "Obsolete", type: "value" }]));
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Generate Dataset · Cancelled"));
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    apiMock.mockResolvedValueOnce(analysis());
    await generateJson(user);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).toHaveBeenCalledTimes(2);
    expect(JSON.parse((screen.getByLabelText("Dataset") as HTMLTextAreaElement).value)).toEqual({ Company: { Name: "" } });
  });

  it.each([
    ["MacIntel", true, "⇧ ⌥ F", "⌘ S"],
    ["Win32", false, "Shift+Alt+F", undefined],
    ["Linux x86_64", true, "Ctrl+Shift+I", "Ctrl+S"]
  ] as const)("groups implemented shortcuts by scope with separate platform keycaps and capability limits (%s)", async (platform, canSave, format, save) => {
    const user = userEvent.setup();
    vi.spyOn(navigator, "platform", "get").mockReturnValue(platform);
    if (canSave) {
      nativeFiles();
      vi.stubGlobal("showSaveFilePicker", vi.fn());
    }
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Application menu" }));
    await user.click(await screen.findByRole("menuitem", { name: /^Keyboard shortcuts/ }));
    const dialog = await screen.findByRole("dialog", { name: "Keyboard shortcuts" });
    const row = (action: string) => within(dialog).getByRole("group", { name: action });
    expect(dialog).toHaveStyle({ width: "440px", maxWidth: "calc(100vw - 32px)" });
    expect(within(dialog).queryByRole("table")).not.toBeInTheDocument();
    expect(within(dialog).queryByRole("columnheader")).not.toBeInTheDocument();
    expect(within(dialog).queryByText(/Action|Shortcut/)).not.toBeInTheDocument();
    const workspace = within(dialog).getByRole("region", { name: "Workspace" });
    const dataEditor = within(dialog).getByRole("region", { name: "Dataset editor" });
    expect(within(workspace).getByRole("group", { name: "Collapse diagnostics" })).toBe(row("Collapse diagnostics"));
    expect(within(dataEditor).getByRole("group", { name: "Format dataset" })).toBe(row("Format dataset"));
    expect(dataEditor.querySelectorAll(".ant-flex-vertical > [role='group']")).toHaveLength(4);
    const keycaps = (action: string) => Array.from(row(action).querySelectorAll("kbd"), key => key.textContent);
    expect(keycaps("Format dataset")).toEqual(format.split(/[ +]+/));
    expect(keycaps("Collapse diagnostics")).toEqual(platform === "MacIntel" ? ["⇧", "⎋"] : ["Shift", "Escape"]);
    expect(keycaps("Keyboard shortcuts")).toEqual(platform === "MacIntel" ? ["⇧", "/"] : ["Shift", "/"]);
    expect(within(row("Collapse diagnostics")).getByRole("group", { name: "Shift + Escape" })).toBeInTheDocument();
    expect(within(row("Keyboard shortcuts")).getByRole("group", { name: "Shift + /" })).toBeInTheDocument();
    expect(dialog).not.toHaveTextContent("(?)");
    expect(dialog).not.toHaveTextContent("Outside editors");
    expect(dialog).not.toHaveTextContent("When not handled");
    expect(row("Format dataset").firstElementChild).toHaveTextContent("Format dataset");
    expect(row("Format dataset").lastElementChild).toHaveClass("shortcut");
    expect(keycaps("Save dataset as")).toEqual(platform === "MacIntel" ? ["⇧", "⌘", "S"] : ["Ctrl", "Shift", "S"]);
    expect(within(row("Save dataset as")).getByRole("group", { name: platform === "MacIntel" ? "Shift + Command + S" : "Ctrl + Shift + S" })).toBeInTheDocument();
    expect(within(dialog).queryByText("Word wrap", { exact: true })).not.toBeInTheDocument();
    await user.keyboard("{Escape}");
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
  });

  it("keeps Dataset shortcuts available for an empty editor", async () => {
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "" } });
    await user.click(screen.getByRole("button", { name: "Application menu" }));
    await user.click(await screen.findByRole("menuitem", { name: /^Keyboard shortcuts/ }));
    const dialog = await screen.findByRole("dialog", { name: "Keyboard shortcuts" });
    expect(within(dialog).getByRole("region", { name: "Dataset editor" })).toBeInTheDocument();
    expect(within(dialog).getByRole("group", { name: "Upload dataset" })).toBeInTheDocument();
    await user.keyboard("{Escape}");
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(screen.getByLabelText("Dataset")).toBeVisible();
  });

  it("opens shortcut help with ? outside editing and keeps typing available in text and Monaco controls", async () => {
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const editor = screen.getByLabelText("Dataset");
    const typed = new KeyboardEvent("keydown", { key: "?", code: "Slash", shiftKey: true, bubbles: true, cancelable: true });
    editor.dispatchEvent(typed);
    expect(typed.defaultPrevented).toBe(false);
    await user.type(editor, "?");
    expect(editor).toHaveValue(sampleData + "?");
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    const controls = document.createElement("div");
    controls.innerHTML = '<input /><div contenteditable="true"><span>Editing</span></div><div class="monaco-editor"><button>Editor widget</button></div>';
    document.body.append(controls);
    for (const target of [controls.querySelector("input")!, controls.querySelector("span")!, controls.querySelector("button")!]) {
      const event = new KeyboardEvent("keydown", { key: "?", code: "Slash", shiftKey: true, bubbles: true, cancelable: true });
      target.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(false);
      expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    }
    controls.remove();
    const prevented = (event: KeyboardEvent) => event.preventDefault();
    const target = screen.getByRole("button", { name: "Generate document" });
    target.addEventListener("keydown", prevented);
    fireEvent.keyDown(target, { key: "?", code: "Slash", shiftKey: true });
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    target.removeEventListener("keydown", prevented);
    for (const extra of [{ ctrlKey: true }, { altKey: true }, { metaKey: true }, { repeat: true }, { isComposing: true }]) {
      fireEvent.keyDown(target, { key: "?", code: "Slash", shiftKey: true, ...extra });
      expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    }
    fireEvent.keyDown(document.body, { key: "?", code: "Slash", shiftKey: true });
    expect(await screen.findByRole("dialog", { name: "Keyboard shortcuts" })).toBeInTheDocument();
    await user.keyboard("{Escape}");
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(editor).toHaveValue(sampleData + "?");
  });

  it("collapses diagnostics with unhandled Shift+Escape while preserving native key handling and removing its listener on unmount", async () => {
    const user = userEvent.setup();
    const view = render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await diagnosticsTab(user, "Problems");
    const footer = screen.getByRole("contentinfo");
    const editor = screen.getByLabelText("Dataset");
    const prevent = (event: KeyboardEvent) => event.preventDefault();
    editor.addEventListener("keydown", prevent);
    fireEvent.keyDown(editor, { key: "Escape", shiftKey: true });
    expect(within(footer).queryByRole("button", { name: "Problems" })).not.toBeInTheDocument();
    editor.removeEventListener("keydown", prevent);
    fireEvent.keyDown(editor, { key: "Escape" });
    expect(within(footer).queryByRole("button", { name: "Problems" })).not.toBeInTheDocument();
    const event = new KeyboardEvent("keydown", { key: "Escape", shiftKey: true, bubbles: true, cancelable: true });
    fireEvent(editor, event);
    expect(event.defaultPrevented).toBe(true);
    expect(within(footer).getByRole("button", { name: "Problems" })).toBeInTheDocument();
    expect(within(footer).getByRole("status")).toBeInTheDocument();
    view.unmount();
    const after = new KeyboardEvent("keydown", { key: "?", code: "Slash", shiftKey: true, bubbles: true, cancelable: true });
    document.body.dispatchEvent(after);
    expect(after.defaultPrevented).toBe(false);
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
  });

  it("restores active native files from fresh disk bytes and keeps a closed template closed across reload", async () => {
    const user = userEvent.setup();
    const one = localHandle("one.json", '{"old":1}');
    const two = localHandle("two.json", '{"second":2}');
    const book = localHandle("remembered.xlsx", "original workbook");
    const fresh = new File(['{"fresh":3}'], "one.json", { lastModified: 9876 });
    vi.mocked(one.handle.getFile).mockResolvedValue(fresh);
    for (const item of [one, two, book]) Object.assign(item.handle, { queryPermission: vi.fn(async () => "granted"), requestPermission: vi.fn() });
    await rememberWorkspace({ datasets: [
      { id: "data-10", name: "one.json", kind: "file", handle: one.handle },
      { id: "data-20", name: "two.json", kind: "file", handle: two.handle }
    ], templates: [{ id: "template-30", name: "remembered.xlsx", handle: book.handle }], datasetId: "data-10", templateId: "template-30", selectedExample: "custom" });
    nativeFiles();
    vi.stubGlobal("showSaveFilePicker", vi.fn());
    let view = render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"fresh":3}'));
    expect(await screen.findByText("Template XLSX preview")).toBeInTheDocument();
    expect(xlsxSources.at(-1)).toBe(book.file);
    expect(one.handle.requestPermission).not.toHaveBeenCalled();
    await dataCommand(user, /^Save(?! As)/);
    expect(one.writable.write).toHaveBeenCalledWith('{"fresh":3}');
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    expect(two.handle.getFile).not.toHaveBeenCalled();
    await documentCommand(user, "Close");
    await waitFor(() => {
      const stored = JSON.parse(localStorage.getItem(workspaceStorage.context)!);
      expect(stored.datasets.map((item: { id: string }) => item.id)).toEqual(["data-10"]);
      expect(stored.templates).toEqual([]);
    });
    view.unmount();
    view = render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"fresh":3}'));
    expect(screen.getByText("Select a DOCX or XLSX template")).toBeInTheDocument();
    expect(screen.queryByRole("combobox")).not.toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent("No operations yet");
  });

  it("preserves partial context and asks for permission only after Grant access, retaining files after denial", async () => {
    const user = userEvent.setup();
    const data = localHandle("permission.json", '{"granted":true}');
    Object.assign(data.handle, { queryPermission: vi.fn(async () => "prompt"), requestPermission: vi.fn().mockResolvedValueOnce("denied").mockResolvedValue("granted") });
    await rememberWorkspace({ datasets: [{ id: "data-1", name: data.file.name, kind: "file", handle: data.handle }], templates: [{ id: "template-2", name: "sample.template.docx", exampleId: manifest.defaultExample }], datasetId: "data-1", templateId: "template-2" });
    nativeFiles();
    const { fetchMock } = mockPlaygroundFetch();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    expect(await screen.findByRole("button", { name: "Grant access" })).toBeInTheDocument();
    expect(screen.getByText("Template preview")).toBeInTheDocument();
    expect(data.handle.getFile).not.toHaveBeenCalled();
    expect(data.handle.requestPermission).not.toHaveBeenCalled();
    expect(fetchMock.mock.calls.some(([url]) => String(url) === "/examples/sample.json")).toBe(false);
    await user.click(screen.getByRole("button", { name: "Grant access" }));
    expect(await screen.findByText("Dataset could not be reopened")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Grant access" })).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Grant access" }));
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"granted":true}'));
    expect(data.handle.requestPermission).toHaveBeenCalledTimes(2);
    expect(data.handle.getFile).toHaveBeenCalledOnce();
    expect(data.handle.createWritable).not.toHaveBeenCalled();
    expect(screen.queryByRole("button", { name: "Grant access" })).not.toBeInTheDocument();
  });

  it("keeps a missing local file named and permits picker replacement without falling back to the default example", async () => {
    const user = userEvent.setup();
    const missing = localHandle("missing.json", "{}");
    Object.assign(missing.handle, { queryPermission: vi.fn(async () => "granted"), requestPermission: vi.fn() });
    vi.mocked(missing.handle.getFile).mockRejectedValue(new DOMException("The file was deleted.", "NotFoundError"));
    await rememberWorkspace({ datasets: [{ id: "data-1", name: "missing.json", kind: "file", handle: missing.handle }], templates: [], datasetId: "data-1" });
    const replacement = localHandle("replacement.json", '{"replacement":true}');
    const picker = nativeFiles([replacement.handle]);
    const { fetchMock } = mockPlaygroundFetch();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    expect(await screen.findByText("missing.json: The file was deleted.")).toBeInTheDocument();
    expect(fetchMock.mock.calls.map(([url]) => String(url))).toEqual(["/examples.json"]);
    await user.click(screen.getByRole("button", { name: "Reopen…" }));
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"replacement":true}'));
    expect(picker).toHaveBeenCalledOnce();
    expect(missing.handle.requestPermission).not.toHaveBeenCalled();
    expect(screen.queryByText(/missing.json: The file/)).not.toBeInTheDocument();
    await waitFor(() => expect(JSON.parse(localStorage.getItem(workspaceStorage.context)!).datasets).toHaveLength(1));
  });

  it("keeps Monaco visible for an upload-only restore reference and recovers it with fallback upload", async () => {
    const user = userEvent.setup();
    await new WorkspacePersistence(localStorage, undefined, new LocalFiles({})).save({ datasets: [{ id: "data-1", name: "session.json", kind: "upload" }], templates: [], datasetId: "data-1" });
    const mount = () => render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    let view = mount();
    const reopen = await screen.findByRole("button", { name: "Reopen…" });
    expect(screen.getByLabelText("Dataset")).toBeVisible();
    expect(screen.getByLabelText("Dataset")).toHaveValue("");
    const input = screen.getByLabelText("Upload dataset") as HTMLInputElement;
    const click = vi.spyOn(input, "click");
    await user.click(reopen);
    expect(click).toHaveBeenCalledOnce();
    await user.upload(input, new File(['{"uploaded":true}'], "reopened.json"));
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"uploaded":true}'));
    await waitFor(() => expect(JSON.parse(localStorage.getItem(workspaceStorage.context)!).datasets[0].name).toBe("reopened.json"));
    view.unmount();
    view = mount();
    await screen.findByRole("button", { name: "Reopen…" });
    expect(screen.getByLabelText("Dataset")).toBeVisible();
    expect(selectedExample()).toHaveTextContent("reopened.json");
  });

  it("restores the last example instead of the default and never persists edited text or generated documents", async () => {
    const second = { ...manifest.examples[0], id: "second", title: "Second", templatePath: "/second.docx", dataPath: "/second.json" };
    const fetchMock = vi.fn(async (input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/examples.json") return new Response(JSON.stringify({ ...manifest, examples: [...manifest.examples, second] }));
      if (url === second.templatePath) return new Response(new Uint8Array([80, 75]));
      if (url === second.dataPath) return new Response('{"second":true}');
      throw new Error(`Unexpected request: ${url}`);
    });
    vi.stubGlobal("fetch", fetchMock);
    await rememberWorkspace({ datasets: [{ id: "data-1", name: "second.json", kind: "example", exampleId: "second" }], templates: [{ id: "template-2", name: "old-template.xlsx", exampleId: "second" }], datasetId: "data-1", templateId: "template-2", selectedExample: "second" });
    const mount = () => render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    let view = mount();
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"second":true}'));
    expect(screen.getByText("Template preview")).toBeInTheDocument();
    expect((previewSources.at(-1) as File).name).toBe("second.docx");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"PRIVATE EDIT":true}' } });
    expect(localStorage.getItem(workspaceStorage.context)).not.toContain("PRIVATE EDIT");
    view.unmount();
    view = mount();
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"second":true}'));
    expect(fetchMock.mock.calls.some(([url]) => String(url).includes("/examples/sample"))).toBe(false);
    expect(screen.getByRole("status")).toHaveTextContent("No operations yet");
  });

  it("ignores a late restored file after a newer user upload and remembers the newer context", async () => {
    const user = userEvent.setup();
    const remembered = localHandle("slow.json", "{}");
    let resolve!: (file: File) => void;
    Object.assign(remembered.handle, { queryPermission: vi.fn(async () => "granted") });
    vi.mocked(remembered.handle.getFile).mockImplementation(() => new Promise<File>(done => { resolve = done; }));
    await rememberWorkspace({ datasets: [{ id: "data-1", name: "slow.json", kind: "file", handle: remembered.handle }], templates: [], datasetId: "data-1" });
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await waitFor(() => expect(remembered.handle.getFile).toHaveBeenCalled());
    await user.upload(screen.getByLabelText("Upload dataset"), new File(['{"newer":true}'], "newer.json"));
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"newer":true}'));
    resolve(new File(['{"obsolete":true}'], "slow.json"));
    await waitFor(() => expect(localStorage.getItem(workspaceStorage.context)).toContain("newer.json"));
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"newer":true}');
    expect(localStorage.getItem(workspaceStorage.context)).not.toContain("slow.json");
  });

  it("cancels Clear Local Data without changes and confirms precisely scoped session/storage reset without disk writes", async () => {
    const user = userEvent.setup();
    const db = testIndexedDb();
    vi.stubGlobal("indexedDB", db.factory);
    db.databases.set("other-app", new Map([["items", new Map([["key", "keep"]])]]));
    const data = localHandle("disk.json", '{"disk":true}');
    nativeFiles([data.handle]);
    const { apiMock } = mockPlaygroundFetch(async () => success());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, /^Open/);
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"disk":true}'));
    await dataCommand(user, "Word wrap");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"unsaved":true}' } });
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    for (const key of [workspaceStorage.desktop, workspaceStorage.mobile, workspaceStorage.desktopCollapsed, workspaceStorage.mobileCollapsed, workspaceStorage.diagnosticsSize, workspaceStorage.diagnosticsCollapsed]) localStorage.setItem(key, "63");
    localStorage.setItem("unrelated", "keep");
    await waitFor(() => expect(localStorage.getItem(workspaceStorage.context)).toContain("disk.json"));
    const before = localStorage.getItem(workspaceStorage.context);
    const status = screen.getByRole("status").textContent;
    let dialog = await clearWorkspaceDialog(user);
    expect(dialog).toHaveTextContent("Unsaved workspace changes will be discarded");
    expect(dialog).toHaveTextContent("Files on disk will not be deleted");
    await user.click(within(dialog).getByRole("button", { name: "Cancel" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(localStorage.getItem(workspaceStorage.context)).toBe(before);
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"unsaved":true}');
    expect(screen.getByText("Result preview")).toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent(status!);
    dialog = await clearWorkspaceDialog(user);
    await user.click(within(dialog).getByRole("button", { name: "Clear" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(await screen.findByText("Local data cleared", { exact: true })).toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.getByLabelText("Dataset")).toHaveAttribute("data-word-wrap", "false");
    expect(screen.getByRole("status")).toHaveTextContent("No operations yet");
    expect(apiMock).toHaveBeenCalledOnce();
    for (const key of [workspaceStorage.desktop, workspaceStorage.mobile, workspaceStorage.desktopCollapsed, workspaceStorage.mobileCollapsed, workspaceStorage.diagnosticsSize, workspaceStorage.diagnosticsCollapsed, workspaceStorage.wordWrap]) expect(localStorage.getItem(key)).toBeNull();
    expect(localStorage.getItem("unrelated")).toBe("keep");
    expect(db.databases.get("other-app")!.get("items")!.get("key")).toBe("keep");
    expect(data.handle.createWritable).not.toHaveBeenCalled();
    await waitFor(() => expect(localStorage.getItem(workspaceStorage.context)).not.toContain("disk.json"));
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(screen.getByText("Not generated")).toBeInTheDocument();
    await diagnosticsTab(user, "Output");
    expect(screen.getByText("No output")).toBeInTheDocument();
  });

  it("Clear Local Data aborts pending generation and late replies cannot restore results or Output", async () => {
    const user = userEvent.setup();
    vi.stubGlobal("indexedDB", testIndexedDb().factory);
    let resolve!: (response: Response) => void;
    let signal: AbortSignal | undefined;
    const { apiMock } = mockPlaygroundFetch(async (_input, init) => { signal = init?.signal as AbortSignal; return new Promise<Response>(done => { resolve = done; }); });
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    const dialog = await clearWorkspaceDialog(user);
    await user.click(within(dialog).getByRole("button", { name: "Clear" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(signal?.aborted).toBe(true);
    resolve(success());
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("No operations yet"));
    expect(screen.queryByText("Result preview")).not.toBeInTheDocument();
    await diagnosticsTab(user, "Output");
    expect(screen.getByText("No output")).toBeInTheDocument();
    expect(apiMock).toHaveBeenCalledOnce();
  });

  it("keeps Clear Local Data confirmation open if unsaved work changes and reports partial storage failure truthfully", async () => {
    const user = userEvent.setup();
    const db = testIndexedDb();
    vi.stubGlobal("indexedDB", db.factory);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    localStorage.setItem(workspaceStorage.desktop, "63");
    let dialog = await clearWorkspaceDialog(user);
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"later edit":true}' } });
    await user.click(within(dialog).getByRole("button", { name: "Clear" }));
    expect(await screen.findByText("The workspace changed. Cancel and review Clear Local Data again.")).toBeInTheDocument();
    expect(localStorage.getItem(workspaceStorage.desktop)).toBe("63");
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"later edit":true}');
    await user.click(within(dialog).getByRole("button", { name: "Cancel" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    db.failure.clear = true;
    db.failure.write = true;
    dialog = await clearWorkspaceDialog(user);
    await user.click(within(dialog).getByRole("button", { name: "Clear" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(await screen.findByText(/Stored file handles could not be checked or cleared/)).toBeInTheDocument();
    expect(screen.queryByText("Local data cleared", { exact: true })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
  });

  it("restores and grants access to only the active template without replacing Dataset content", async () => {
    const user = userEvent.setup();
    const accessible = localHandle("accessible.docx", "original document");
    const blocked = localHandle("blocked.xlsx", "original workbook");
    Object.assign(accessible.handle, { queryPermission: vi.fn(async () => "granted"), requestPermission: vi.fn() });
    Object.assign(blocked.handle, { queryPermission: vi.fn(async () => "prompt"), requestPermission: vi.fn(async () => "granted") });
    await rememberWorkspace({ datasets: [{ id: "data-1", name: "sample.json", kind: "example", exampleId: manifest.defaultExample }], templates: [
      { id: "template-2", name: accessible.file.name, handle: accessible.handle },
      { id: "template-3", name: blocked.file.name, handle: blocked.handle }
    ], datasetId: "data-1", templateId: "template-3" });
    nativeFiles();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    const grant = await screen.findByRole("button", { name: "Grant access" });
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(blocked.handle.getFile).not.toHaveBeenCalled();
    expect(accessible.handle.getFile).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Grant access" }));
    expect(await screen.findByText("Template XLSX preview")).toBeInTheDocument();
    expect(xlsxSources.at(-1)).toBe(blocked.file);
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(blocked.handle.requestPermission).toHaveBeenCalledWith({ mode: "read" });
    expect(accessible.handle.requestPermission).not.toHaveBeenCalled();
    expect(blocked.handle.createWritable).not.toHaveBeenCalled();
    expect(grant).not.toBeInTheDocument();
  });

  it("does not let a file read started before Clear Local Data repopulate the cleared workspace", async () => {
    const user = userEvent.setup();
    vi.stubGlobal("indexedDB", testIndexedDb().factory);
    const slow = localHandle("slow.json", "{}");
    let resolve!: (text: string) => void;
    vi.spyOn(slow.file, "text").mockImplementation(() => new Promise<string>(done => { resolve = done; }));
    nativeFiles([slow.handle]);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, /^Open/);
    await waitFor(() => expect(slow.file.text).toHaveBeenCalled());
    const dialog = await clearWorkspaceDialog(user);
    await user.click(within(dialog).getByRole("button", { name: "Clear" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    resolve('{"late":true}');
    await waitFor(() => expect(localStorage.getItem(workspaceStorage.context)).toContain(manifest.defaultExample));
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(localStorage.getItem(workspaceStorage.context)).not.toContain("slow.json");
    expect(screen.getByRole("status")).toHaveTextContent("No operations yet");
  });

});
