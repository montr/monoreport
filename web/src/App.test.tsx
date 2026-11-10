import { act, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
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
  return { MonacoJsonEditor: ({ value, onChange, ref, wordWrap, onWordWrapChange, readOnly, ariaLabel = "Dataset" }: { value: string; onChange(value: string): void; ref: import("react").Ref<import("./components/MonacoJsonEditor").JsonEditorCommands>; wordWrap: boolean; onWordWrapChange(enabled: boolean): void; readOnly?: boolean; ariaLabel?: string }) => {
    useImperativeHandle(ref, () => ({
      async format() { onChange(JSON.stringify(JSON.parse(value), null, 2)); },
      async copy() { await navigator.clipboard.writeText(value); },
      toggleWrap() { onWordWrapChange(!wordWrap); }
    }));
    return <div className="monaco-editor"><textarea aria-label={ariaLabel} data-word-wrap={wordWrap} readOnly={readOnly} value={value} onChange={event => onChange(event.target.value)} /></div>;
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
const success = (bytes = new Uint8Array([80, 75, 7, 8]), mediaType = mime, filename = mediaType === xlsxMime ? "document.xlsx" : "document.docx", warnings: unknown[] = []) => {
  const boundary = "playground-test-boundary";
  const encoder = new TextEncoder();
  const format = mediaType === xlsxMime ? "xlsx" : "docx";
  const warningBytes = encoder.encode(JSON.stringify({ format, fileName: filename, warnings }));
  const parts = [
    encoder.encode(`--${boundary}\r\nContent-Type: ${mediaType}\r\nContent-Disposition: attachment; filename="${filename}"\r\n\r\n`),
    bytes,
    encoder.encode(`\r\n--${boundary}\r\nContent-Type: application/json\r\nContent-Disposition: attachment; filename="diagnostics.json"\r\n\r\n`),
    warningBytes,
    encoder.encode(`\r\n--${boundary}--\r\n`)
  ];
  const body = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) { body.set(part, offset); offset += part.length; }
  return new Response(body.buffer, { status: 200, headers: { "Content-Type": `multipart/mixed; boundary="${boundary}"` } });
};
const analysis = (requirements: unknown[] = [{ name: "Company", type: "object", fields: [{ name: "Name", type: "value" }] }], valid = true) =>
  new Response(JSON.stringify({ valid, errors: valid ? [] : [{ code: "template.invalid_structure", message: "Invalid block" }], warnings: [], requirements, occurrences: [] }), { status: 200, headers: { "Content-Type": "application/json" } });
const sampleJson = (data: unknown = { Company: { Name: "Sample Name 1" } }, warnings: unknown[] = []) =>
  new Response(JSON.stringify({ data, warnings }), { status: 200, headers: { "Content-Type": "application/json" } });

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
  await user.upload(document.querySelector('input[type="file"][accept*=".docx"]') as HTMLInputElement, file);
}

async function expectTemplateBytes(value: unknown, expected: File) {
  expect(value).toBeInstanceOf(File);
  expect((value as File).name).toBe(expected.name);
  expect(new Uint8Array(await (value as File).arrayBuffer())).toEqual(new Uint8Array(await expected.arrayBuffer()));
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

async function analyzeTemplate(user: ReturnType<typeof userEvent.setup>) {
  await user.click(screen.getByRole("button", { name: "Generate actions" }));
  await user.click(await screen.findByRole("menuitem", { name: "Analyze Template" }));
}

async function dataCommand(user: ReturnType<typeof userEvent.setup>, name: string | RegExp) {
  await openMenu(user, "Dataset");
  await user.click(await screen.findByRole(name === "Word wrap" ? "menuitemcheckbox" : "menuitem", { name: name === "Format" ? /^Format/ : name }));
}

async function generateJson(user: ReturnType<typeof userEvent.setup>, count = 10) {
  const modified = Boolean(screen.queryByLabelText("Dataset modified"));
  await user.click(screen.getByRole("button", { name: "Generate actions" }));
  await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
  const dialog = await screen.findByRole("dialog");
  if (count !== 10) fireEvent.change(within(dialog).getByRole("spinbutton", { name: "Items per collection" }), { target: { value: String(count) } });
  await user.click(within(dialog).getByRole("button", { name: "Generate sample" }));
  await user.click(await within(dialog).findByRole("button", { name: "Replace Dataset" }));
  if (modified) await confirmSampleReplacement(user);
}

async function confirmSampleReplacement(user: ReturnType<typeof userEvent.setup>, accept = true) {
  const confirmation = await screen.findByRole("dialog", { name: "Replace Dataset?" });
  await user.click(within(confirmation).getByRole("button", { name: accept ? "Replace" : "Cancel" }));
  await waitFor(() => expect(screen.queryByRole("dialog", { name: "Replace Dataset?" })).not.toBeInTheDocument());
}

function expectResultOutdated(outdated = true) {
  const indicator = screen.queryByText("Outdated", { exact: true });
  const resultActive = (screen.getByRole("radio", { name: "Result" }) as HTMLInputElement).checked;
  if (outdated && resultActive) expect(indicator).toBeInTheDocument();
  else expect(indicator).not.toBeInTheDocument();
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
    vi.stubGlobal("showDirectoryPicker", undefined);
    vi.stubGlobal("FileSystemObserver", undefined);
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

  it("uses stable Preview and fresh Generate bytes after a native Template is replaced with Live Reload off", async () => {
    const source = localHandle("current.docx", "before");
    nativeFiles([source.handle]);
    const { apiMock } = mockPlaygroundFetch(async () => success());
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await documentCommand(user, "Open…");
    await screen.findByText("current.docx");
    const accepted = previewSources.at(-1)!;
    vi.spyOn(source.file, "arrayBuffer").mockRejectedValue(new DOMException("Changed", "NotReadableError"));
    expect(await accepted.text()).toBe("before");
    vi.mocked(source.handle.getFile).mockResolvedValue(new File(["after"], source.file.name));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await waitFor(() => expect(apiMock).toHaveBeenCalledOnce());
    const body = apiMock.mock.calls[0][1]!.body as FormData;
    expect(await (body.get("template") as File).text()).toBe("after");
    await screen.findByText("Result preview");
    vi.mocked(source.handle.getFile).mockRejectedValue(new DOMException("Gone", "NotReadableError"));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Generation failed", { exact: true });
    expect(apiMock).toHaveBeenCalledOnce();
    expect(screen.getByText("Result preview")).toBeInTheDocument();
  });

  it("debounces editor Live Reload, pauses invalid JSON and avoids retrying an unchanged server failure", async () => {
    const { apiMock } = mockPlaygroundFetch(async () => success());
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
    await waitFor(() => expect(apiMock).toHaveBeenCalledOnce());
    fireEvent.change(screen.getByRole("textbox", { name: "Dataset" }), { target: { value: "{" } });
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 600)); });
    expect(apiMock).toHaveBeenCalledOnce();
    apiMock.mockRejectedValue(new Error("Offline"));
    fireEvent.change(screen.getByRole("textbox", { name: "Dataset" }), { target: { value: '{"new":true}' } });
    await waitFor(() => expect(apiMock).toHaveBeenCalledTimes(2));
    await screen.findByText("Server unavailable", { exact: true });
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 650)); });
    expect(apiMock).toHaveBeenCalledTimes(2);
  });

  it("obsoletes Generate preparation when editor inputs change during a fresh native read", async () => {
    const source = localHandle("preparation.docx", "before");
    nativeFiles([source.handle]);
    const { apiMock } = mockPlaygroundFetch(async () => success());
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await documentCommand(user, "Open…");
    await screen.findByText("preparation.docx");
    let finish!: (file: File) => void;
    vi.mocked(source.handle.getFile).mockImplementationOnce(() => new Promise(resolve => { finish = resolve; }));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    fireEvent.change(screen.getByRole("textbox", { name: "Dataset" }), { target: { value: '{"edit":true}' } });
    await act(async () => { finish(new File(["after"], source.file.name)); });
    expect(apiMock).not.toHaveBeenCalled();
    expect(screen.getByRole("status")).toHaveTextContent("Cancelled");
  });

  it.each(["docx", "xlsx"])("polls production %s handles from Open and Drop through current multipart inputs without no-op generation", async format => {
    const intervals = vi.spyOn(window, "setInterval");
    const user = userEvent.setup();
    const initial = localHandle(`production.${format}`, "one");
    nativeFiles([initial.handle]);
    const { apiMock } = mockPlaygroundFetch(async () => success(undefined, format === "xlsx" ? xlsxMime : mime));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    let expectedRequests = 0;
    for (const origin of ["Open", "Drop"]) {
      await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
      const native = origin === "Open" ? initial.handle : { ...localHandle(`dropped.${format}`, "one").handle, kind: "file" as const };
      if (origin === "Open") await documentCommand(user, "Open…");
      else await act(async () => { fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, {
        dataTransfer: { files: [new File(["one"], native.name)], types: ["Files"], items: [{ kind: "file", getAsFileSystemHandle: () => Promise.resolve(native) }] }
      }); });
      await screen.findByText(native.name);
      await user.click(screen.getByRole("button", { name: "Generate actions" }));
      await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
      expectedRequests++;
      await waitFor(() => expect(apiMock).toHaveBeenCalledTimes(expectedRequests));
      const poll = intervals.mock.calls.filter(call => call[1] === 1000).at(-1)![0] as () => void;
      vi.mocked(native.getFile).mockResolvedValue(new File(["two"], native.name, { lastModified: 42 }));
      act(() => poll());
      expectedRequests++;
      await waitFor(() => expect(apiMock).toHaveBeenCalledTimes(expectedRequests));
      expect(await ((apiMock.mock.calls.at(-1)![1]!.body as FormData).get("template") as File).text()).toBe("two");
      await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
      expect(await (format === "xlsx" ? xlsxSources : previewSources).at(-1)!.text()).toBe("two");
      act(() => poll());
      await act(async () => { await new Promise(resolve => setTimeout(resolve, 600)); });
      expect(apiMock).toHaveBeenCalledTimes(expectedRequests);
      await user.click(screen.getByRole("button", { name: "Generate actions" }));
      await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
    }
  });

  it("keeps long Generate ownership across unchanged polls and rejects a response after a confirmed Template change", async () => {
    const intervals = vi.spyOn(window, "setInterval");
    const source = localHandle("long.docx", "one");
    nativeFiles([source.handle]);
    let finish!: (response: Response) => void;
    const { apiMock } = mockPlaygroundFetch(() => new Promise(resolve => { finish = resolve; }));
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await documentCommand(user, "Open…");
    await screen.findByText("long.docx");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await waitFor(() => expect(apiMock).toHaveBeenCalledOnce());
    const signal = apiMock.mock.calls[0][1]!.signal!;
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
    const poll = intervals.mock.calls.find(call => call[1] === 1000)![0] as () => void;
    await act(async () => { poll(); });
    expect(signal.aborted).toBe(false);
    vi.mocked(source.handle.getFile).mockResolvedValue(new File(["two"], source.file.name));
    await act(async () => { poll(); });
    expect(signal.aborted).toBe(true);
    await act(async () => finish(success()));
    expect(screen.queryByText("Result preview")).not.toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent("Cancelled");
  });

  it("polls external JSON while dirty and confirms equal-metadata divergence before Save", async () => {
    const intervals = vi.spyOn(window, "setInterval");
    const source = localHandle("external.json", '{"v":1}');
    nativeFiles([source.handle]);
    vi.stubGlobal("showSaveFilePicker", vi.fn());
    mockPlaygroundFetch(async () => success());
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await dataCommand(user, /^Open…/);
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"v":1}'));
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
    const poll = intervals.mock.calls.find(call => call[1] === 1000)![0] as () => void;
    vi.mocked(source.handle.getFile).mockResolvedValue(new File(['{"v":2}'], source.file.name, { lastModified: source.file.lastModified }));
    await act(async () => { poll(); });
    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue('{"v":2}'));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"v":3}' } });
    vi.mocked(source.handle.getFile).mockResolvedValue(new File(['{"v":4}'], source.file.name, { lastModified: source.file.lastModified }));
    await act(async () => { poll(); });
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"v":3}');
    expect(screen.getByLabelText("Dataset modified")).toHaveTextContent("Disk differs");
    await dataCommand(user, /^Save(?! As)/);
    const confirmation = await screen.findByRole("dialog", { name: "File changed on disk" });
    await user.click(within(confirmation).getByRole("button", { name: "Cancel" }));
    expect(source.writable.write).not.toHaveBeenCalled();
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"v":3}');
  });

  it("reports Template monitoring only after a successful sample and removes the claim on read failure", async () => {
    const intervals = vi.spyOn(window, "setInterval");
    const source = localHandle("status.docx", "one");
    nativeFiles([source.handle]);
    mockPlaygroundFetch(async () => success());
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await documentCommand(user, "Open…");
    await screen.findByText("status.docx");
    let finish!: (file: File) => void;
    vi.mocked(source.handle.getFile).mockImplementationOnce(() => new Promise(resolve => { finish = resolve; }));
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
    expect(screen.getByLabelText("Template Live Reload")).toHaveTextContent("Reading Template");
    await act(async () => finish(source.file));
    expect(screen.getByLabelText("Template Live Reload")).toHaveTextContent("Template monitored");
    const poll = intervals.mock.calls.find(call => call[1] === 1000)![0] as () => void;
    vi.mocked(source.handle.getFile).mockRejectedValueOnce(new DOMException("Busy", "NotReadableError"));
    await act(async () => { poll(); });
    expect(screen.getByLabelText("Template Live Reload")).toHaveTextContent("Template unavailable");
    expect(screen.getByRole("button", { name: "Reopen…" })).toBeInTheDocument();
    await act(async () => { poll(); });
    expect(screen.getByLabelText("Template Live Reload")).toHaveTextContent("Template monitored");
    expect(screen.queryByRole("button", { name: "Reopen…" })).not.toBeInTheDocument();
  });

  it("disables automatic work without cancelling a subsequent explicit Analyze", async () => {
    let finish!: (response: Response) => void;
    const { apiMock } = mockPlaygroundFetch(async input => String(input).endsWith("/analyze") ? new Promise(resolve => { finish = resolve; }) : success());
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
    await screen.findByText("Result preview");
    await analyzeTemplate(user);
    await waitFor(() => expect(apiMock).toHaveBeenCalledTimes(2));
    const signal = apiMock.mock.calls[1][1]!.signal!;
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
    expect(signal.aborted).toBe(false);
    await act(async () => finish(analysis()));
    expect(screen.getByRole("status")).toHaveTextContent("Success");
  });

  it.each(["success", "failure"])("reconciles pending file reads before publishing Generate %s or diagnostics", async outcome => {
    const intervals = vi.spyOn(window, "setInterval");
    const source = localHandle("publication.docx", "one");
    nativeFiles([source.handle]);
    let finish!: (response: Response) => void;
    let fail!: (error: Error) => void;
    const { apiMock } = mockPlaygroundFetch(() => new Promise((resolve, reject) => { finish = resolve; fail = reject; }));
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await documentCommand(user, "Open…");
    await screen.findByText("publication.docx");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await waitFor(() => expect(apiMock).toHaveBeenCalledOnce());
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
    let finishRead!: (file: File) => void;
    vi.mocked(source.handle.getFile).mockImplementationOnce(() => new Promise(resolve => { finishRead = resolve; }));
    const poll = intervals.mock.calls.find(call => call[1] === 1000)![0] as () => void;
    await act(async () => { poll(); });
    await act(async () => { if (outcome === "success") finish(success()); else fail(new Error("Obsolete failure")); });
    expect(screen.queryByText("Result preview")).not.toBeInTheDocument();
    expect(screen.queryByText("Server unavailable", { exact: true })).not.toBeInTheDocument();
    await act(async () => finishRead(new File(["two"], source.file.name)));
    expect(screen.queryByText("Result preview")).not.toBeInTheDocument();
    expect(screen.queryByText("Server unavailable", { exact: true })).not.toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent("Cancelled");
  });

  it("enables Live Reload without regenerating a matching successful Result", async () => {
    const { apiMock } = mockPlaygroundFetch(async () => success());
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 650)); });
    expect(apiMock).toHaveBeenCalledOnce();
  });

  it("disabling Live Reload prevents an automatic preparation from publishing late Template bytes", async () => {
    const source = localHandle("disabled.docx", "one");
    nativeFiles([source.handle]);
    const { apiMock } = mockPlaygroundFetch(async () => success());
    const user = userEvent.setup();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await documentCommand(user, "Open…");
    await screen.findByText("disabled.docx");
    const accepted = previewSources.at(-1);
    let finish!: (file: File) => void;
    vi.mocked(source.handle.getFile).mockImplementationOnce(() => new Promise(resolve => { finish = resolve; }));
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
    await waitFor(() => expect(screen.getByRole("button", { name: "Generate document" })).toBeDisabled());
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitemcheckbox", { name: "Live Reload" }));
    await act(async () => finish(new File(["late"], source.file.name)));
    expect(apiMock).not.toHaveBeenCalled();
    expect(previewSources.at(-1)).toBe(accepted);
    expect(screen.getByRole("status")).toHaveTextContent("Cancelled");
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
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
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
    expect(screen.getByRole("status")).toHaveTextContent("Success");
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
    await act(async () => {
      finish();
    });
    expect(screen.getByLabelText("Dataset modified")).toBeInTheDocument();
    expect(savePicker).not.toHaveBeenCalled();
    await dataCommand(user, /^Save As…/);
    await waitFor(() => expect(destination.writable.close).toHaveBeenCalledOnce());
    expect(destination.writable.write).toHaveBeenCalledWith('{"value":3}');
    expect(screen.queryByLabelText("Dataset modified")).not.toBeInTheDocument();
    expect(selectedExample()).toHaveTextContent("copy.json");
    vi.mocked(destination.handle.getFile).mockResolvedValue(new File(['{"value":3}'], "copy.json", { lastModified: destination.file.lastModified }));
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
    fireEvent(format, save);
    expect(save.defaultPrevented).toBe(true);
    await waitFor(() => expect(original.writable.write).toHaveBeenCalledWith('{"edited":true}'));
    expect(savePicker).not.toHaveBeenCalled();

    const editor = screen.getByLabelText("Dataset");
    const saveAs = new KeyboardEvent("keydown", { key: "s", ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true });
    fireEvent(editor, saveAs);
    expect(saveAs.defaultPrevented).toBe(true);
    await waitFor(() => expect(destination.writable.write).toHaveBeenCalledWith('{"edited":true}'));
    expect(savePicker).toHaveBeenCalledOnce();

    const input = document.createElement("input");
    document.body.append(input);
    const inputSave = new KeyboardEvent("keydown", { key: "s", ctrlKey: true, bubbles: true, cancelable: true });
    fireEvent(input, inputSave);
    expect(inputSave.defaultPrevented).toBe(false);
    expect(savePicker).toHaveBeenCalledOnce();
    input.remove();

    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const spinbutton = await screen.findByRole("spinbutton", { name: "Items per collection" });
    const dialogSave = new KeyboardEvent("keydown", { key: "s", ctrlKey: true, bubbles: true, cancelable: true });
    fireEvent(spinbutton, dialogSave);
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
    const expectedDocumentBytes = new Uint8Array([80, 75, 7, 8]);
    mockPlaygroundFetch(async () => response);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await documentCommand(user, "Open…");
    await waitFor(async () => expectTemplateBytes(previewSources.at(-1), one.file));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
    await documentCommand(user, "Open…");
    await waitFor(async () => expectTemplateBytes(xlsxSources.at(-1), two.file));
    expectResultOutdated();
    await documentCommand(user, "Open…");
    await waitFor(() => expect(picker).toHaveBeenCalledTimes(3));
    await waitFor(async () => expectTemplateBytes(previewSources.at(-1), one.file));
    expectResultOutdated(false);
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    await documentCommand(user, "Open…");
    await expectTemplateBytes(xlsxSources.at(-1), two.file);
    await documentCommand(user, "Close");
    expect(screen.getByText("Select a DOCX or XLSX template")).toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(new Uint8Array(await previewSources.at(-1)!.arrayBuffer())).toEqual(expectedDocumentBytes);
    let downloadedFilename = "";
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) { downloadedFilename = this.download; });
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(previewSources.at(-1));
    expect(downloadedFilename).toBe("one-result.docx");
  });

  it("activates Template when native Open finishes while Result is selected and when selecting an existing template", async () => {
    const user = userEvent.setup();
    const opened = localHandle("local.xlsx", "template bytes");
    let finish!: (file: File) => void;
    vi.mocked(opened.handle.getFile).mockImplementationOnce(() => new Promise<File>(resolve => { finish = resolve; }));
    nativeFiles([opened.handle]);
    const response = success();
    const { apiMock } = mockPlaygroundFetch(async () => response);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
    await documentCommand(user, "Open…");
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    await act(async () => {
      finish(opened.file);
    });
    await screen.findByText("Template XLSX preview");
    expect(screen.getByRole("radio", { name: "Template" })).toBeChecked();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(new Uint8Array(await previewSources.at(-1)!.arrayBuffer())).toEqual(new Uint8Array([80, 75, 7, 8]));

    await selectTemplate(user, new File(["fallback"], "uploaded.docx"));
    expect(screen.getByRole("radio", { name: "Template" })).toBeChecked();
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(new Uint8Array(await previewSources.at(-1)!.arrayBuffer())).toEqual(new Uint8Array([80, 75, 7, 8]));
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
    await act(async () => {
      finish();
    });
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
    expect(screen.getByRole("status")).toHaveTextContent("Ready");
    expect(screen.getByRole("link", { name: "MonoReport" })).toHaveAttribute("href", "https://github.com/montr/monoreport");
    expect(screen.getByText("dev")).toBeInTheDocument();
    expect(screen.queryByRole("link", { name: "API documentation" })).not.toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Application menu" }));
    const docs = await screen.findByRole("link", { name: "API documentation" });
    expect(docs).toHaveAttribute("href", "/docs/");
    expect(docs).toHaveAttribute("target", "_blank");
    expect(docs).toHaveAttribute("rel", "noopener noreferrer");
    expect(screen.getByRole("menuitem", { name: /^Keyboard shortcuts/ })).toBeInTheDocument();
    expect(screen.getByRole("menuitem", { name: "Clear Local Data…" })).toBeInTheDocument();
  });

  it("keeps template download disabled without a template and opens the upload picker", async () => {
    const user = userEvent.setup();
    vi.stubGlobal("fetch", vi.fn(async () => new Response("Unavailable", { status: 503 })));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    expect(await screen.findByText("Examples unavailable")).toBeInTheDocument();
    await openMenu(user, "Document");
    expect(screen.getByRole("menuitem", { name: "Download" })).toHaveAttribute("aria-disabled", "true");
    const input = document.querySelector('input[type="file"][accept*=".docx"]') as HTMLInputElement;
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
    await act(async () => { fireEvent.drop(screen.getByText("Select a DOCX or XLSX template"), { dataTransfer: { files: [book], types: ["Files"] } }); });
    expect(await screen.findByText("Template XLSX preview")).toBeInTheDocument();
    await expectTemplateBytes(xlsxSources.at(-1), book);
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
    await act(async () => {
      resolveXlsxTemplate(new Response(new Uint8Array([80, 75, 3, 4]), { status: 200, headers: { "Content-Type": xlsxMime } }));
      resolveXlsxData(new Response('{"stale":true}', { status: 200 }));
    });

    await waitFor(() => expect(screen.getByText("custom.xlsx")).toBeInTheDocument());
    expect(selectedExample()).toHaveTextContent(/sample.json|Untitled.json/);
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    await expectTemplateBytes(xlsxSources.at(-1), custom);
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
    await act(async () => {
      resolveTemplate(new Response(new Uint8Array([80, 75, 3, 4]), { status: 200, headers: { "Content-Type": mime } }));
      resolveData(new Response(sampleData, { status: 200 }));
    });

    await waitFor(() => expect(screen.getByText("custom.xlsx")).toBeInTheDocument());
    expect(screen.getByLabelText("Dataset")).toHaveValue("{}");
    expect(screen.getByText("Template XLSX preview")).toBeInTheDocument();
  });

  it("formats and copies valid JSON while Generate reports invalid JSON from the API", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => new Response(JSON.stringify({
      title: "Invalid request",
      detail: "The data field must contain one valid JSON value.",
      code: "data.invalid_json"
    }), { status: 400, headers: { "Content-Type": "application/problem+json" } }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const editor = screen.getByLabelText("Dataset");
    const clipboard = vi.spyOn(navigator.clipboard, "writeText");
    fireEvent.change(editor, { target: { value: '{"a":1}' } });
    await dataCommand(user, "Format");
    expect(editor).toHaveValue('{\n  "a": 1\n}');
    await dataCommand(user, "Copy");
    expect(clipboard).toHaveBeenCalledWith('{\n  "a": 1\n}');
    expect(await screen.findByText("Copied to clipboard.")).toBeInTheDocument();
    fireEvent.change(editor, { target: { value: "{" } });
    await dataCommand(user, "Format");
    expect(editor).toHaveValue("{");
    expect(screen.getByText("Invalid JSON")).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(await screen.findByText(/data.invalid_json/)).toBeInTheDocument();
    expect(apiMock).toHaveBeenCalledOnce();
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
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    expectResultOutdated(false);
    expect(screen.queryByText("Result preview")).not.toBeInTheDocument();
    expect(screen.queryByText("sample.docx")).not.toBeInTheDocument();
    expect(screen.getByRole("radio", { name: hasTemplate ? "Template" : "Result" })).toBeChecked();
    if (hasTemplate) {
      expect(previewSources.at(-1)).toBe(originalTemplate);
      await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    }
    expect(screen.getByText("Not generated", { exact: true })).toBeInTheDocument();
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
    expect(screen.getByRole("status")).toHaveTextContent("Success");
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
    await act(async () => { fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [replacement], types: ["Files"] } }); });
    expect(screen.getByText("second.docx")).toBeInTheDocument();
    expect(editor).toHaveValue('{"kept":true}');
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    expect(screen.getByRole("radio", { name: "Result" })).toBeChecked();
    await act(async () => { fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [new File([new Uint8Array([1])], "bad.pdf")], types: ["Files"] } }); });
    expect(screen.getByText("Unsupported template format")).toBeInTheDocument();
    expect(screen.getByText("second-result.docx")).toBeInTheDocument();
    await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
    expect(screen.getByText("second.docx")).toBeInTheDocument();
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(screen.getByRole("radio", { name: "Result" })).toBeChecked();
    const third = new File([new Uint8Array([2])], "third.docx", { type: mime });
    await act(async () => { fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [third], types: ["Files"] } }); });
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
    const picker = document.querySelector('input[type="file"][accept*=".docx"]')!;
    expect(picker).toHaveAttribute("accept", expect.stringContaining(".xlsx"));
    const book = new File([new Uint8Array([1])], "tender.XLSX", { type: mime });
    await selectTemplate(user, book);
    expect(screen.getByText("Template XLSX preview")).toBeInTheDocument();
    await expectTemplateBytes(xlsxSources.at(-1), book);
    expect(previewSources.at(-1)).not.toBe(book);
    expectResultOutdated();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    const dropped = new File([new Uint8Array([2])], "second.xlsx");
    await act(async () => { fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [dropped], types: ["Files"] } }); });
    await expectTemplateBytes(xlsxSources.at(-1), dropped);
    expect(screen.getByText("second.xlsx")).toBeInTheDocument();
    await act(async () => { fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [new File([], "bad.xls")], types: ["Files"] } }); });
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
    await act(async () => {
      resolve(success(new Uint8Array(await responseBlob.arrayBuffer()), mime, "document.docx"));
    });
    await screen.findByText("Result preview");
    expect(screen.getByText("sample.docx")).toBeInTheDocument();
    expect(screen.queryByText("sample.template.docx")).not.toBeInTheDocument();
    expect(new Uint8Array(await previewSources.at(-1)!.arrayBuffer())).toEqual(new Uint8Array(await responseBlob.arrayBuffer()));
    let downloadName = "";
    const click = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) { downloadName = this.download; });
    await documentCommand(user, "Download");
    const downloaded = vi.mocked(URL.createObjectURL).mock.calls.at(-1)![0] as Blob;
    expect(new Uint8Array(await downloaded.arrayBuffer())).toEqual(new Uint8Array(await responseBlob.arrayBuffer()));
    expect(click).toHaveBeenCalled();
    expect(downloadName).toBe("sample.docx");
  });

  it("shows ordered generation warnings as Problems while keeping successful status and document", async () => {
    const user = userEvent.setup();
    const warnings = [
      { message: "Removed 2 prohibited XML 1.0 characters.", field: "{Name}", fieldCode: "{Name}", location: "A1", part: "/xl/worksheets/sheet1.xml", dataPath: "$.Name", code: "generation.xml_characters_removed", removedCharacterCount: 2 },
      { message: "Another warning.", field: "{Other}", dataPath: "$.Other" }
    ];
    mockPlaygroundFetch(async () => success(new Uint8Array([80, 75, 1, 2, 255]), mime, "document.docx", warnings));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");

    await user.click(screen.getByRole("button", { name: "Generate document" }));

    expect(await screen.findByRole("tab", { name: "Problems 2" })).toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    expect(screen.getByText(/removed 2/)).toBeInTheDocument();
    const issues = [...document.querySelectorAll(".issue-heading strong")].map(node => node.textContent);
    expect(issues).toEqual(["generation.xml_characters_removed", "Generation warning"]);
    expect(screen.getByText("Result preview")).toBeInTheDocument();
  });

  it("keeps the previous successful document when a later multipart response is malformed", async () => {
    const user = userEvent.setup();
    const responses = [success(), new Response("--broken", { status: 200, headers: { "Content-Type": "multipart/mixed; boundary=broken" } })];
    const { apiMock } = mockPlaygroundFetch(async () => responses.shift()!);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    const previousBytes = new Uint8Array(await previewSources.at(-1)!.arrayBuffer());
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"Project":{"Name":"changed"}}' } });

    await user.click(screen.getByRole("button", { name: "Generate document" }));

    expect(await screen.findByText(/invalid document and diagnostics response/)).toBeInTheDocument();
    expect(screen.getByText("Result preview")).toBeInTheDocument();
    expect(new Uint8Array(await previewSources.at(-1)!.arrayBuffer())).toEqual(previousBytes);
    expect(apiMock).toHaveBeenCalledTimes(2);
  });

  it("manually generates XLSX without analysis, previews the response format, and keeps the original download after preview failure", async () => {
    const user = userEvent.setup();
    previewTestState.failResultPreview = true;
    const resultBlob = new Blob([new Uint8Array([80, 75, 7, 8, 9])], { type: xlsxMime });
    const { apiMock } = mockPlaygroundFetch(async () => success(new Uint8Array(await resultBlob.arrayBuffer()), xlsxMime, "document.xlsx"));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    const sourceFile = new File([new Uint8Array([1, 2])], "manual.xlsx", { type: xlsxMime });
    await selectTemplate(user, sourceFile);
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"manual":true}' } });
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(await screen.findByText("Result XLSX preview failed")).toBeInTheDocument();
    expect(screen.getByRole("radio", { name: "Result" })).toBeChecked();
    expect(new Uint8Array(await xlsxSources.at(-1)!.arrayBuffer())).toEqual(new Uint8Array(await resultBlob.arrayBuffer()));
    expect(apiMock).toHaveBeenCalledOnce();
    expect(String((apiMock.mock.calls[0] as [string])[0])).toBe("/api/documents/render");
    const requestBody = (apiMock.mock.calls[0] as [string, RequestInit])[1].body as FormData;
    await expectTemplateBytes(requestBody.get("template"), sourceFile);
    expect(requestBody.get("data")).toBe('{"manual":true}');

    let downloadedFilename = "";
    const click = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) { downloadedFilename = this.download; });
    await documentCommand(user, "Download");
    const downloaded = vi.mocked(URL.createObjectURL).mock.calls.at(-1)![0] as Blob;
    expect(new Uint8Array(await downloaded.arrayBuffer())).toEqual(new Uint8Array(await resultBlob.arrayBuffer()));
    expect(click).toHaveBeenCalled();
    expect(downloadedFilename).toBe("manual-result.xlsx");
  });

  it("renders real API diagnostics and multiple issues without mock content", async () => {
    const user = userEvent.setup();
    mockPlaygroundFetch(async () => new Response(JSON.stringify({ traceId: "trace-123", limitReached: true, errors: [{ code: "first", message: "First issue", field: "Amount", dataPath: "Items[3].Amount", part: "/xl/worksheets/sheet2.xml", location: "C14" }, { code: "second", detail: "Second issue" }] }), { status: 422, headers: { "Content-Type": "application/problem+json" } }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime }));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(await screen.findByRole("tab", { name: "Problems 2" })).toBeInTheDocument();
    expect(screen.queryByText("2 problems")).not.toBeInTheDocument();
    expect(screen.getByText(/First issue/)).toBeInTheDocument();
    expect(screen.getByText(/Second issue/)).toBeInTheDocument();
    expect(screen.getByText(/Items\[3\]\.Amount/)).toBeInTheDocument();
    expect(screen.getByText(/\/xl\/worksheets\/sheet2\.xml · C14/)).toBeInTheDocument();
    expect(screen.getAllByText(/trace trace-123/)).toHaveLength(2);
    expect(screen.getAllByText(/error limit reached/)).toHaveLength(2);
    expect(screen.queryByText("Unmatched collection region")).not.toBeInTheDocument();
  });

  it("retains exact successful bytes after edits, invalid JSON, and API failure while reporting freshness", async () => {
    const user = userEvent.setup();
    const response = success();
    const expectedDocumentBytes = new Uint8Array([80, 75, 7, 8]);
    const responses = [response,
      new Response(JSON.stringify({ title: "Invalid request", detail: "The data field must contain one valid JSON value.", code: "data.invalid_json" }), { status: 400, headers: { "Content-Type": "application/problem+json" } }),
      new Response(JSON.stringify({ title: "Failed", detail: "Incompatible", code: "generation.invalid_data" }), { status: 422, headers: { "Content-Type": "application/problem+json" } })];
    const { apiMock } = mockPlaygroundFetch(async () => responses.shift()!);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "{}" } });
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    expectResultOutdated(false);
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    await dataCommand(user, "Format");
    expectResultOutdated(false);
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "{" } });
    expectResultOutdated();
    expect(screen.getByText("Result preview")).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expectResultOutdated();
    expect(screen.getByRole("status")).toHaveTextContent("Error");
    expect(screen.getByText("data.invalid_json", { selector: ".issue-heading strong" })).toBeInTheDocument();
    expect(apiMock).toHaveBeenCalledTimes(2);
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: "{}" } });
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(await screen.findByText(/generation.invalid_data/)).toBeInTheDocument();
    expectResultOutdated(false);
    expect(screen.getByRole("status")).toHaveTextContent("Error");
    expect(screen.getByRole("radio", { name: "Result" })).toBeChecked();
    expect(new Uint8Array(await previewSources.at(-1)!.arrayBuffer())).toEqual(expectedDocumentBytes);
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(previewSources.at(-1));
  });

  it("reports HTTP 413 for a large Dataset while retaining the Dataset and previous result preview", async () => {
    const user = userEvent.setup();
    const oversizedProblem = new Response(JSON.stringify({
      title: "Request too large",
      detail: "The whole request exceeds the configured size limit.",
      code: "request.too_large"
    }), { status: 413, headers: { "Content-Type": "application/problem+json" } });
    const responses = [success(undefined, xlsxMime), oversizedProblem];
    const { apiMock } = mockPlaygroundFetch(async () => responses.shift()!);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime }));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result XLSX preview");
    const previousResult = xlsxSources.at(-1)!;

    const largeDataset = JSON.stringify({ Rows: Array.from({ length: 20_000 }, (_, index) => `Sample Name ${index + 1}`) });
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: largeDataset } });
    await user.click(screen.getByRole("button", { name: "Generate document" }));

    expect(await screen.findByText("request.too_large", { selector: ".issue-heading strong" })).toBeInTheDocument();
    expect(screen.getAllByText("request.too_large", { selector: ".issue-heading strong" })).toHaveLength(1);
    expect(screen.getByLabelText("Dataset")).toHaveValue(largeDataset);
    expect(screen.getByText("Result XLSX preview")).toBeInTheDocument();
    expect(xlsxSources.at(-1)).toBe(previousResult);
    expect(screen.getByRole("status")).toHaveTextContent("Error");
    expect(screen.getByRole("radio", { name: "Result" })).toBeChecked();
    expectResultOutdated();
    expect(apiMock).toHaveBeenCalledTimes(2);
    const [url, init] = apiMock.mock.calls[1] as [string, RequestInit];
    expect(url).toBe("/api/documents/render");
    expect(((init.body as FormData).get("data") as string)).toBe(largeDataset);
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(previousResult);
    expect(new Uint8Array(await previousResult.arrayBuffer())).toEqual(new Uint8Array([80, 75, 7, 8]));
  });

  it("keeps the previous XLSX result when a later successful HTTP response has truncated multipart framing", async () => {
    const user = userEvent.setup();
    const responses = [success(undefined, xlsxMime), new Response("--broken", {
      status: 200,
      headers: { "Content-Type": "multipart/mixed; boundary=broken" }
    })];
    mockPlaygroundFetch(async () => responses.shift()!);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime }));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result XLSX preview");
    const previousResult = xlsxSources.at(-1)!;

    await user.click(screen.getByRole("button", { name: "Generate document" }));

    expect(await screen.findByText("Unexpected response", { selector: ".issue-heading strong" })).toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent("Error");
    expect(xlsxSources.at(-1)).toBe(previousResult);
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(previousResult);
  });

  it("retains XLSX result format and download bytes after switching and closing the template", async () => {
    const user = userEvent.setup();
    const response = success(undefined, xlsxMime);
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
    expect(screen.getByText("Result XLSX preview")).toBeInTheDocument();
    expectResultOutdated();
    await user.click(screen.getByRole("radio", { name: "Template" }).closest("label")!);
    await documentCommand(user, "Close");
    expect(screen.getByText("Select a DOCX or XLSX template")).toBeInTheDocument();
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(screen.getByText("Result XLSX preview")).toBeInTheDocument();
    expect(new Uint8Array(await xlsxSources.at(-1)!.arrayBuffer())).toEqual(new Uint8Array([80, 75, 7, 8]));
    let downloadedFilename = "";
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) { downloadedFilename = this.download; });
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(xlsxSources.at(-1));
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
    expect(screen.getByRole("status")).toHaveTextContent("Error");
    await user.click(screen.getByRole("button", { name: "Hide diagnostics" }));
    expect(localStorage.getItem(workspaceStorage.diagnosticsCollapsed)).toBe("true");
    const toggle = screen.getByRole("button", { name: "Problems 1" });
    expect(toggle).toHaveAttribute("aria-expanded", "false");
    expect(preview).toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent("Error");
    await user.click(toggle);
    expect(localStorage.getItem(workspaceStorage.diagnosticsCollapsed)).toBe("false");
    expect(toggle).not.toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Problems 1" })).toBeInTheDocument();
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
    await act(async () => {
      resolve(success());
    });
    expect(screen.getByRole("status")).toHaveTextContent("Cancelled");
    expect(screen.queryByText("Result preview")).not.toBeInTheDocument();
    expect(screen.getByRole("radio", { name: "Template" })).toBeChecked();
  });

  it("keeps the newer XLSX generation when an older request resolves afterward", async () => {
    const user = userEvent.setup();
    const older = success(new Uint8Array([1]), xlsxMime);
    const newer = success(new Uint8Array([2]), xlsxMime);
    const resolvers: Array<(response: Response) => void> = [];
    const { apiMock } = mockPlaygroundFetch(() => new Promise<Response>(resolve => resolvers.push(resolve)));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime }));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"revision":2}' } });
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(apiMock).toHaveBeenCalledTimes(2);

    await act(async () => {
      resolvers[1](newer);
    });
    expect(await screen.findByText("Result XLSX preview")).toBeInTheDocument();
    await act(async () => {
      resolvers[0](older);
    });
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Success"));
    expect(new Uint8Array(await xlsxSources.at(-1)!.arrayBuffer())).toEqual(new Uint8Array([2]));
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
    await act(async () => { fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [xlsx], types: ["Files"] } }); });
    await act(async () => {
      resolveGeneration(success());
    });

    await waitFor(() => expect(screen.getByText("replacement.xlsx")).toBeInTheDocument());
    expect(screen.getByText("Template XLSX preview")).toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent("Cancelled");
  });

  it("ignores an analysis response after the template is cleared", async () => {
    const user = userEvent.setup();
    let resolveAnalysis!: (response: Response) => void;
    mockPlaygroundFetch(() => new Promise<Response>(resolve => { resolveAnalysis = resolve; }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime }));
    const editor = screen.getByLabelText("Dataset");
    await analyzeTemplate(user);
    await documentCommand(user, "Close");
    await act(async () => {
      resolveAnalysis(analysis([{ name: "Stale", type: "value" }]));
    });

    await waitFor(() => expect(screen.getByText("Select a DOCX or XLSX template")).toBeInTheDocument());
    expect(editor).toHaveValue(sampleData);
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    expect(await screen.findByRole("menuitem", { name: "Analyze Template" })).toHaveAttribute("aria-disabled", "true");
    expect(screen.getByRole("status")).toHaveTextContent("Cancelled");
    expect(screen.getByText("No analysis")).toBeInTheDocument();
  });

  it("replaces a clean Dataset without confirmation and retains an outdated Result", async () => {
    const user = userEvent.setup();
    const responses = [success(), sampleJson()];
    const { apiMock } = mockPlaygroundFetch(async () => responses.shift()!);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await selectTemplate(user);
    expect(apiMock).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");

    await generateJson(user, 1);
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();

    await waitFor(() => expect(screen.getByLabelText("Dataset")).toHaveValue(JSON.stringify({ Company: { Name: "Sample Name 1" } }, null, 2)));
    expectResultOutdated();
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    expect(screen.getByText("Result preview")).toBeInTheDocument();
    const [sampleUrl, sampleInit] = apiMock.mock.calls[1] as [string, RequestInit];
    expect(sampleUrl).toBe("/api/documents/sample");
    expect(((sampleInit.body as FormData).get("template") as File).name).toBe("sample.template.docx");
  });

  it("requests an XLSX sample from the server without posting analysis data", async () => {
    const user = userEvent.setup();
    const xlsx = new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime });
    const { apiMock } = mockPlaygroundFetch(async () => sampleJson());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user, xlsx);
    const editor = screen.getByLabelText("Dataset");
    fireEvent.change(editor, { target: { value: "" } });
    await generateJson(user, 1);

    await waitFor(() => expect(editor).toHaveValue(JSON.stringify({ Company: { Name: "Sample Name 1" } }, null, 2)));
    expect(apiMock).toHaveBeenCalledOnce();
    const [url, init] = apiMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe("/api/documents/sample");
    await expectTemplateBytes((init.body as FormData).get("template"), xlsx);
    expect((init.body as FormData).get("itemsPerCollection")).toBe("1");
  });

  it("protects dirty Dataset content when generated data would replace it", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => sampleJson(undefined, [{ message: "Review generated values" }]));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const editor = screen.getByLabelText("Dataset");
    fireEvent.change(editor, { target: { value: "keep me" } });
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const settings = await screen.findByRole("dialog", { name: "Generate Dataset" });
    fireEvent.change(within(settings).getByRole("spinbutton", { name: "Items per collection" }), { target: { value: "3" } });
    await user.click(within(settings).getByRole("radio", { name: "Empty" }));
    await user.click(within(settings).getByRole("button", { name: "Generate sample" }));
    await waitFor(() => expect(within(settings).getByLabelText("Sample JSON preview")).toBeInTheDocument());
    expect(within(settings).queryByText(/discard unsaved changes/)).not.toBeInTheDocument();
    expect(editor).toHaveValue("keep me");
    expect(apiMock).toHaveBeenCalledOnce();
    vi.spyOn(navigator.clipboard, "writeText").mockRejectedValueOnce(new Error("Clipboard denied"));
    await user.click(within(settings).getByRole("button", { name: "Copy" }));
    expect(await screen.findByText("The sample JSON could not be copied to the clipboard.")).toBeInTheDocument();
    expect(within(settings).getByLabelText("Sample JSON preview")).toBeInTheDocument();
    await user.click(within(settings).getByRole("button", { name: "Replace Dataset" }));
    expect(await screen.findByRole("dialog", { name: "Replace Dataset?" })).toHaveTextContent("Replacing the Dataset will discard unsaved changes.");
    expect(editor).toHaveValue("keep me");
    await confirmSampleReplacement(user, false);
    expect(editor).toHaveValue("keep me");
    expect(settings).toBeVisible();
    expect(within(settings).getByRole("spinbutton", { name: "Items per collection" })).toHaveValue("3");
    expect(within(settings).getByRole("radio", { name: "Empty" })).toBeChecked();
    expect(within(settings).getByLabelText("Sample JSON preview")).toHaveValue(JSON.stringify({ Company: { Name: "Sample Name 1" } }, null, 2));
    expect(within(settings).getByText(/Review generated values/)).toBeInTheDocument();
    expect(apiMock).toHaveBeenCalledOnce();
    await user.click(within(settings).getByRole("button", { name: "Replace Dataset" }));
    await confirmSampleReplacement(user);
    await waitFor(() => expect(editor).toHaveValue(JSON.stringify({ Company: { Name: "Sample Name 1" } }, null, 2)));
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
  });

  it.each(["dataset", "settings", "template", "dialog"])("ignores a sample replacement confirmation after the %s changes", async change => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => sampleJson());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const editor = screen.getByLabelText("Dataset");
    fireEvent.change(editor, { target: { value: "keep me" } });
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const settings = await screen.findByRole("dialog", { name: "Generate Dataset" });
    await user.click(within(settings).getByRole("button", { name: "Generate sample" }));
    await user.click(await within(settings).findByRole("button", { name: "Replace Dataset" }));
    await screen.findByRole("dialog", { name: "Replace Dataset?" });

    if (change === "dataset") fireEvent.change(editor, { target: { value: "newer edits" } });
    if (change === "settings") fireEvent.change(within(settings).getByRole("spinbutton", { name: "Items per collection" }), { target: { value: "2" } });
    if (change === "template") fireEvent.change(document.querySelector('input[type="file"][accept*=".docx"]')!, { target: { files: [new File([new Uint8Array([80, 75])], "replacement.docx", { type: mime })] } });
    if (change === "dialog") fireEvent.click(within(settings).getByRole("button", { name: "Close" }));

    await confirmSampleReplacement(user);
    expect(editor).toHaveValue(change === "dataset" ? "newer edits" : "keep me");
    expect(apiMock).toHaveBeenCalledOnce();
    if (change !== "dialog") expect(settings).toBeVisible();
  });

  it("never applies stale analysis after JSON or template changes", async () => {
    const user = userEvent.setup();
    const resolvers: Array<(response: Response) => void> = [];
    mockPlaygroundFetch(() => new Promise<Response>(resolve => resolvers.push(resolve)));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await selectTemplate(user);
    const editor = screen.getByLabelText("Dataset");
    await analyzeTemplate(user);
    fireEvent.change(editor, { target: { value: "edited during analysis" } });
    await act(async () => {
      resolvers[0](analysis());
    });
    expect(editor).toHaveValue("edited during analysis");
    expect(screen.getByText("No analysis")).toBeInTheDocument();
    expect(screen.queryByText("Company")).not.toBeInTheDocument();

    await analyzeTemplate(user);
    const replacement = new File([new Uint8Array([1])], "replacement.docx", { type: mime });
    await act(async () => { fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [replacement], types: ["Files"] } }); });
    await act(async () => {
      resolvers[1](analysis([{ name: "Stale", type: "value" }]));
    });
    await waitFor(() => expect(screen.getByText("replacement.docx")).toBeInTheDocument());
    expect(editor).toHaveValue("edited during analysis");
    expect(screen.getByText("No analysis")).toBeInTheDocument();
    expect(screen.queryByText("Stale")).not.toBeInTheDocument();
  });

  it("ignores an in-flight XLSX analysis after switching to a DOCX template", async () => {
    const user = userEvent.setup();
    let resolveAnalysis!: (response: Response) => void;
    mockPlaygroundFetch(() => new Promise<Response>(resolve => { resolveAnalysis = resolve; }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user, new File([new Uint8Array([80, 75])], "source.xlsx", { type: xlsxMime }));
    const editor = screen.getByLabelText("Dataset");
    await analyzeTemplate(user);
    const docx = template();
    await act(async () => { fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [docx], types: ["Files"] } }); });
    await act(async () => {
      resolveAnalysis(analysis([{ name: "Stale XLSX", type: "value" }]));
    });

    await waitFor(() => expect(screen.getByText(docx.name)).toBeInTheDocument());
    expect(editor).toHaveValue(sampleData);
    expect(screen.getByText("No analysis")).toBeInTheDocument();
    expect(screen.queryByText("Stale XLSX")).not.toBeInTheDocument();
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
    await analyzeTemplate(user);
    const xlsx = new File([new Uint8Array([80, 75])], "replacement.xlsx", { type: xlsxMime });
    await act(async () => { fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [xlsx], types: ["Files"] } }); });
    await act(async () => {
      resolveAnalysis(analysis([{ name: "Stale DOCX", type: "value" }]));
    });

    await waitFor(() => expect(screen.getByText(xlsx.name)).toBeInTheDocument());
    expect(editor).toHaveValue(sampleData);
    expect(screen.getByText("No analysis")).toBeInTheDocument();
    expect(screen.queryByText("Stale DOCX")).not.toBeInTheDocument();
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
    await analyzeTemplate(user);
    expect(await screen.findByText("template.unmatched_brace", { selector: ".issue-heading strong" })).toBeInTheDocument();
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
    await act(async () => {
      resolveTemplate(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
      resolveData(new Response('{"stale":true}', { status: 200 }));
    });
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
    await act(async () => {
      resolveTemplate(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
      resolveData(new Response('{"stale":true}', { status: 200 }));
    });
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
    await act(async () => {
      resolveTemplate(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
      resolveData(new Response(sampleData, { status: 200 }));
    });
    await waitFor(() => expect(screen.getByText("sample.template.docx")).toBeInTheDocument());
    await act(async () => {
      resolveGenerate(success());
    });
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Cancelled"));
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
    await act(async () => {
      resolveTemplate(new Response(new Uint8Array([1, 2]), { status: 200, headers: { "Content-Type": mime } }));
      resolveData(new Response('{"second":true}', { status: 200 }));
    });
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
  it("opens dataset generation settings with defaults and cancels without requesting or changing data", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    expect(await screen.findByRole("menuitem", { name: "Analyze Template" })).toBeInTheDocument();
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog");
    const count = within(dialog).getByRole("spinbutton", { name: "Items per collection" });
    expect(within(dialog).getByRole("radio", { name: "Filled" })).toBeInTheDocument();
    expect(within(dialog).getByText("Values", { exact: true })).toBeInTheDocument();
    expect(count).toHaveValue("10");
    expect(count).not.toHaveAttribute("aria-valuemax");
    expect(within(dialog).getByRole("radio", { name: "Filled" })).toBeChecked();
    await user.click(within(dialog).getByRole("button", { name: "Close" }));
    expect(apiMock).not.toHaveBeenCalled();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.getByRole("status")).toHaveTextContent("Ready");
  });

  it("generates from Cmd+Enter or Ctrl+Enter, including Monaco focus, and prevents the browser default", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => success());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    const editor = screen.getByLabelText("Dataset");
    await waitFor(() => expect(editor).toHaveValue(sampleData));

    expect(fireEvent.keyDown(editor, { key: "Enter", metaKey: true })).toBe(false);
    await waitFor(() => expect(apiMock).toHaveBeenCalledTimes(1));
    await screen.findByText("Result preview");

    expect(fireEvent.keyDown(document, { key: "Enter", ctrlKey: true })).toBe(false);
    await waitFor(() => expect(apiMock).toHaveBeenCalledTimes(2));
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Success"));
  });

  it("closes the sample review without replacing the current dataset", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => sampleJson());
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog");
    fireEvent.click(within(dialog).getByRole("button", { name: "Generate sample" }));
    fireEvent.click(within(dialog).getByRole("button", { name: "Close" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).not.toHaveBeenCalled();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
  });

  it("rejects invalid settings without a request or dataset", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog");
    const count = within(dialog).getByRole("spinbutton", { name: "Items per collection" });
    for (const value of ["0", "-2", "1.5", ""]) {
      fireEvent.change(count, { target: { value } });
      await user.click(within(dialog).getByRole("button", { name: "Generate sample" }));
      expect(count).toHaveValue(value);
    }
    expect(apiMock).not.toHaveBeenCalled();
    await user.click(within(dialog).getByRole("button", { name: "Close" }));
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.getByRole("status")).toHaveTextContent("Ready");
  });

  it("renders actual nested analysis, occurrences and warnings, then reuses it across JSON edits and dataset creation", async () => {
    const user = userEvent.setup();
    const payload = {
      valid: true, errors: [],
      retainedWireField: "preserved",
      warnings: [
        { code: "template.notice", message: "A supplied template warning", field: "Rows", part: "/xl/worksheets/sheet1.xml", location: "C4", formatter: "format", dataPath: "$.Rows" },
        { code: "template.distinct", title: "Distinct warning title", message: "Another template warning" }
      ],
      requirements: [{ name: "Rows", type: "array", fields: [{ name: "Name", type: "value" }, { name: "Enabled", type: "boolean" }, { name: "Tags", type: "array" }] }],
      occurrences: [
        { kind: "for", expression: "{For:Rows}", path: "$.Rows", part: "/xl/worksheets/sheet1.xml", location: "C4", semanticType: "number" },
        { kind: "if", expression: "{If:Rows[0].Description contains 'a long value'}", path: "$.Rows[0].Description/Details", part: "/xl/worksheets/sheet1.xml", location: "C5" }
      ]
    };
    const { apiMock } = mockPlaygroundFetch(async input => String(input) === "/api/documents/sample"
      ? sampleJson()
      : new Response(JSON.stringify(payload), { status: 200 }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await analyzeTemplate(user);
    const panel = await screen.findByRole("tabpanel", { name: "Analysis" });
    expect(await within(panel).findByText("Valid", { exact: true })).toBeInTheDocument();
    for (const text of ["Rows", "Name", "Enabled", "Tags", "{For:Rows}", "A supplied template warning", "format"]) expect(within(panel).getAllByText(text, { exact: true }).length).toBeGreaterThan(0);
    for (const type of ["array", "value", "boolean"]) expect(within(panel).getAllByText(type, { exact: true }).length).toBeGreaterThan(0);
    expect(within(panel).getByText("{If:Rows[0].Description contains 'a long value'}")).toBeInTheDocument();
    expect(within(panel).getByText("$.Rows[0].Description/Details")).toBeInTheDocument();
    expect(within(panel).getAllByText("C4")).toHaveLength(2);
    expect(within(panel).getByText("number")).toBeInTheDocument();
    expect(within(panel).getByText("Requirement types describe JSON structure; semantic types are advisory and describe value meaning.")).toBeInTheDocument();
    expect(within(panel).getByText("template.notice", { exact: true })).toBeInTheDocument();
    expect(within(panel).getByText("Distinct warning title", { exact: true })).toBeInTheDocument();
    expect(within(panel).getByText("template.distinct", { exact: true })).toBeInTheDocument();
    expect(within(panel).getByRole("columnheader", { name: "Semantic type" })).toBeInTheDocument();
    const clipboard = vi.spyOn(navigator.clipboard, "writeText").mockResolvedValue(undefined);
    const analysisActions = screen.getByRole("toolbar", { name: "Analysis actions" });
    expect(await within(analysisActions).findByRole("button", { name: "Copy" })).toBeEnabled();
    await user.click(within(analysisActions).getByRole("button", { name: "Copy" }));
    await waitFor(() => expect(clipboard).toHaveBeenCalledWith(JSON.stringify(payload, null, 2)));
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.queryByText("Result preview")).not.toBeInTheDocument();
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"keep":"dirty"}' } });
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog");
    fireEvent.change(within(dialog).getByRole("spinbutton", { name: "Items per collection" }), { target: { value: "3" } });
    await user.click(within(dialog).getByRole("radio", { name: "Filled" }));
    await user.click(within(dialog).getByRole("button", { name: "Generate sample" }));
    await user.click(await within(dialog).findByRole("button", { name: "Replace Dataset" }));
    await confirmSampleReplacement(user);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    const data = JSON.parse((screen.getByLabelText("Dataset") as HTMLTextAreaElement).value);
    expect(data).toEqual({ Company: { Name: "Sample Name 1" } });
    expect(apiMock).toHaveBeenCalledTimes(2);
    await generateJson(user, 2);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).toHaveBeenCalledTimes(3);
    expect(JSON.parse((screen.getByLabelText("Dataset") as HTMLTextAreaElement).value)).toEqual({ Company: { Name: "Sample Name 1" } });
    expect(selectedExample()).toHaveTextContent("sample-generated.json");
    await diagnosticsTab(user, "Analysis");
    expect(await screen.findByText("A supplied template warning")).toBeInTheDocument();
  });

  it("preserves displayed invalid analysis while sample generation uses its own endpoint", async () => {
    const user = userEvent.setup();
    const payload = {
      valid: false,
      errors: [{ code: "template.bad", message: "Bad template block", field: "{Name", dataPath: "$.Name", part: "/word/document.xml", location: "paragraph 8" }],
      warnings: [{ code: "template.warning", message: "Additional warning" }],
      requirements: [{ name: "Name", type: "value" }],
      occurrences: [{ kind: "field", expression: "{Name", path: "$.Name", part: "/word/document.xml", location: "paragraph 8" }]
    };
    let first = true;
    const { apiMock } = mockPlaygroundFetch(async input => {
      if (String(input) === "/api/documents/sample") return sampleJson();
      if (first) { first = false; return analysis(); }
      return new Response(JSON.stringify(payload), { status: 200 });
    });
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await analyzeTemplate(user);
    expect(await screen.findByText("Company")).toBeInTheDocument();
    await analyzeTemplate(user);
    await diagnosticsTab(user, "Analysis");
    const panel = await screen.findByRole("tabpanel", { name: "Analysis" });
    expect(await within(panel).findByText("Invalid", { exact: true })).toBeInTheDocument();
    expect(within(panel).getByText("Name")).toBeInTheDocument();
    expect(within(panel).getAllByText("$.Name")).toHaveLength(2);
    expect(within(panel).getByText("Additional warning")).toBeInTheDocument();
    expect(within(panel).getAllByText("paragraph 8")).toHaveLength(2);
    expect(screen.getByRole("status")).toHaveTextContent("Error");
    await generateJson(user);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).toHaveBeenCalledTimes(3);
    expect(screen.getByLabelText("Dataset")).toHaveValue(JSON.stringify({ Company: { Name: "Sample Name 1" } }, null, 2));
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    await diagnosticsTab(user, "Analysis");
    expect(screen.getByText("Bad template block")).toBeInTheDocument();
    expect(screen.getByText("Additional warning")).toBeInTheDocument();
  });

  it("keeps the displayed and copied analysis paired when a later request has no analysis payload", async () => {
    const user = userEvent.setup();
    const firstWire = { valid: true, errors: [], warnings: [], requirements: [{ name: "Original", type: "value" }], occurrences: [], serverRevision: "first" };
    const responses = [
      new Response(JSON.stringify(firstWire), { status: 200 }),
      new Response(JSON.stringify({ title: "Analysis failed", detail: "Temporary failure", code: "request.invalid_input" }), { status: 422 })
    ];
    mockPlaygroundFetch(async () => responses.shift()!);
    const clipboard = vi.spyOn(navigator.clipboard, "writeText").mockResolvedValue(undefined);
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await analyzeTemplate(user);
    expect(await screen.findByText("Original")).toBeInTheDocument();
    await analyzeTemplate(user);
    expect(await screen.findByText("Temporary failure")).toBeInTheDocument();
    await diagnosticsTab(user, "Analysis");

    const panel = await screen.findByRole("tabpanel", { name: "Analysis" });
    expect(within(panel).getByText("Valid", { exact: true })).toBeInTheDocument();
    expect(within(panel).getByText("Original")).toBeInTheDocument();
    await user.click(within(screen.getByRole("toolbar", { name: "Analysis actions" })).getByRole("button", { name: "Copy" }));

    await waitFor(() => expect(clipboard).toHaveBeenCalledWith(JSON.stringify(firstWire, null, 2)));
  });

  it("clears analysis when the template changes and performs a fresh analysis for new data", async () => {
    const user = userEvent.setup();
    const clipboard = vi.spyOn(navigator.clipboard, "writeText").mockResolvedValue(undefined);
    const { apiMock } = mockPlaygroundFetch(async input => String(input) === "/api/documents/sample"
      ? sampleJson()
      : analysis([{ name: "Rows", type: "array", fields: [{ name: "Name", type: "value" }] }]));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await analyzeTemplate(user);
    expect(await screen.findByText("Rows")).toBeInTheDocument();
    const replacement = new File([new Uint8Array([80, 75])], "new.xlsx", { type: xlsxMime });
    await selectTemplate(user, replacement);
    expect(screen.queryByText("Rows")).not.toBeInTheDocument();
    expect(screen.getByText("No analysis")).toBeInTheDocument();
    await diagnosticsTab(user, "Analysis");
    await user.click(within(screen.getByRole("toolbar", { name: "Analysis actions" })).getByRole("button", { name: "Copy" }));
    await waitFor(() => expect(clipboard).toHaveBeenCalledWith("{}"));
    await generateJson(user);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(JSON.parse((screen.getByLabelText("Dataset") as HTMLTextAreaElement).value)).toEqual({ Company: { Name: "Sample Name 1" } });
    expect(selectedExample()).toHaveTextContent("new-generated.json");
    expect(apiMock).toHaveBeenCalledTimes(2);
    expect(apiMock.mock.calls[1][0]).toBe("/api/documents/sample");
    await expectTemplateBytes(((apiMock.mock.calls[1][1] as RequestInit).body as FormData).get("template"), replacement);
  });

  it("rejects server sample failures without a candidate and accepts a later sample", async () => {
    const user = userEvent.setup();
    let samples = 0;
    const { apiMock } = mockPlaygroundFetch(async input => {
      if (String(input) === "/api/documents/sample" && samples++ === 0) return new Response(JSON.stringify({
        title: "Sample generation failed",
        detail: "The 1,000,000 element limit was exceeded.",
        code: "sample.element_limit",
        errors: [
          { code: "sample.element_limit", message: "The 1,000,000 element limit was exceeded.", part: "/word/document.xml", location: "paragraph 3" },
          { code: "sample.template_issue", message: "A separate template limitation.", location: "paragraph 4" }
        ]
      }), { status: 422 });
      return sampleJson();
    });
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: '{"keep":true}' } });
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog");
    await user.click(within(dialog).getByRole("button", { name: "Generate sample" }));
    expect(await within(dialog).findByText(/The 1,000,000 element limit was exceeded/)).toBeInTheDocument();
    expect(within(dialog).getAllByText(/The 1,000,000 element limit was exceeded/)).toHaveLength(1);
    expect(within(dialog).getByText(/\/word\/document.xml/)).toBeInTheDocument();
    expect(within(dialog).getByText(/A separate template limitation/)).toBeInTheDocument();
    expect(within(dialog).queryByRole("button", { name: "Replace Dataset" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue('{"keep":true}');
    expect(selectedExample()).toHaveTextContent("sample.json");
    await user.click(within(dialog).getByRole("button", { name: "Close" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    await generateJson(user, 2);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).toHaveBeenCalledTimes(2);
    expect(JSON.parse((screen.getByLabelText("Dataset") as HTMLTextAreaElement).value)).toEqual({ Company: { Name: "Sample Name 1" } });
    expect(selectedExample()).toHaveTextContent("sample-generated.json");
  });

  it("clears sample diagnostics on replacement and retains later Generate diagnostics", async () => {
    const user = userEvent.setup();
    const sampleWarning = { code: "sample.unresolved_shape", message: "Sample-only warning", location: "paragraph 2" };
    const generateWarning = { message: "Generate warning", code: "generation.notice", field: "Name" };
    const { apiMock } = mockPlaygroundFetch(async input => String(input) === "/api/documents/sample"
      ? sampleJson({}, [sampleWarning])
      : success(new Uint8Array([80, 75, 7, 8]), mime, "document.docx", [generateWarning]));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog", { name: "Generate Dataset" });
    await user.click(within(dialog).getByRole("button", { name: "Generate sample" }));
    expect(await within(dialog).findByText(/Sample-only warning/)).toBeInTheDocument();
    expect(within(dialog).queryByText(/discard unsaved changes/)).not.toBeInTheDocument();
    expect(within(dialog).getByLabelText("Sample JSON preview")).toHaveAttribute("readonly");
    await user.click(await within(dialog).findByRole("button", { name: "Replace Dataset" }));
    await diagnosticsTab(user, "Problems");
    expect(await screen.findByText("No problems")).toBeInTheDocument();

    await user.click(screen.getByRole("button", { name: "Generate document" }));

    await waitFor(() => expect(apiMock).toHaveBeenCalledTimes(2));
    expect(await screen.findByRole("tab", { name: "Problems 1" })).toBeInTheDocument();
    expect(document.querySelector(".issue p")).toHaveTextContent("Generate warning");
    expect(screen.queryByText("Sample-only warning")).not.toBeInTheDocument();
  });

  it("shows compact operation status and measured duration without an Output history", async () => {
    const user = userEvent.setup();
    let now = 1000;
    vi.spyOn(performance, "now").mockImplementation(() => now);
    const resolvers: Array<(response: Response) => void> = [];
    const { apiMock } = mockPlaygroundFetch(() => new Promise<Response>(resolve => resolvers.push(resolve)));
    const view = render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const sampleDialog = await screen.findByRole("dialog");
    await user.click(within(sampleDialog).getByRole("button", { name: "Generate sample" }));
    expect(screen.getByRole("status")).toHaveTextContent("Generating sample...");
    now = 1125;
    await act(async () => {
      resolvers[0](sampleJson());
    });
    await within(sampleDialog).findByRole("button", { name: "Replace Dataset" });
    await user.click(within(sampleDialog).getByRole("button", { name: "Replace Dataset" }));
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    expect(screen.getByRole("contentinfo")).toHaveTextContent("125 ms");
    now = 2000;
    await analyzeTemplate(user);
    now = 2290;
    await act(async () => {
      resolvers[1](analysis());
    });
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Success"));
    expect(screen.getByRole("contentinfo")).toHaveTextContent("290 ms");
    now = 3000;
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    now = 3450;
    await act(async () => {
      resolvers[2](success());
    });
    await screen.findByText("Result preview");
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    expect(apiMock.mock.calls.map(([url]) => String(url))).toEqual(["/api/documents/sample", "/api/documents/analyze", "/api/documents/render"]);
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    expect(screen.getByRole("contentinfo")).toHaveTextContent("450 ms");
    expect(screen.getByRole("status").textContent).not.toMatch(/\d{1,2}:\d{2}/);
    expect(screen.queryByRole("tab", { name: "Output" })).not.toBeInTheDocument();
    expect(screen.queryByText("Operation history")).not.toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Workspace status: Success" }));
    expect(screen.getByRole("tabpanel", { name: "Problems" })).toBeInTheDocument();
    await diagnosticsTab(user, "Analysis");
    expect(screen.getByText("Company")).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Hide diagnostics" }));
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    view.unmount();
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    expect(screen.getByRole("status")).toHaveTextContent("Ready");
  });

  it("finishes cancelled dataset generation even when the transport ignores abort and accepts a subsequent analysis", async () => {
    const user = userEvent.setup();
    const resolvers: Array<(response: Response) => void> = [];
    mockPlaygroundFetch(() => new Promise<Response>(resolve => resolvers.push(resolve)));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog");
    await user.click(within(dialog).getByRole("button", { name: "Generate sample" }));
    await user.click(within(dialog).getByRole("button", { name: "Close" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(screen.getByRole("status")).toHaveTextContent("Cancelled");
    await act(async () => {
      resolvers[0](analysis([{ name: "Obsolete", type: "value" }]));
    });
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    await analyzeTemplate(user);
    await act(async () => {
      resolvers[1](analysis([], false));
    });
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Error"));
    await diagnosticsTab(user, "Problems 1");
    expect(screen.getByText("template.invalid_structure", { selector: ".issue-heading strong" })).toBeInTheDocument();
  });

  it("clears Problems while retaining analysis, data and successful result bytes", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch();
    const response = success();
    const clipboard = vi.spyOn(navigator.clipboard, "writeText").mockResolvedValue(undefined);
    const warning = { code: "template.notice", message: "Review this field", part: "/word/document.xml", location: "paragraph 2" };
    apiMock.mockResolvedValueOnce(response).mockResolvedValueOnce(new Response(JSON.stringify({ valid: true, errors: [], warnings: [warning], requirements: [], occurrences: [] })))
      .mockResolvedValueOnce(new Response(JSON.stringify({ data: {}, warnings: [warning] }), { status: 200 }));
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    await analyzeTemplate(user);
    await screen.findByText("Valid", { exact: true });
    await diagnosticsTab(user, "Problems 1");
    const panel = screen.getByRole("tabpanel", { name: "Problems 1" });
    const actions = screen.getByRole("toolbar", { name: "Problems actions" });
    await user.click(within(actions).getByRole("button", { name: "Clear" }));
    expect(within(actions).getByRole("button", { name: "Clear" })).toBeEnabled();
    expect(within(actions).getByRole("button", { name: "Copy" })).toBeEnabled();
    expect(await within(panel).findByText("No problems")).toBeInTheDocument();
    await user.click(within(actions).getByRole("button", { name: "Copy" }));
    await waitFor(() => expect(clipboard).toHaveBeenCalledWith("[]"));
    expect(screen.getByRole("tab", { name: "Problems" })).toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.getByText("Result preview")).toBeInTheDocument();
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
    await documentCommand(user, "Download");
    expect(URL.createObjectURL).toHaveBeenCalledWith(previewSources.at(-1));
    await diagnosticsTab(user, "Analysis");
    expect(screen.getByText("Valid", { exact: true })).toBeInTheDocument();
    expect(screen.getByText("Review this field")).toBeInTheDocument();
    expect(screen.getByRole("status")).toHaveTextContent("Success");
    // A new operation publishes its diagnostics normally, even when it reuses cached analysis.
    await generateJson(user);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).toHaveBeenCalledTimes(3);
    await diagnosticsTab(user, "Problems");
    expect(await within(screen.getByRole("tabpanel", { name: "Problems" })).findByText("No problems")).toBeInTheDocument();
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
    await analyzeTemplate(user);
    await screen.findByText("Invalid", { exact: true });
    await diagnosticsTab(user, "Problems 2");
    const panel = screen.getByRole("tabpanel", { name: "Problems 2" });
    expect(panel).not.toHaveTextContent("Analyze");
    expect(panel).not.toHaveTextContent("Error ·");
    expect(panel).not.toHaveTextContent("Warning ·");
    await user.click(within(screen.getByRole("toolbar", { name: "Problems actions" })).getByRole("button", { name: "Copy" }));
    expect(await screen.findByText("Diagnostics copied to clipboard.")).toBeInTheDocument();
    expect(JSON.parse(clipboard.mock.calls[0][0])).toEqual([
      { title: "template.bad", code: "template.bad", detail: "Bad expression", field: "Amount", dataPath: "$.Rows.Amount", part: "/xl/worksheets/sheet1.xml", location: "C4", formatter: "currency", action: "Analyze", severity: "Error" },
      { title: "template.notice", code: "template.notice", detail: "Review this field", part: "/xl/worksheets/sheet2.xml", location: "D5", action: "Analyze", severity: "Warning" }
    ]);
    clipboard.mockRejectedValueOnce(new DOMException("Denied", "NotAllowedError"));
    await user.click(within(screen.getByRole("toolbar", { name: "Problems actions" })).getByRole("button", { name: "Copy" }));
    expect(await screen.findByText("The diagnostics could not be copied to the clipboard.")).toBeInTheDocument();
    expect(within(panel).getByText("template.bad", { selector: ".issue-heading strong" })).toBeInTheDocument();
    expect(within(panel).getByText("template.notice", { selector: ".issue-heading strong" })).toBeInTheDocument();
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(screen.getByRole("status")).toHaveTextContent("Error");
  });

  it("aborts generation on edits and retains cancelled status after a late response", async () => {
    const user = userEvent.setup();
    let resolve!: (response: Response) => void;
    let signal: AbortSignal | undefined;
    mockPlaygroundFetch(async (_input, init) => {
      signal = init?.signal as AbortSignal;
      return new Promise<Response>(done => { resolve = done; });
    });
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await selectTemplate(user);
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    const status = screen.getByRole("status");
    expect(status).toHaveTextContent("Generating...");
    expect(status).not.toHaveTextContent(/ms|%/);
    fireEvent.change(screen.getByLabelText("Dataset"), { target: { value: sampleData + " " } });
    expect(signal?.aborted).toBe(true);
    expect(status).toHaveTextContent("Cancelled");
    expect(screen.getByRole("contentinfo")).toHaveTextContent(/\d+ ms/);
    await user.click(screen.getByRole("button", { name: "Workspace status: Cancelled" }));
    expect(screen.getByRole("tabpanel", { name: "Problems" })).toBeInTheDocument();
    expect(screen.queryByRole("tab", { name: "Output" })).not.toBeInTheDocument();
    await act(async () => {
      resolve(success());
    });
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Cancelled"));
  });

  it.each(["Cancel", "Escape"])("cancels Generate Dataset through native %s while pending and ignores the obsolete analysis", async action => {
    const user = userEvent.setup();
    let resolve!: (response: Response) => void;
    let signal: AbortSignal | undefined;
    const { apiMock } = mockPlaygroundFetch((_input, init) => {
      signal = init?.signal as AbortSignal;
      return new Promise<Response>(done => { resolve = done; });
    });
    render(<ConfigProvider theme={{ token: { motion: false } }}><App /></ConfigProvider>);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate actions" }));
    await user.click(await screen.findByRole("menuitem", { name: "Generate Dataset..." }));
    const dialog = await screen.findByRole("dialog", { name: "Generate Dataset" });
    await user.click(within(dialog).getByRole("button", { name: "Generate sample" }));
    expect(apiMock).toHaveBeenCalledOnce();
    if (action === "Escape") await user.keyboard("{Escape}");
    else await user.click(within(dialog).getByRole("button", { name: "Cancel" }));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(signal?.aborted).toBe(true);
    expect(screen.getByRole("status")).toHaveTextContent("Cancelled");
    await act(async () => {
      resolve(analysis([{ name: "Obsolete", type: "value" }]));
    });
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Cancelled"));
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    apiMock.mockResolvedValueOnce(sampleJson());
    await generateJson(user);
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    expect(apiMock).toHaveBeenCalledTimes(2);
    expect(JSON.parse((screen.getByLabelText("Dataset") as HTMLTextAreaElement).value)).toEqual({ Company: { Name: "Sample Name 1" } });
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
    fireEvent(editor, typed);
    expect(typed.defaultPrevented).toBe(false);
    await user.type(editor, "?");
    expect(editor).toHaveValue(sampleData + "?");
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    const controls = document.createElement("div");
    controls.innerHTML = '<input /><div contenteditable="true"><span>Editing</span></div><div class="monaco-editor"><button>Editor widget</button></div>';
    document.body.append(controls);
    for (const target of [controls.querySelector("input")!, controls.querySelector("span")!, controls.querySelector("button")!]) {
      const event = new KeyboardEvent("keydown", { key: "?", code: "Slash", shiftKey: true, bubbles: true, cancelable: true });
      fireEvent(target, event);
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
    fireEvent(document.body, after);
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
    await expectTemplateBytes(xlsxSources.at(-1), book.file);
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
    expect(screen.getByRole("status")).toHaveTextContent("Ready");
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
    expect(screen.getByRole("status")).toHaveTextContent("Ready");
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
    await act(async () => {
      resolve(new File(['{"obsolete":true}'], "slow.json"));
    });
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
    expect(screen.getByRole("status")).toHaveTextContent("Ready");
    expect(apiMock).toHaveBeenCalledOnce();
    for (const key of [workspaceStorage.desktop, workspaceStorage.mobile, workspaceStorage.desktopCollapsed, workspaceStorage.mobileCollapsed, workspaceStorage.diagnosticsSize, workspaceStorage.diagnosticsCollapsed, workspaceStorage.wordWrap]) expect(localStorage.getItem(key)).toBeNull();
    expect(localStorage.getItem("unrelated")).toBe("keep");
    expect(db.databases.get("other-app")!.get("items")!.get("key")).toBe("keep");
    expect(data.handle.createWritable).not.toHaveBeenCalled();
    await waitFor(() => expect(localStorage.getItem(workspaceStorage.context)).not.toContain("disk.json"));
    await user.click(screen.getByRole("radio", { name: "Result" }).closest("label")!);
    expect(screen.getByText("Not generated")).toBeInTheDocument();
    expect(screen.queryByRole("tab", { name: "Output" })).not.toBeInTheDocument();
  });

  it("Clear Local Data aborts pending generation and late replies cannot restore results", async () => {
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
    await act(async () => {
      resolve(success());
    });
    await waitFor(() => expect(screen.getByRole("status")).toHaveTextContent("Ready"));
    expect(screen.queryByText("Result preview")).not.toBeInTheDocument();
    expect(screen.queryByRole("tab", { name: "Output" })).not.toBeInTheDocument();
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
    await expectTemplateBytes(xlsxSources.at(-1), blocked.file);
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
    await act(async () => {
      resolve('{"late":true}');
    });
    await waitFor(() => expect(localStorage.getItem(workspaceStorage.context)).toContain(manifest.defaultExample));
    expect(screen.getByLabelText("Dataset")).toHaveValue(sampleData);
    expect(localStorage.getItem(workspaceStorage.context)).not.toContain("slow.json");
    expect(screen.getByRole("status")).toHaveTextContent("Ready");
  });

});
