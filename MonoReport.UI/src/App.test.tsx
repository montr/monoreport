import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import examplesRegistry from "../public/examples.json";

vi.mock("./components/MonacoJsonEditor", () => ({
  MonacoJsonEditor: ({ value, onChange }: { value: string; onChange(value: string): void }) => <textarea aria-label="JSON data" value={value} onChange={event => onChange(event.target.value)} />
}));

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
  await user.upload(screen.getByLabelText("Template actions").querySelector("input[type=file]") as HTMLInputElement, file);
}

describe("MonoReport playground", () => {
  beforeEach(() => { previewSources.length = 0; xlsxSources.length = 0; previewTestState.failResultPreview = false; mockPlaygroundFetch(); });

  it("loads the default example without generation and keeps the approved controls and accessible splitter", async () => {
    render(<App />);
    expect(await screen.findByText("Template preview")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Form" })).toBeDisabled();
    expect(screen.getByLabelText("JSON data")).toHaveValue(sampleData);
    expect(screen.getByRole("combobox", { name: "Examples" })).toHaveValue("construction-cost-estimate");
    expect(screen.getByRole("button", { name: "Download template" })).toBeEnabled();
    expect(screen.getByRole("button", { name: "Generate JSON" })).toBeEnabled();
    const repositoryLink = screen.getByRole("link", { name: "MonoReport" });
    expect(repositoryLink).toHaveAttribute("href", "https://github.com/montr/monoreport");
    expect(repositoryLink.querySelector("svg")).not.toBeNull();
    expect(screen.getByText("dev")).toBeInTheDocument();
    const splitter = screen.getByRole("separator");
    fireEvent.keyDown(splitter, { key: "ArrowRight" });
    expect(splitter).toHaveAttribute("aria-valuenow", "48");
  });

  it("loads an XLSX registry example with format-correct file metadata and preview dispatch", async () => {
    const xlsxExample = examplesRegistry.examples.find(example => example.id === "xlsx-tender")!;
    expect(xlsxExample.format).toBe("xlsx");
    const registry = { defaultExample: xlsxExample.id, examples: [xlsxExample] };
    const fetchMock = vi.fn((input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/examples.json") return Promise.resolve(new Response(JSON.stringify(registry), { status: 200 }));
      if (url === xlsxExample.templatePath) return Promise.resolve(new Response(new Uint8Array([80, 75, 3, 4]), { status: 200, headers: { "Content-Type": "application/octet-stream" } }));
      if (url === xlsxExample.dataPath) return Promise.resolve(new Response('{"Title":"Office equipment"}', { status: 200 }));
      return Promise.reject(new Error(`Unexpected request: ${url}`));
    });
    vi.stubGlobal("fetch", fetchMock);

    render(<App />);

    expect(await screen.findByText("Template XLSX preview")).toBeInTheDocument();
    expect(screen.getByRole("combobox", { name: "Examples" })).toHaveValue("xlsx-tender");
    expect(screen.getByLabelText("JSON data")).toHaveValue('{"Title":"Office equipment"}');
    expect(xlsxSources.at(-1)).toBeInstanceOf(File);
    expect(xlsxSources.at(-1)).toMatchObject({ name: "xlsx-tender.template.xlsx", type: xlsxMime });
    expect(fetchMock).toHaveBeenCalledWith(xlsxExample.templatePath);
    expect(fetchMock).toHaveBeenCalledWith(xlsxExample.dataPath);
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
    render(<App />);
    await screen.findByText("Template preview");
    fireEvent.change(screen.getByRole("combobox", { name: "Examples" }), { target: { value: "xlsx-tender" } });
    await waitFor(() => expect(resolveXlsxTemplate).toBeTypeOf("function"));

    const custom = new File([new Uint8Array([9])], "custom.xlsx", { type: xlsxMime });
    await selectTemplate(user, custom);
    resolveXlsxTemplate(new Response(new Uint8Array([80, 75, 3, 4]), { status: 200, headers: { "Content-Type": xlsxMime } }));
    resolveXlsxData(new Response('{"stale":true}', { status: 200 }));

    await waitFor(() => expect(screen.getByText("custom.xlsx")).toBeInTheDocument());
    expect(screen.getByRole("combobox", { name: "Examples" })).toHaveValue("custom");
    expect(screen.getByLabelText("JSON data")).toHaveValue(sampleData);
    expect(xlsxSources.at(-1)).toBe(custom);
  });

  it("formats and copies valid JSON while invalid JSON remains unchanged and never requests", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch();
    render(<App />);
    const editor = screen.getByLabelText("JSON data");
    const clipboard = vi.spyOn(navigator.clipboard, "writeText");
    fireEvent.change(editor, { target: { value: '{"a":1}' } });
    await user.click(screen.getByRole("button", { name: "Format" }));
    expect(editor).toHaveValue('{\n  "a": 1\n}');
    await user.click(screen.getByRole("button", { name: "Copy" }));
    expect(clipboard).toHaveBeenCalledWith('{\n  "a": 1\n}');
    fireEvent.change(editor, { target: { value: "{" } });
    await user.click(screen.getByRole("button", { name: "Format" }));
    expect(editor).toHaveValue("{");
    expect(screen.getByText("Invalid JSON")).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(apiMock).not.toHaveBeenCalled();
  });

  it("uses picker and drop through one replacement path, preserves JSON, and rejects unsupported files without replacement", async () => {
    const user = userEvent.setup();
    mockPlaygroundFetch(async () => success());
    render(<App />);
    const editor = screen.getByLabelText("JSON data");
    fireEvent.change(editor, { target: { value: '{"kept":true}' } });
    await selectTemplate(user);
    expect(screen.getByText("sample.template.docx")).toBeInTheDocument();
    const replacement = new File([new Uint8Array([1])], "second.docx", { type: mime });
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [replacement], types: ["Files"] } });
    expect(screen.getByText("second.docx")).toBeInTheDocument();
    expect(editor).toHaveValue('{"kept":true}');
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    expect(screen.getByRole("button", { name: /Download/ })).toBeEnabled();
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [new File([new Uint8Array([1])], "bad.pdf")], types: ["Files"] } });
    expect(screen.getByText("Unsupported template format")).toBeInTheDocument();
    expect(screen.getByText("second.docx")).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: /Download/ })).not.toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Result" }));
    expect(screen.getByRole("button", { name: "Download result" })).toBeEnabled();
    const third = new File([new Uint8Array([2])], "third.docx", { type: mime });
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [third], types: ["Files"] } });
    expect(screen.getByText("third.docx")).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Download result" })).not.toBeInTheDocument();
    expect(editor).toHaveValue('{"kept":true}');
  });

  it("selects XLSX by extension for picker/drop, preserves JSON, invalidates results, and switches back to DOCX", async () => {
    const user = userEvent.setup();
    const { apiMock } = mockPlaygroundFetch(async () => success());
    render(<App />);
    await screen.findByText("Template preview");
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    await user.click(screen.getByRole("button", { name: "Template" }));
    const picker = screen.getByLabelText("Template actions").querySelector("input")!;
    expect(picker).toHaveAttribute("accept", expect.stringContaining(".xlsx"));
    const book = new File([new Uint8Array([1])], "tender.XLSX", { type: mime });
    await selectTemplate(user, book);
    expect(screen.getByText("Template XLSX preview")).toBeInTheDocument();
    expect(xlsxSources.at(-1)).toBe(book);
    expect(previewSources.at(-1)).not.toBe(book);
    expect(screen.getByRole("button", { name: "Result" })).toBeDisabled();
    expect(screen.getByLabelText("JSON data")).toHaveValue(sampleData);
    const dropped = new File([new Uint8Array([2])], "second.xlsx");
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [dropped], types: ["Files"] } });
    expect(xlsxSources.at(-1)).toBe(dropped);
    expect(screen.getByRole("button", { name: "Download template" })).toBeEnabled();
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [new File([], "bad.xls")], types: ["Files"] } });
    expect(screen.getByText("Select a .docx or .xlsx file.")).toBeInTheDocument();
    expect(screen.getByText("second.xlsx")).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Template" }));
    await selectTemplate(user);
    expect(screen.getByText("Template preview")).toBeInTheDocument();
    expect(screen.queryByText("Template XLSX preview")).not.toBeInTheDocument();
    expect(screen.getByLabelText("JSON data")).toHaveValue(sampleData);
    expect(apiMock).toHaveBeenCalledOnce();
  });

  it("sends generation only on click, prevents duplicates, previews real response bytes, and downloads them", async () => {
    const user = userEvent.setup();
    let resolve!: (response: Response) => void;
    const responseBlob = new Blob([new Uint8Array([80, 75, 7, 8])], { type: mime });
    const { apiMock } = mockPlaygroundFetch(() => new Promise<Response>(done => { resolve = done; }));
    render(<App />);
    await selectTemplate(user);
    expect(apiMock).not.toHaveBeenCalled();
    const generate = screen.getByRole("button", { name: "Generate document" });
    await user.click(generate);
    expect(generate).toBeDisabled();
    await user.click(generate);
    expect(apiMock).toHaveBeenCalledTimes(1);
    resolve({ ok: true, status: 200, blob: async () => responseBlob, headers: new Headers({ "Content-Type": mime, "Content-Disposition": "attachment; filename=generated.docx" }) } as Response);
    await screen.findByText("Result preview");
    expect(previewSources.at(-1)).toBe(responseBlob);
    const download = screen.getByRole("button", { name: /Download/ });
    expect(download).toBeEnabled();
    const click = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
    await user.click(download);
    expect(URL.createObjectURL).toHaveBeenCalledWith(responseBlob);
    expect(click).toHaveBeenCalled();
  });

  it("manually generates XLSX without analysis, previews the response format, and keeps the original download after preview failure", async () => {
    const user = userEvent.setup();
    previewTestState.failResultPreview = true;
    const resultBlob = new Blob([new Uint8Array([80, 75, 7, 8, 9])], { type: xlsxMime });
    const { apiMock } = mockPlaygroundFetch(async () => ({ ok: true, status: 200, blob: async () => resultBlob, headers: new Headers({ "Content-Type": `${xlsxMime}; charset=binary`, "Content-Disposition": "attachment; filename=result.xlsx" }) } as Response));
    render(<App />);
    const sourceFile = new File([new Uint8Array([1, 2])], "manual.xlsx", { type: xlsxMime });
    await selectTemplate(user, sourceFile);
    fireEvent.change(screen.getByLabelText("JSON data"), { target: { value: '{"manual":true}' } });
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(await screen.findByText("Result XLSX preview failed")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Download result" })).toBeEnabled();
    expect(xlsxSources.at(-1)).toBe(resultBlob);
    expect(previewSources).not.toContain(resultBlob);
    expect(apiMock).toHaveBeenCalledOnce();
    expect(String((apiMock.mock.calls[0] as [string])[0])).toBe("/api/documents/generate");
    const requestBody = (apiMock.mock.calls[0] as [string, RequestInit])[1].body as FormData;
    expect(requestBody.get("template")).toBe(sourceFile);
    expect(requestBody.get("data")).toBe('{"manual":true}');

    let downloadedFilename = "";
    const click = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) { downloadedFilename = this.download; });
    await user.click(screen.getByRole("button", { name: "Download result" }));
    expect(URL.createObjectURL).toHaveBeenCalledWith(resultBlob);
    expect(click).toHaveBeenCalled();
    expect(downloadedFilename).toBe("result.xlsx");
  });

  it("renders real API diagnostics and multiple issues without mock content", async () => {
    const user = userEvent.setup();
    mockPlaygroundFetch(async () => new Response(JSON.stringify({ errors: [{ code: "first", message: "First issue", dataPath: "Items[3].Amount", location: "/word/document.xml, paragraph 2" }, { code: "second", detail: "Second issue" }] }), { status: 422, headers: { "Content-Type": "application/problem+json" } }));
    render(<App />);
    await selectTemplate(user);
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(await screen.findByText("2 errors")).toBeInTheDocument();
    expect(screen.getByText(/First issue/)).toBeInTheDocument();
    expect(screen.getByText(/Second issue/)).toBeInTheDocument();
    expect(screen.getByText(/Items\[3\]\.Amount/)).toBeInTheDocument();
    expect(screen.queryByText("Unmatched collection region")).not.toBeInTheDocument();
  });

  it("invalidates a successful result on JSON or template change and after failure", async () => {
    const user = userEvent.setup();
    const responses = [success(), new Response(JSON.stringify({ title: "Failed", detail: "Incompatible", code: "generation.invalid_data" }), { status: 422, headers: { "Content-Type": "application/problem+json" } })];
    mockPlaygroundFetch(async () => responses.shift()!);
    render(<App />);
    await selectTemplate(user);
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");
    expect(screen.getByRole("button", { name: /Download/ })).toBeEnabled();
    await user.type(screen.getByLabelText("JSON data"), " ");
    expect(screen.queryByRole("button", { name: "Download result" })).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Result" })).toBeDisabled();
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    expect(await screen.findByText(/generation.invalid_data/)).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Download result" })).not.toBeInTheDocument();
  });

  it("ignores an obsolete success after JSON changes during generation", async () => {
    const user = userEvent.setup();
    let resolve!: (response: Response) => void;
    mockPlaygroundFetch(() => new Promise<Response>(done => { resolve = done; }));
    render(<App />);
    await selectTemplate(user);
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await user.type(screen.getByLabelText("JSON data"), " ");
    resolve(success());
    await waitFor(() => expect(screen.getByRole("button", { name: "Result" })).toBeDisabled());
    expect(screen.queryByRole("button", { name: "Download result" })).not.toBeInTheDocument();
  });

  it("analyzes only on explicit action, confirms replacement, and invalidates an existing Result", async () => {
    const user = userEvent.setup();
    const confirm = vi.spyOn(window, "confirm").mockReturnValue(true);
    const responses = [success(), analysis([
      { name: "Company", type: "object", fields: [{ name: "Name", type: "value" }] },
      { name: "Orders", type: "array", fields: [{ name: "Number", type: "value" }] }
    ])];
    const { apiMock } = mockPlaygroundFetch(async () => responses.shift()!);
    render(<App />);
    await selectTemplate(user);
    expect(apiMock).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    await screen.findByText("Result preview");

    await user.click(screen.getByRole("button", { name: "Generate JSON" }));

    await waitFor(() => expect(screen.getByLabelText("JSON data")).toHaveValue(JSON.stringify({ Company: { Name: "" }, Orders: [{ Number: "" }] }, null, 2)));
    expect(confirm).toHaveBeenCalledWith("Replace current JSON with generated example?");
    expect(screen.getByRole("button", { name: "Result" })).toBeDisabled();
    const [, analysisInit] = apiMock.mock.calls[1] as [string, RequestInit];
    expect(((analysisInit.body as FormData).get("template") as File).name).toBe("sample.template.docx");
  });

  it("does not confirm for empty JSON and leaves nonempty JSON unchanged when replacement is cancelled", async () => {
    const user = userEvent.setup();
    const confirm = vi.spyOn(window, "confirm").mockReturnValue(false);
    mockPlaygroundFetch(async () => analysis());
    render(<App />);
    await selectTemplate(user);
    const editor = screen.getByLabelText("JSON data");
    fireEvent.change(editor, { target: { value: "" } });
    await user.click(screen.getByRole("button", { name: "Generate JSON" }));
    await waitFor(() => expect(editor).toHaveValue(JSON.stringify({ Company: { Name: "" } }, null, 2)));
    expect(confirm).not.toHaveBeenCalled();

    fireEvent.change(editor, { target: { value: "keep me" } });
    await user.click(screen.getByRole("button", { name: "Generate JSON" }));
    await waitFor(() => expect(confirm).toHaveBeenCalledTimes(1));
    expect(editor).toHaveValue("keep me");
  });

  it("never applies stale analysis after JSON or template changes", async () => {
    const user = userEvent.setup();
    const resolvers: Array<(response: Response) => void> = [];
    mockPlaygroundFetch(() => new Promise<Response>(resolve => resolvers.push(resolve)));
    render(<App />);
    await selectTemplate(user);
    const editor = screen.getByLabelText("JSON data");
    await user.click(screen.getByRole("button", { name: "Generate JSON" }));
    fireEvent.change(editor, { target: { value: "edited during analysis" } });
    resolvers[0](analysis());
    await waitFor(() => expect(editor).toHaveValue("edited during analysis"));

    await user.click(screen.getByRole("button", { name: "Generate JSON" }));
    const replacement = new File([new Uint8Array([1])], "replacement.docx", { type: mime });
    fireEvent.drop(screen.getByLabelText("Preview").querySelector(".canvas")!, { dataTransfer: { files: [replacement], types: ["Files"] } });
    resolvers[1](analysis([{ name: "Stale", type: "value" }]));
    await waitFor(() => expect(screen.getByText("replacement.docx")).toBeInTheDocument());
    expect(editor).toHaveValue("edited during analysis");
  });

  it("shows analysis failures without replacing JSON", async () => {
    const user = userEvent.setup();
    mockPlaygroundFetch(async () => new Response(JSON.stringify({ valid: false,
      errors: [{ code: "template.unmatched_brace", field: "{Name", location: "/word/document.xml, paragraph 40",
        message: "Template braces are not balanced or escaped correctly. Token: «{Name». Nearby source text: «Invoice reference: {Name»." }],
      warnings: [], requirements: [], occurrences: [] }), { status: 200 }));
    render(<App />);
    await selectTemplate(user);
    await user.click(screen.getByRole("button", { name: "Generate JSON" }));
    expect(await screen.findByText("template.unmatched_brace", { selector: "strong" })).toBeInTheDocument();
    expect(screen.getByText(/Invoice reference: \{Name/, { selector: ".issue p" })).toHaveTextContent("/word/document.xml, paragraph 40");
    expect(screen.getByLabelText("JSON data")).toHaveValue(sampleData);
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
    render(<App />);
    await waitFor(() => expect(resolveTemplate).toBeTypeOf("function"));
    fireEvent.change(screen.getByLabelText("JSON data"), { target: { value: "user edit" } });
    resolveTemplate(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
    resolveData(new Response('{"stale":true}', { status: 200 }));
    await waitFor(() => expect(screen.getByLabelText("JSON data")).toHaveValue("user edit"));
    expect(screen.getByText("No file selected")).toBeInTheDocument();
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
    render(<App />);
    await waitFor(() => expect(resolveTemplate).toBeTypeOf("function"));
    await selectTemplate(user, new File([new Uint8Array([9])], "custom.docx", { type: mime }));
    resolveTemplate(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
    resolveData(new Response('{"stale":true}', { status: 200 }));
    await waitFor(() => expect(screen.getByText("custom.docx")).toBeInTheDocument());
    expect(screen.getByRole("combobox", { name: "Examples" })).toHaveValue("custom");
    expect(screen.getByLabelText("JSON data")).toHaveValue("{}");
  });

  it("marks edits, preserves Custom JSON, confirms restoration, and downloads exact template bytes", async () => {
    const user = userEvent.setup();
    const confirm = vi.spyOn(window, "confirm");
    render(<App />);
    await screen.findByText("Template preview");
    fireEvent.change(screen.getByLabelText("JSON data"), { target: { value: "byte-for-byte JSON" } });
    expect(screen.getByLabelText("Example modified")).toBeInTheDocument();
    const custom = new File([new Uint8Array([4, 5, 6])], "custom.docx", { type: mime });
    await selectTemplate(user, custom);
    expect(screen.getByLabelText("JSON data")).toHaveValue("byte-for-byte JSON");
    expect(screen.getByRole("combobox", { name: "Examples" })).toHaveValue("custom");
    const click = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
    await user.click(screen.getByRole("button", { name: "Download template" }));
    expect(URL.createObjectURL).toHaveBeenCalledWith(custom);
    expect(click).toHaveBeenCalled();
    confirm.mockReturnValueOnce(false);
    fireEvent.change(screen.getByRole("combobox", { name: "Examples" }), { target: { value: "construction-cost-estimate" } });
    expect(screen.getByRole("combobox", { name: "Examples" })).toHaveValue("custom");
    expect(screen.getByLabelText("JSON data")).toHaveValue("byte-for-byte JSON");
    confirm.mockReturnValueOnce(true);
    fireEvent.change(screen.getByRole("combobox", { name: "Examples" }), { target: { value: "construction-cost-estimate" } });
    await waitFor(() => expect(screen.getByLabelText("JSON data")).toHaveValue(sampleData));
  });

  it("prevents a Generate response from applying after an example commits", async () => {
    let templateLoads = 0;
    let dataLoads = 0;
    let resolveTemplate!: (response: Response) => void;
    let resolveData!: (response: Response) => void;
    let resolveGenerate!: (response: Response) => void;
    vi.spyOn(window, "confirm").mockReturnValue(true);
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
    render(<App />);
    await screen.findByText("Template preview");
    await selectTemplate(user, new File([new Uint8Array([9])], "custom.docx", { type: mime }));
    fireEvent.change(screen.getByRole("combobox", { name: "Examples" }), { target: { value: "construction-cost-estimate" } });
    await waitFor(() => expect(resolveTemplate).toBeTypeOf("function"));
    await user.click(screen.getByRole("button", { name: "Generate document" }));
    resolveTemplate(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
    resolveData(new Response(sampleData, { status: 200 }));
    await waitFor(() => expect(screen.getByText("sample.template.docx")).toBeInTheDocument());
    resolveGenerate(success());
    await waitFor(() => expect(screen.getByRole("button", { name: "Result" })).toBeDisabled());
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
    render(<App />);
    await screen.findByText("Template preview");
    const selector = screen.getByRole("combobox", { name: "Examples" });
    fireEvent.change(selector, { target: { value: "second" } });
    await waitFor(() => expect(resolveTemplate).toBeTypeOf("function"));
    fireEvent.change(selector, { target: { value: "construction-cost-estimate" } });
    resolveTemplate(new Response(new Uint8Array([1, 2]), { status: 200, headers: { "Content-Type": mime } }));
    resolveData(new Response('{"second":true}', { status: 200 }));
    await waitFor(() => expect(selector).toHaveValue("construction-cost-estimate"));
    expect(screen.getByLabelText("JSON data")).toHaveValue(sampleData);
    expect(screen.getByText("sample.template.docx")).toBeInTheDocument();
  });

  it("shows an active malformed example failure without partially replacing content", async () => {
    vi.stubGlobal("fetch", vi.fn((input: RequestInfo | URL) => {
      const url = String(input);
      if (url === "/examples.json") return Promise.resolve(new Response(JSON.stringify(manifest), { status: 200 }));
      if (url.endsWith(".docx")) return Promise.resolve(new Response(new Uint8Array([80, 75]), { status: 200, headers: { "Content-Type": mime } }));
      return Promise.resolve(new Response("{", { status: 200 }));
    }));
    render(<App />);
    expect(await screen.findByText("Example could not be loaded")).toBeInTheDocument();
    expect(screen.getByLabelText("JSON data")).toHaveValue("{}");
    expect(screen.getByText("No file selected")).toBeInTheDocument();
  });
});
