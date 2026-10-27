import { Button, ConfigProvider, Select, Spin } from "antd";
import { CodeOutlined, DownloadOutlined, FileExcelFilled, FileWordFilled, GithubFilled, UploadOutlined } from "@ant-design/icons";
import { ChangeEvent, DragEvent, useEffect, useRef, useState } from "react";
import { XlsxPreview } from "./components/XlsxPreview";
import { documentMimes, templateFormat, templateAccept } from "./services/documentFormat";
import { DocxPreview } from "./components/DocxPreview";
import { MonacoJsonEditor } from "./components/MonacoJsonEditor";
import { Splitter } from "./components/Splitter";
import { analyzeTemplate, AnalysisFailure } from "./services/analysis";
import { generateDocument, GenerationFailure } from "./services/generation";
import { requirementsToSampleJson } from "./services/sampleJson";
import type { Diagnostic, GeneratedDocument } from "./types";
import type { DocumentFormat } from "./services/documentFormat";
import "./styles.css";
import "./mobile.css";
import "./examples.css";

type Tab = "template" | "result" | "errors";
type Operation = "generate" | "analyze";
type ExampleSelection = string | "custom";

interface ExampleDefinition {
  id: string;
  title: string;
  format: DocumentFormat;
  templatePath: string;
  dataPath: string;
}

interface ExampleManifest {
  defaultExample: string;
  examples: ExampleDefinition[];
}

const repositoryUrl = "https://github.com/montr/monoreport";
function diagnostic(title: string, detail: string): Diagnostic[] {
  return [{ title, detail }];
}

function validateManifest(value: unknown): ExampleManifest {
  if (!value || typeof value !== "object") throw new Error("The examples registry is invalid.");
  const candidate = value as Partial<ExampleManifest>;
  if (typeof candidate.defaultExample !== "string" || !Array.isArray(candidate.examples) || candidate.examples.length === 0) throw new Error("The examples registry is invalid.");
  const ids = new Set<string>();
  for (const example of candidate.examples) {
    if (!example || typeof example.id !== "string" || !example.id || ids.has(example.id) || typeof example.title !== "string" || !example.title || (example.format !== "docx" && example.format !== "xlsx") || typeof example.templatePath !== "string" || !example.templatePath || typeof example.dataPath !== "string" || !example.dataPath) throw new Error("The examples registry is invalid.");
    ids.add(example.id);
  }
  if (!ids.has(candidate.defaultExample)) throw new Error("The default example is missing from the registry.");
  return candidate as ExampleManifest;
}

function filenameFromPath(path: string) {
  const segment = path.split("/").filter(Boolean).at(-1) ?? "template.docx";
  try { return decodeURIComponent(segment); } catch { return segment; }
}

function FormatIcon({ format }: { format: DocumentFormat }) {
  const Icon = format === "docx" ? FileWordFilled : FileExcelFilled;
  return <Icon className={`format-icon format-icon-${format}`} style={{ color: format === "docx" ? "#185abd" : "#217346" }} aria-hidden="true" />;
}

function FormatLabel({ format, title }: { format: DocumentFormat; title: string }) {
  return <span className="format-label"><FormatIcon format={format} /><span className="format-title">{title}</span></span>;
}

function DocumentEmpty({ children }: { children: string }) {
  return <div className="empty"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><path d="M14 2v6h6M8 13h8M8 17h8" /></svg><strong>{children}</strong></div>;
}

export default function App() {
  const [template, setTemplate] = useState<File>();
  const [json, setJson] = useState("{}");
  const [result, setResult] = useState<GeneratedDocument>();
  const [errors, setErrors] = useState<Diagnostic[]>([]);
  const [active, setActive] = useState<Tab>("template");
  const [pending, setPending] = useState<Operation>();
  const [copyLabel, setCopyLabel] = useState("Copy");
  const [dragging, setDragging] = useState(false);
  const [examples, setExamples] = useState<ExampleDefinition[]>([]);
  const [selectedExample, setSelectedExample] = useState<ExampleSelection>();
  const [pendingExample, setPendingExample] = useState<string>();
  const [baselineJson, setBaselineJson] = useState("");
  const [exampleLoading, setExampleLoading] = useState(true);
  const fileInput = useRef<HTMLInputElement>(null);
  const contentRevision = useRef(0);
  const operationId = useRef(0);
  const exampleLoadId = useRef(0);
  const request = useRef<AbortController | undefined>(undefined);
  const currentTemplate = useRef<File | undefined>(undefined);
  const currentJson = useRef(json);
  const copyTimer = useRef<number | undefined>(undefined);

  useEffect(() => {
    let active = true;
    const loadId = ++exampleLoadId.current;
    const expectedRevision = contentRevision.current;
    void (async () => {
      try {
        const response = await fetch("/examples.json", { cache: "no-cache" });
        if (!response.ok) throw new Error("The examples registry could not be loaded.");
        const manifest = validateManifest(await response.json());
        if (!active) return;
        setExamples(manifest.examples);
        if (exampleLoadId.current !== loadId || contentRevision.current !== expectedRevision) return;
        const initial = manifest.examples.find(example => example.id === manifest.defaultExample)!;
        await loadExample(initial, loadId, expectedRevision);
      } catch (error) {
        if (!active) return;
        setExampleLoading(false);
        setErrors(diagnostic("Examples unavailable", error instanceof Error ? error.message : "The examples registry could not be loaded."));
        setActive("errors");
      }
    })();
    return () => {
      active = false;
      operationId.current += 1;
      exampleLoadId.current += 1;
      request.current?.abort();
      if (copyTimer.current) window.clearTimeout(copyTimer.current);
    };
  }, []);

  function invalidateOperation(nextTab: Tab = "template") {
    operationId.current += 1;
    request.current?.abort();
    request.current = undefined;
    setPending(undefined);
    setResult(undefined);
    setActive(current => current === "result" ? nextTab : current);
  }

  function commitContent(nextTab: Tab = "template") {
    contentRevision.current += 1;
    exampleLoadId.current += 1;
    setPendingExample(undefined);
    setExampleLoading(false);
    invalidateOperation(nextTab);
  }

  async function loadExample(example: ExampleDefinition, loadId = ++exampleLoadId.current, expectedRevision = contentRevision.current) {
    setExampleLoading(true);
    try {
      const [templateResponse, dataResponse] = await Promise.all([
        fetch(example.templatePath, { cache: "no-cache" }),
        fetch(example.dataPath, { cache: "no-cache" })
      ]);
      if (!templateResponse.ok || !dataResponse.ok) throw new Error("The selected example assets could not be loaded.");
      const [blob, data] = await Promise.all([templateResponse.blob(), dataResponse.text()]);
      JSON.parse(data);
      if (exampleLoadId.current !== loadId || contentRevision.current !== expectedRevision) return;
      const file = new File([blob], filenameFromPath(example.templatePath), { type: documentMimes[example.format] });
      contentRevision.current += 1;
      invalidateOperation("template");
      currentTemplate.current = file;
      currentJson.current = data;
      setTemplate(file);
      setJson(data);
      setBaselineJson(data);
      setSelectedExample(example.id);
      setPendingExample(undefined);
      setErrors([]);
      setActive("template");
    } catch (error) {
      if (exampleLoadId.current !== loadId || contentRevision.current !== expectedRevision) return;
      setErrors(diagnostic("Example could not be loaded", error instanceof SyntaxError ? "The example JSON data is malformed." : error instanceof Error ? error.message : "The selected example could not be loaded."));
      setPendingExample(undefined);
      setActive("errors");
    } finally {
      if (exampleLoadId.current === loadId) setExampleLoading(false);
    }
  }

  function chooseExample(id: string) {
    const example = examples.find(item => item.id === id);
    if (!example) return;
    if (id === pendingExample) return;
    if (id === selectedExample) {
      if (exampleLoading) {
        exampleLoadId.current += 1;
        setPendingExample(undefined);
        setExampleLoading(false);
      }
      return;
    }
    if ((selectedExample === "custom" || json !== baselineJson) && !window.confirm("Replace the current template and JSON with the selected example?")) return;
    setPendingExample(id);
    void loadExample(example);
  }

  function acceptTemplate(file: File) {
    if (!templateFormat(file.name)) {
      setErrors(diagnostic("Unsupported template format", "Select a .docx or .xlsx file."));
      setActive("errors");
      return;
    }
    commitContent("template");
    currentTemplate.current = file;
    setTemplate(file);
    setSelectedExample("custom");
    setErrors([]);
    setActive("template");
  }

  function selected(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) acceptTemplate(file);
    event.target.value = "";
  }

  function drop(event: DragEvent) {
    event.preventDefault();
    setDragging(false);
    const file = event.dataTransfer.files?.[0];
    if (file) acceptTemplate(file);
  }

  function changeJson(value: string) {
    commitContent("template");
    currentJson.current = value;
    setJson(value);
  }

  async function copyJson() {
    try {
      await navigator.clipboard.writeText(json);
      setCopyLabel("Copied");
      if (copyTimer.current) window.clearTimeout(copyTimer.current);
      copyTimer.current = window.setTimeout(() => setCopyLabel("Copy"), 1200);
    } catch {
      setErrors(diagnostic("Clipboard unavailable", "The JSON could not be copied to the clipboard."));
      setActive("errors");
    }
  }

  function formatJson() {
    try {
      changeJson(JSON.stringify(JSON.parse(json), null, 2));
      setErrors([]);
    } catch {
      setErrors(diagnostic("Invalid JSON", "Fix the JSON syntax before formatting."));
      setActive("errors");
    }
  }

  async function generate() {
    if (pending) return;
    if (!template) {
      setErrors(diagnostic("Template is required", "Select a DOCX or XLSX file before generating a document."));
      setActive("errors");
      return;
    }
    try {
      JSON.parse(json);
    } catch {
      invalidateOperation("errors");
      setErrors(diagnostic("Invalid JSON", "Fix the JSON syntax before generating a document."));
      setActive("errors");
      return;
    }

    invalidateOperation("template");
    const operation = operationId.current;
    const expectedRevision = contentRevision.current;
    const controller = new AbortController();
    request.current = controller;
    setErrors([]);
    setPending("generate");
    try {
      const generated = await generateDocument(template, json, controller.signal);
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return;
      setResult(generated);
      setActive("result");
    } catch (error) {
      if (operationId.current !== operation || contentRevision.current !== expectedRevision || (error instanceof DOMException && error.name === "AbortError")) return;
      setErrors(error instanceof GenerationFailure ? error.diagnostics : diagnostic("Generation failed", "The document could not be generated."));
      setActive("errors");
    } finally {
      if (operationId.current === operation) {
        request.current = undefined;
        setPending(undefined);
      }
    }
  }

  async function generateJson() {
    if (pending || !template) return;
    const operation = ++operationId.current;
    request.current?.abort();
    const expectedRevision = contentRevision.current;
    const source = template;
    const controller = new AbortController();
    request.current = controller;
    setErrors([]);
    setPending("analyze");
    try {
      const analysis = await analyzeTemplate(source, controller.signal);
      if (operationId.current !== operation || contentRevision.current !== expectedRevision || currentTemplate.current !== source) return;
      const proposed = requirementsToSampleJson(analysis.requirements);
      const previous = currentJson.current;
      if (proposed === previous) return;
      if (previous.length > 0 && !window.confirm("Replace current JSON with generated example?")) return;
      if (operationId.current !== operation || contentRevision.current !== expectedRevision || currentTemplate.current !== source || currentJson.current !== previous) return;
      changeJson(proposed);
      setErrors([]);
    } catch (error) {
      if (operationId.current !== operation || contentRevision.current !== expectedRevision || (error instanceof DOMException && error.name === "AbortError")) return;
      setErrors(error instanceof AnalysisFailure ? error.diagnostics : diagnostic("Analysis failed", "The template could not be analyzed."));
      setActive("errors");
    } finally {
      if (operationId.current === operation) {
        request.current = undefined;
        setPending(undefined);
      }
    }
  }

  function clearTemplate() {
    commitContent("template");
    currentTemplate.current = undefined;
    setTemplate(undefined);
    setSelectedExample("custom");
    setErrors([]);
    setActive("template");
  }

  function downloadBlob(blob: Blob, filename: string) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  const modified = selectedExample !== undefined && selectedExample !== "custom" && json !== baselineJson;
  const selectorValue = pendingExample ?? selectedExample;
  const exampleFormat = selectorValue === "custom"
    ? template && templateFormat(template.name)
    : examples.find(example => example.id === selectorValue)?.format;
  const exampleOptions = examples.map(example => ({
    value: example.id,
    label: <FormatLabel format={example.format} title={example.title} />
  }));
  if (selectorValue === "custom" && exampleFormat) {
    exampleOptions.push({ value: "custom", label: <FormatLabel format={exampleFormat} title="Custom" />, disabled: true } as typeof exampleOptions[number]);
  }

  const content = active === "errors" ? (
    errors.length ? <div className="issues"><div className="issues-title"><strong>{errors.length} {errors.length === 1 ? "error" : "errors"}</strong></div>{errors.map((error, index) => <div className="issue" key={`${error.code ?? error.title}-${index}`}><span>{index + 1}.</span><div><strong>{error.title}</strong><p>{[error.code, error.field, error.dataPath, error.part, error.location, error.formatter, error.detail].filter(Boolean).join(" · ")}</p></div></div>)}</div> : <DocumentEmpty>No errors</DocumentEmpty>
  ) : active === "result" && result ? (
    result.format === "xlsx" ? <XlsxPreview source={result.blob} label="Result" /> : <DocxPreview source={result.blob} label="Result" />
  ) : template ? (
    templateFormat(template.name) === "xlsx" ? <XlsxPreview source={template} label="Template" /> : <DocxPreview source={template} label="Template" />
  ) : <DocumentEmpty>Select a template</DocumentEmpty>;

  return <ConfigProvider theme={{ token: { colorPrimary: "#cf4d48", borderRadius: 6, fontSize: 13 } }}>
    <div className="app">
      <header><span className="brand"><b>mono</b>report</span><span className="subbrand">Playground</span></header>
      <main className="layout">
        <section className="source" aria-label="Template and input data">
          <div className="filebar">
            <span className="label">Template</span>
            <span className="filename">{template?.name ?? "No file selected"}</span>
            {template && <Button type="text" className="clear" aria-label="Remove template" onClick={clearTemplate}>×</Button>}
            <Button type="primary" className="generate" aria-label="Generate document" loading={pending === "generate"} disabled={Boolean(pending)} onClick={() => void generate()}>Generate <span aria-hidden="true">→</span></Button>
          </div>
          <div className="editbar">
            <div className="mode" aria-label="Input mode"><Button disabled>Form</Button><Button className="active" aria-current="true">JSON</Button></div>
            <div className="edit-actions"><Button size="small" aria-label="Generate JSON" title="Creates a starting point; edit it because XLSX row counts and other data-dependent requirements are not validated." loading={pending === "analyze"} disabled={!template || Boolean(pending)} onClick={() => void generateJson()}>Generate JSON</Button><Button size="small" onClick={() => void copyJson()}>{copyLabel}</Button><Button size="small" onClick={formatJson}>Format</Button></div>
          </div>
          <MonacoJsonEditor value={json} onChange={changeJson} />
        </section>
        <Splitter />
        <section className="preview" aria-label="Preview">
          <div className="previewbar">
            <nav className="tabs" aria-label="Preview tabs">
              <button className={active === "template" ? "active" : ""} onClick={() => setActive("template")}>Template</button>
              <button className={active === "result" ? "active" : ""} disabled={!result} onClick={() => result && setActive("result")}>Result</button>
              <button className={active === "errors" ? "active" : ""} onClick={() => setActive("errors")}>Errors {errors.length > 0 && <span className="count">{errors.length}</span>}</button>
            </nav>
            <span className="spacer" />
            {(pending || exampleLoading) && <Spin size="small" aria-label={exampleLoading ? "Loading example" : pending === "analyze" ? "Analyzing template" : "Generating document"} />}
          </div>
          {active === "template" && <div className="document-toolbar" aria-label="Template actions">
            <div className="examples-control">
              <label htmlFor="examples">Examples</label>
              <div className="examples-picker">
                <Select
                  id="examples"
                  className="examples-select"
                  size="small"
                  aria-label="Examples"
                  value={selectorValue ?? undefined}
                  placeholder="Loading…"
                  disabled={examples.length === 0}
                  options={exampleOptions}
                  popupMatchSelectWidth={false}
                  onChange={chooseExample}
                />
              </div>
              {modified && <span className="modified" aria-label="Example modified">Modified</span>}
            </div>
            <div className="document-actions">
              <Button size="small" icon={<UploadOutlined aria-hidden="true" />} onClick={() => fileInput.current?.click()}>Upload</Button>
              <Button size="small" aria-label="Download template" disabled={!template} onClick={() => template && downloadBlob(template, template.name)} icon={<DownloadOutlined aria-hidden="true" />}>Download</Button>
              <input ref={fileInput} type="file" accept={templateAccept} onChange={selected} />
            </div>
          </div>}
          {active === "result" && <div className="document-toolbar result-toolbar" aria-label="Result actions">
            <Button size="small" aria-label="Download result" disabled={!result} onClick={() => result && downloadBlob(result.blob, result.filename)} icon={<DownloadOutlined aria-hidden="true" />}>Download</Button>
          </div>}
          <div
            className={`canvas ${(active === "template" && template && templateFormat(template.name) === "xlsx") || (active === "result" && result?.format === "xlsx") ? "xlsx-view" : ""} ${active === "errors" && errors.length > 0 ? "errors-view" : ""} ${dragging ? "drop-target" : ""}`}
            onDragEnter={event => { if (event.dataTransfer.types.includes("Files")) { event.preventDefault(); setDragging(true); } }}
            onDragOver={event => { if (event.dataTransfer.types.includes("Files")) event.preventDefault(); }}
            onDragLeave={event => { if (!event.currentTarget.contains(event.relatedTarget as Node)) setDragging(false); }}
            onDrop={drop}
          >{content}</div>
        </section>
      </main>
      <footer>
        <span className="project-links"><a href={repositoryUrl} target="_blank" rel="noreferrer" title="MonoReport on GitHub"><GithubFilled aria-hidden="true" />MonoReport</a>{__APP_VERSION__ === "dev" ? <span>dev</span> : <a href={`${repositoryUrl}/releases/tag/v${__APP_VERSION__}`} target="_blank" rel="noreferrer">v{__APP_VERSION__}</a>}</span>
        <a href="/docs/" title="API documentation"><CodeOutlined aria-hidden="true" />API docs</a>
      </footer>
    </div>
  </ConfigProvider>;
}
