import { Alert, Badge, Button, ConfigProvider, Dropdown, Empty, Flex, Form, InputNumber, message, Modal, Radio, Space, Spin, Tabs, Tag, theme, Tooltip, Typography, Upload } from "antd";
import { BulbOutlined, CheckCircleFilled, CheckOutlined, CloseCircleFilled, CloseOutlined, CopyOutlined, DeleteOutlined, DownOutlined, EllipsisOutlined, ExclamationCircleFilled, ExclamationCircleOutlined, FileExcelFilled, FileWordFilled, GithubFilled, StopOutlined, UploadOutlined } from "@ant-design/icons";
import { ChangeEvent, DragEvent, useEffect, useRef, useState } from "react";
import { XlsxPreview } from "./components/XlsxPreview";
import { documentMimes, templateFormat, templateAccept } from "./services/documentFormat";
import { DocxPreview } from "./components/DocxPreview";
import { MonacoJsonEditor, type JsonEditorCommands } from "./components/MonacoJsonEditor";
import { WorkspaceLayout } from "./components/WorkspaceLayout";
import { analyzeTemplate, AnalysisFailure } from "./services/analysis";
import { generateDocument, GenerationFailure } from "./services/generation";
import { SampleJson, type SampleJsonSettings } from "./services/sampleJson";
import { AnalysisView } from "./components/AnalysisView";
import type { Diagnostic, GeneratedDocument, TemplateAnalysis } from "./types";
import { LocalFiles, type LocalFileHandle, type FileAccess } from "./services/localFiles";
import { WorkspacePersistence, workspaceStorage, type WorkspaceContext, type DataReference, type TemplateReference } from "./services/workspacePersistence";
import type { DocumentFormat } from "./services/documentFormat";
import "./styles.css";

type Tab = "template" | "result";
type Operation = "generate" | "analyze" | "generate-json";
type DiagnosticsTab = "problems" | "analysis";
type OperationAction = "Generate" | "Analyze" | "Generate Dataset";
type OperationOutcome = "running" | "success" | "error" | "cancelled" | "ready";
interface OperationStatus { action?: OperationAction; outcome: OperationOutcome; started?: number; duration?: number; }
interface Problem extends Diagnostic { action?: OperationAction; severity?: "Error" | "Warning"; }
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

interface Dataset {
  id: string;
  name: string;
  kind: "example" | "generated" | "upload" | "file";
  text: string;
  baseline?: string;
  handle?: LocalFileHandle;
  lastModified?: number;
  exampleId?: string;
  restoreError?: string;
  needsPermission?: boolean;
}
interface TemplateEntry extends TemplateReference { file?: File; restoreError?: string; needsPermission?: boolean; }
const emptyDataset: Dataset = { id: "initial", name: "Untitled.json", kind: "generated", text: "", baseline: "" };

const repositoryUrl = "https://github.com/montr/monoreport";
// Native alignment and shifting keep right-edge triggers and long menus inside the viewport.
const dropdownProps = {
  placement: "bottomRight" as const,
  autoAdjustOverflow: true,
  align: { overflow: { adjustX: true, adjustY: true, shiftX: true, shiftY: true } },
  styles: { root: { maxWidth: "calc(100vw - 16px)" } }
};
const wordWrapPreference = workspaceStorage.wordWrap;
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

function ShortcutKeys({ value, color }: { value: string; color?: string }) {
  const keys = value.split(/[ +]+/);
  const names: Record<string, string> = { "⇧": "Shift", "⌘": "Command", "⌥": "Option", "⎋": "Escape" };
  return <Space size={4} wrap={false} className="shortcut" role="group" aria-label={keys.map(key => names[key] ?? key).join(" + ")}>
    {keys.map((key, index) => <Typography.Text keyboard aria-hidden="true" key={index} style={color ? { color } : undefined}>{key}</Typography.Text>)}
  </Space>;
}

function Filename({ name, source }: { name?: string; source?: "Example file" | "Local file" }) {
  const { token } = theme.useToken();
  return <Tooltip title={source}>
    <Typography.Text ellipsis title={name} style={{ minWidth: 0, flex: "0 1 auto", color: source === "Example file" ? token.colorTextTertiary : token.colorText }}>{name}</Typography.Text>
  </Tooltip>;
}

function restoredDiagnosticsOpen() {
  try { return localStorage.getItem(workspaceStorage.diagnosticsCollapsed) === "false"; }
  catch { return false; }
}

function resultFilename(templateName: string) {
  const templatedName = templateName.replace(/\.template(?=\.[^.]+$)/, "");
  if (templatedName !== templateName) return templatedName;
  const extension = templateName.match(/\.[^.]+$/)?.[0];
  return extension ? `${templateName.slice(0, -extension.length)}-result${extension}` : `${templateName}-result`;
}

function generatedDatasetFilename(templateName: string) {
  const basename = templateName.replace(/\.[^.]+$/, "").replace(/\.template$/, "");
  return `${basename}-generated.json`;
}

export default function App() {
  const { token } = theme.useToken();
  const [template, setTemplate] = useState<File>();
  const [datasets, setDatasets] = useState<Dataset[]>([{ id: "initial", name: "Untitled.json", kind: "generated", text: "{}", baseline: "{}" }]);
  const datasetRecords = useRef(datasets);
  const [datasetId, setDatasetId] = useState<string | undefined>("initial");
  const currentDatasetId = useRef<string | undefined>("initial");
  const dataset = datasets.find(item => item.id === datasetId);
  const json = dataset?.text ?? "";
  const [templates, setTemplates] = useState<TemplateEntry[]>([]);
  const templateRecords = useRef<TemplateEntry[]>([]);
  const [templateId, setTemplateId] = useState<string>();
  const [files] = useState(() => new LocalFiles(window as FileAccess));
  const [persistence] = useState(() => new WorkspacePersistence({
    getItem: key => localStorage.getItem(key), setItem: (key, value) => localStorage.setItem(key, value), removeItem: key => localStorage.removeItem(key)
  }, window.indexedDB, files));
  const [persistenceReady, setPersistenceReady] = useState(false);
  const [storageWarning, setStorageWarning] = useState<string>();
  const [workspaceKey, setWorkspaceKey] = useState(0);
  const resetRevision = useRef(0);
  const [clearing, setClearing] = useState(false);
  const manifestRef = useRef<ExampleManifest | undefined>(undefined);
  const reopenDataId = useRef<string | undefined>(undefined);
  const reopenTemplateId = useRef<string | undefined>(undefined);
  const [saving, setSaving] = useState(false);
  const savingRef = useRef(false);
  const editor = useRef<JsonEditorCommands>(null);
  const [wordWrap, setWordWrap] = useState(() => {
    try { return localStorage.getItem(wordWrapPreference) === "true"; }
    catch { return false; }
  });
  const nextId = useRef(0);
  const [result, setResult] = useState<GeneratedDocument>();
  const [errors, setErrors] = useState<Problem[]>([]);
  const [analysis, setAnalysis] = useState<TemplateAnalysis>();
  const analysisCache = useRef<{ source: File; analysis: TemplateAnalysis } | undefined>(undefined);
  const [sampleJson] = useState(() => new SampleJson(Math.random));
  const [jsonSettingsOpen, setJsonSettingsOpen] = useState(false);
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  const [jsonSettings] = Form.useForm<SampleJsonSettings>();
  const [diagnosticsTab, setDiagnosticsTab] = useState<DiagnosticsTab>("problems");
  const [operationStatus, setOperationStatus] = useState<OperationStatus>({ outcome: "ready" });
  const runningOperation = useRef<OperationStatus | undefined>(undefined);
  const generateAction = useRef<() => void>(() => undefined);
  generateAction.current = () => void generate();
  const [active, setActive] = useState<Tab>("template");
  const [pending, setPending] = useState<Operation>();
  const [dragging, setDragging] = useState(false);
  const [problemsOpen, setProblemsOpen] = useState(restoredDiagnosticsOpen);
  const [resultInputs, setResultInputs] = useState<{ datasetId: string | undefined; json: string; template: File; templateSource: "Example file" | "Local file" }>();
  const [modal, modalContext] = Modal.useModal();
  const [feedback, feedbackContext] = message.useMessage();
  const [examples, setExamples] = useState<ExampleDefinition[]>([]);
  const [selectedExample, setSelectedExample] = useState<ExampleSelection>();
  const [pendingExample, setPendingExample] = useState<string>();
  const [exampleLoading, setExampleLoading] = useState(true);
  const fileInput = useRef<HTMLInputElement>(null);
  const dataInput = useRef<HTMLInputElement>(null);
  const contentRevision = useRef(0);
  const operationId = useRef(0);
  const exampleLoadId = useRef(0);
  const request = useRef<AbortController | undefined>(undefined);
  const currentTemplate = useRef<File | undefined>(undefined);
  const currentJson = useRef(json);

  function changeWordWrap(value: boolean) {
    setWordWrap(value);
    try { localStorage.setItem(wordWrapPreference, String(value)); }
    catch { /* Keep editor commands usable when preference storage is unavailable. */ }
  }

  function changeProblemsOpen(open: boolean) {
    setProblemsOpen(open);
    try { localStorage.setItem(workspaceStorage.diagnosticsCollapsed, String(!open)); }
    catch { /* Keep diagnostics usable when preference storage is unavailable. */ }
  }

  useEffect(() => {
    if (!persistenceReady || clearing) return;
    let active = true;
    void persistence.save({ datasets, templates, datasetId, templateId, selectedExample }).then(warning => { if (active && warning) setStorageWarning(current => current ?? warning); });
    return () => { active = false; };
  }, [persistenceReady, clearing, datasets, templates, datasetId, templateId, selectedExample]);

  function updateDatasets(records: Dataset[]) {
    datasetRecords.current = records;
    setDatasets(records);
  }

  function selectDataset(id: string | undefined) {
    if (currentDatasetId.current === id) return;
    commitContent();
    currentDatasetId.current = id;
    currentJson.current = datasetRecords.current.find(item => item.id === id)?.text ?? "";
    setDatasetId(id);
  }

  function addDataset(record: Omit<Dataset, "id">) {
    const entry = { ...record, id: `data-${++nextId.current}` };
    updateDatasets([entry]);
    selectDataset(entry.id);
  }

  function updateTemplates(records: TemplateEntry[]) {
    templateRecords.current = records;
    setTemplates(records);
  }

  function selectTemplateEntry(entry: TemplateEntry | undefined) {
    commitContent();
    setTemplateId(entry?.id);
    if (currentTemplate.current !== entry?.file) {
      analysisCache.current = undefined;
      setAnalysis(undefined);
    }
    currentTemplate.current = entry?.file;
    setTemplate(entry?.file);
    setSelectedExample(entry?.exampleId ?? "custom");
    setErrors([]);
    if (entry) setActive("template");
  }

  async function confirmDiscard(record: Dataset | undefined, title: string, content: string, okText: string) {
    return !record || record.text === record.baseline || Boolean(await modal.confirm({ title, content, okText }));
  }

  function fileError(error: unknown, title: string) {
    if (error instanceof DOMException && error.name === "AbortError") return;
    setErrors(diagnostic(title, error instanceof Error || error instanceof DOMException ? error.message : "The file operation failed."));
    openDiagnostics("problems");
  }

  async function openData(file?: File, replaceId?: string, handle?: LocalFileHandle) {
    const reset = resetRevision.current;
    const revision = contentRevision.current;
    const record = datasetRecords.current.find(item => item.id === currentDatasetId.current);
    try {
      const opened = file ? { file, handle } : await files.open("json");
      if (!opened || reset !== resetRevision.current || revision !== contentRevision.current || (replaceId && record?.id !== replaceId)) return;
      if (record && !record.restoreError && await files.sameFile(record.handle, opened.handle)) return;
      const text = await files.readJson(opened.file);
      if (reset !== resetRevision.current || revision !== contentRevision.current) return;
      if (!await confirmDiscard(record, "Open dataset?", "Replace the current dataset, discarding unsaved changes?", "Replace")) return;
      if (revision !== contentRevision.current) return;
      addDataset({ name: opened.file.name, kind: opened.handle ? "file" : "upload", text, baseline: text, handle: opened.handle, lastModified: opened.handle ? opened.file.lastModified : undefined });
      setErrors([]);
    } catch (error) { if (reset === resetRevision.current) fileError(error, "Dataset could not be opened"); }
  }

  function chooseDataFile() {
    reopenDataId.current = undefined;
    if (files.canOpen) void openData();
    else dataInput.current?.click();
  }

  async function saveData(saveAs = false) {
    const snapshot = datasetRecords.current.find(item => item.id === currentDatasetId.current);
    if (!snapshot || snapshot.restoreError || savingRef.current) return;
    if (!files.canSave) {
      downloadBlob(new Blob([snapshot.text], { type: "application/json" }), snapshot.name);
      return;
    }
    savingRef.current = true;
    setSaving(true);
    try {
      const handle = !saveAs && snapshot.handle ? snapshot.handle : await files.destination(snapshot.name);
      for (const item of datasetRecords.current) {
        if (item.id !== snapshot.id && await files.sameFile(item.handle, handle)) throw new Error("This destination is now the active dataset. Save that dataset directly instead.");
      }
      if (await files.sameFile(snapshot.handle, handle) && await files.hasChanged(handle, snapshot.lastModified)) {
        if (!await modal.confirm({ title: "File changed on disk", content: `${snapshot.name} has been modified outside MonoReport. Overwrite it with this dataset?`, okText: "Overwrite", cancelText: "Cancel" })) return;
      }
      const lastModified = await files.write(handle, snapshot.text);
      updateDatasets(datasetRecords.current.map(item => item.id === snapshot.id ? { ...item, kind: "file", name: handle.name, handle, baseline: snapshot.text, lastModified } : item));
      void feedback.success("Saved.");
    } catch (error) { fileError(error, "Dataset could not be saved"); }
    finally { savingRef.current = false; setSaving(false); }
  }

  async function openTemplate(replaceId?: string) {
    const reset = resetRevision.current;
    if (!files.canOpen) { reopenTemplateId.current = replaceId; fileInput.current?.click(); return; }
    try {
      const opened = await files.open("template");
      if (!opened || reset !== resetRevision.current || (replaceId && !templateRecords.current.some(item => item.id === replaceId))) return;
      for (const entry of templateRecords.current) {
        if (entry.id !== replaceId && !entry.restoreError && await files.sameFile(entry.handle, opened.handle)) {
          if (reset !== resetRevision.current) return;
          if (replaceId) updateTemplates(templateRecords.current.filter(item => item.id !== replaceId));
          selectTemplateEntry(entry); return;
        }
      }
      if (reset === resetRevision.current) acceptTemplate(opened.file, opened.handle, replaceId);
    } catch (error) { if (reset === resetRevision.current) fileError(error, "Template could not be opened"); }
  }

  async function restoreWorkspace(context: WorkspaceContext, catalog: ExampleDefinition[], loadId: number, expectedRevision: number) {
    const restoreData = async (entry: DataReference): Promise<Dataset> => {
      try {
        let file: File;
        if (entry.kind === "example" && entry.exampleId) {
          const example = catalog.find(item => item.id === entry.exampleId);
          if (!example) throw new Error("This example is no longer available. Reopen a JSON file to continue.");
          const response = await fetch(example.dataPath, { cache: "no-cache" });
          if (!response.ok) throw new Error("The example dataset could not be reopened.");
          file = new File([await response.text()], filenameFromPath(example.dataPath));
        } else {
          if (!entry.handle) throw new Error("Only file names and reopen references are stored. Session-only content is not recoverable; Reopen this dataset to continue.");
          file = await files.reopen(entry.handle);
        }
        const text = await files.readJson(file);
        return { ...entry, name: file.name, text, baseline: text, lastModified: entry.handle ? file.lastModified : undefined };
      } catch (error) {
        return { ...entry, text: "", baseline: "", restoreError: error instanceof Error || error instanceof DOMException ? error.message : "This dataset requires Reopen.", needsPermission: error instanceof DOMException && error.name === "NotAllowedError" };
      }
    };
    const restoreTemplate = async (entry: TemplateReference): Promise<TemplateEntry> => {
      try {
        let file: File;
        if (entry.exampleId) {
          const example = catalog.find(item => item.id === entry.exampleId);
          if (!example) throw new Error("This example is no longer available. Reopen the template to continue.");
          const response = await fetch(example.templatePath, { cache: "no-cache" });
          if (!response.ok) throw new Error("The example template could not be reopened.");
          file = new File([await response.blob()], filenameFromPath(example.templatePath), { type: documentMimes[example.format] });
        } else {
          if (!entry.handle) throw new Error("Template bytes are not stored. Reopen this template to continue.");
          file = await files.reopen(entry.handle);
        }
        if (!templateFormat(file.name)) throw new Error("Select a .docx or .xlsx file.");
        return { ...entry, name: file.name, file };
      } catch (error) {
        return { ...entry, restoreError: error instanceof Error || error instanceof DOMException ? error.message : "This template requires Reopen.", needsPermission: error instanceof DOMException && error.name === "NotAllowedError" };
      }
    };
    const [restoredData, restoredTemplates] = await Promise.all([Promise.all(context.datasets.filter(entry => entry.id === context.datasetId).map(restoreData)), Promise.all(context.templates.filter(entry => entry.id === context.templateId).map(restoreTemplate))]);
    if (exampleLoadId.current !== loadId || contentRevision.current !== expectedRevision) return;
    nextId.current = Math.max(nextId.current, ...[...restoredData, ...restoredTemplates].map(item => Number(item.id.split("-").at(-1)) || 0));
    if (!restoredData.length) restoredData.push({ ...emptyDataset });
    updateDatasets(restoredData);
    updateTemplates(restoredTemplates);
    currentDatasetId.current = restoredData[0].id;
    currentJson.current = restoredData[0].text;
    setDatasetId(restoredData[0].id);
    const entry = restoredTemplates.find(item => item.id === context.templateId);
    setTemplateId(entry?.id);
    currentTemplate.current = entry?.file;
    setTemplate(entry?.file);
    setSelectedExample(context.selectedExample);
    setActive("template");
    setExampleLoading(false);
  }

  useEffect(() => {
    let active = true;
    const loadId = ++exampleLoadId.current;
    const expectedRevision = contentRevision.current;
    void (async () => {
      const stored = await persistence.load();
      if (!active) return;
      setStorageWarning(stored.warning);
      try {
        const response = await fetch("/examples.json", { cache: "no-cache" });
        if (!response.ok) throw new Error("The examples registry could not be loaded.");
        const manifest = validateManifest(await response.json());
        if (!active) return;
        manifestRef.current = manifest;
        setExamples(manifest.examples);
        if (exampleLoadId.current !== loadId || contentRevision.current !== expectedRevision) return;
        if (stored.context) await restoreWorkspace(stored.context, manifest.examples, loadId, expectedRevision);
        else await loadExample(manifest.examples.find(example => example.id === manifest.defaultExample)!, loadId, expectedRevision);
      } catch (error) {
        if (!active || contentRevision.current !== expectedRevision) return;
        if (stored.context) await restoreWorkspace(stored.context, [], loadId, expectedRevision);
        setExampleLoading(false);
        setErrors(diagnostic("Examples unavailable", error instanceof Error ? error.message : "The examples registry could not be loaded."));
        openDiagnostics("problems");
      } finally { if (active) setPersistenceReady(true); }
    })();
    return () => {
      active = false;
      operationId.current += 1;
      exampleLoadId.current += 1;
      request.current?.abort();
    };
  }, []);

  async function recoverData(record: Dataset) {
    if (!record.needsPermission || !record.handle?.requestPermission) {
      if (files.canOpen) await openData(undefined, record.id);
      else { reopenDataId.current = record.id; dataInput.current?.click(); }
      return;
    }
    const revision = contentRevision.current;
    try {
      const file = await files.reopen(record.handle, true);
      if (revision === contentRevision.current && datasetRecords.current.includes(record)) await openData(file, record.id, record.handle);
    } catch (error) { if (revision === contentRevision.current) fileError(error, "Dataset could not be reopened"); }
  }

  async function recoverTemplate(record: TemplateEntry) {
    if (!record.needsPermission || !record.handle?.requestPermission) { await openTemplate(record.id); return; }
    const revision = contentRevision.current;
    try {
      const file = await files.reopen(record.handle, true);
      if (revision === contentRevision.current && templateRecords.current.includes(record)) acceptTemplate(file, record.handle, record.id);
    } catch (error) { if (revision === contentRevision.current) fileError(error, "Template could not be reopened"); }
  }

  function openDiagnostics(tab: DiagnosticsTab) {
    setDiagnosticsTab(tab);
    changeProblemsOpen(true);
  }

  function startOperation(action: OperationAction, started = performance.now()) {
    const operation: OperationStatus = { action, started, outcome: "running" };
    runningOperation.current = operation;
    setOperationStatus(operation);
  }

  function finishOperation(outcome: Exclude<OperationOutcome, "running" | "ready">) {
    const operation = runningOperation.current;
    if (!operation) return;
    const finished = { ...operation, outcome, duration: performance.now() - (operation.started ?? performance.now()) };
    runningOperation.current = undefined;
    setOperationStatus(finished);
  }

  useEffect(() => {
    function generateShortcut(event: KeyboardEvent) {
      if (event.defaultPrevented || event.isComposing || event.repeat || event.altKey || event.shiftKey || jsonSettingsOpen || shortcutsOpen) return;
      if (event.key !== "Enter" || event.metaKey === event.ctrlKey) return;
      const target = event.target instanceof Element ? event.target : undefined;
      if (target?.closest('[role="dialog"], [role="alertdialog"]')) return;
      // Monaco binds Ctrl/Cmd+Enter to insert a line below; generation takes precedence for this workspace shortcut.
      event.preventDefault();
      event.stopPropagation();
      generateAction.current();
    }
    function keydown(event: KeyboardEvent) {
      if (event.defaultPrevented || event.isComposing || event.repeat || event.altKey || jsonSettingsOpen || shortcutsOpen) return;
      const target = event.target instanceof Element ? event.target : undefined;
      if (target?.closest('[role="dialog"], [role="alertdialog"]')) return;
      const workspaceModifier = event.metaKey !== event.ctrlKey;
      const saveShortcut = event.key.toLowerCase() === "s" && workspaceModifier;
      if (saveShortcut) {
        const editable = target?.closest('input, textarea, select, [contenteditable]:not([contenteditable="false"]), [role="textbox"], [role="combobox"]');
        if (editable && !target?.closest(".monaco-editor")) return;
        event.preventDefault();
        if (event.shiftKey) void saveData(true);
        else void saveData();
        return;
      }
      if (event.ctrlKey || event.metaKey) return;
      // Native controls and Monaco handle their bindings before this bubble listener.
      if (event.key === "Escape" && event.shiftKey && problemsOpen) {
        event.preventDefault();
        changeProblemsOpen(false);
      } else if (event.key === "?" && event.shiftKey && !target?.closest('input, textarea, select, [contenteditable]:not([contenteditable="false"]), [role="textbox"], [role="combobox"], .monaco-editor')) {
        event.preventDefault();
        setShortcutsOpen(true);
      }
    }
    document.addEventListener("keydown", generateShortcut, true);
    document.addEventListener("keydown", keydown);
    return () => {
      document.removeEventListener("keydown", generateShortcut, true);
      document.removeEventListener("keydown", keydown);
    };
  }, [problemsOpen, jsonSettingsOpen, shortcutsOpen]);

  function invalidateOperation() {
    finishOperation("cancelled");
    operationId.current += 1;
    request.current?.abort();
    request.current = undefined;
    setPending(undefined);
  }

  function commitContent() {
    contentRevision.current += 1;
    exampleLoadId.current += 1;
    setPendingExample(undefined);
    setExampleLoading(false);
    invalidateOperation();
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
      addDataset({ name: filenameFromPath(example.dataPath), kind: "example", text: data, baseline: data, exampleId: example.id });
      const entry = { id: `template-${++nextId.current}`, name: file.name, file, exampleId: example.id };
      updateTemplates([entry]);
      selectTemplateEntry(entry);
      setSelectedExample(example.id);
      setActive("template");
      setPendingExample(undefined);
      setExampleLoading(false);
      setErrors([]);
    } catch (error) {
      if (exampleLoadId.current !== loadId || contentRevision.current !== expectedRevision) return;
      setErrors(diagnostic("Example could not be loaded", error instanceof SyntaxError ? "The example JSON file is malformed." : error instanceof Error ? error.message : "The selected example could not be loaded."));
      setPendingExample(undefined);
      openDiagnostics("problems");
    } finally {
      if (exampleLoadId.current === loadId) setExampleLoading(false);
    }
  }

  async function chooseExample(id: string) {
    const example = examples.find(item => item.id === id);
    if (!example) return;
    if (id === pendingExample) return;
    if (id === selectedExample && templateRecords.current.find(item => item.id === templateId)?.exampleId === id && dataset?.kind === "example" && dataset.exampleId === id && dataset.text === dataset.baseline) {
      setActive("template");
      if (exampleLoading) {
        exampleLoadId.current += 1;
        setPendingExample(undefined);
        setExampleLoading(false);
      }
      return;
    }
    const expectedRevision = contentRevision.current;
    const record = datasetRecords.current.find(item => item.id === currentDatasetId.current);
    if (!await confirmDiscard(record, "Open example?", "Replace the current template and dataset with the selected example, discarding unsaved changes?", "Replace")) return;
    if (contentRevision.current !== expectedRevision) return;
    setPendingExample(id);
    void loadExample(example);
  }

  function acceptTemplate(file: File, handle?: LocalFileHandle, replaceId?: string) {
    if (!templateFormat(file.name)) {
      setErrors(diagnostic("Unsupported template format", "Select a .docx or .xlsx file."));
      openDiagnostics("problems");
      return;
    }
    if (replaceId && !templateRecords.current.some(item => item.id === replaceId)) return;
    const entry = { id: `template-${++nextId.current}`, name: file.name, file, handle };
    updateTemplates([entry]);
    selectTemplateEntry(entry);
  }

  function selected(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    const replaceId = reopenTemplateId.current;
    reopenTemplateId.current = undefined;
    if (file) acceptTemplate(file, undefined, replaceId);
    event.target.value = "";
  }

  function drop(event: DragEvent) {
    event.preventDefault();
    setDragging(false);
    const file = event.dataTransfer.files?.[0];
    if (file) acceptTemplate(file);
  }

  function changeJson(value: string) {
    if (value === currentJson.current) return;
    commitContent();
    currentJson.current = value;
    updateDatasets(datasetRecords.current.map(item => item.id === currentDatasetId.current ? { ...item, text: value } : item));
  }

  async function copyJson() {
    try {
      await editor.current?.copy();
      void feedback.success("Copied to clipboard.");
    } catch {
      void feedback.error("The JSON could not be copied to the clipboard.");
    }
  }

  async function copyProblems() {
    try {
      await navigator.clipboard.writeText(JSON.stringify(errors, null, 2));
      void feedback.success("Diagnostics copied to clipboard.");
    } catch {
      void feedback.error("The diagnostics could not be copied to the clipboard.");
    }
  }

  async function formatJson() {
    try {
      JSON.parse(json);
      await editor.current?.format();
      setErrors([]);
    } catch {
      setErrors(diagnostic("Invalid JSON", "Fix the JSON syntax before formatting."));
      openDiagnostics("problems");
    }
  }

  async function generate() {
    if (pending || runningOperation.current || jsonSettingsOpen) return;
    invalidateOperation();
    startOperation("Generate");
    if (!template) {
      setErrors(diagnostic("Template is required", "Select a DOCX or XLSX file before generating a document.").map(item => ({ ...item, action: "Generate" })));
      finishOperation("error");
      openDiagnostics("problems");
      return;
    }
    try {
      JSON.parse(json);
    } catch {
      setErrors(diagnostic("Invalid JSON", "Fix the JSON syntax before generating a document.").map(item => ({ ...item, action: "Generate" })));
      finishOperation("error");
      openDiagnostics("problems");
      return;
    }

    const operation = operationId.current;
    const expectedRevision = contentRevision.current;
    const controller = new AbortController();
    request.current = controller;
    setErrors([]);
    setPending("generate");
    try {
      const generated = await generateDocument(template, json, controller.signal);
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return;
      setResult({ ...generated, filename: resultFilename(template.name) });
      const warnings = (generated.warnings ?? []).map(item => ({ ...item, action: "Generate" as const, severity: "Warning" as const }));
      setErrors(warnings);
      setResultInputs({ datasetId, json, template, templateSource: selectedTemplate?.exampleId ? "Example file" : "Local file" });
      setActive("result");
      finishOperation("success");
      if (warnings.length) openDiagnostics("problems");
    } catch (error) {
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return;
      if (error instanceof DOMException && error.name === "AbortError") { finishOperation("cancelled"); return; }
      setErrors((error instanceof GenerationFailure ? error.diagnostics : diagnostic("Generation failed", "The document could not be generated.")).map(item => ({ ...item, action: "Generate" })));
      finishOperation("error");
      openDiagnostics("problems");
    } finally {
      if (operationId.current === operation) {
        request.current = undefined;
        setPending(undefined);
      }
    }
  }

  function analysisProblems(value: TemplateAnalysis, action: OperationAction): Problem[] {
    return [...value.errors.map(item => ({ ...item, action, severity: "Error" as const })), ...value.warnings.map(item => ({ ...item, action, severity: "Warning" as const }))];
  }

  async function analyze(settings?: SampleJsonSettings) {
    if (pending || runningOperation.current || !template || (!settings && jsonSettingsOpen)) return;
    invalidateOperation();
    const action: OperationAction = settings ? "Generate Dataset" : "Analyze";
    startOperation(action);
    const operation = operationId.current;
    const expectedRevision = contentRevision.current;
    const source = template;
    const controller = new AbortController();
    request.current = controller;
    setErrors([]);
    setPending(settings ? "generate-json" : "analyze");
    if (!settings) {
      analysisCache.current = undefined;
      openDiagnostics("analysis");
    }
    try {
      const analyzed = settings && analysisCache.current?.source === source ? analysisCache.current.analysis : await analyzeTemplate(source, controller.signal);
      if (operationId.current !== operation || contentRevision.current !== expectedRevision || currentTemplate.current !== source) return;
      analysisCache.current = { source, analysis: analyzed };
      setAnalysis(analyzed);
      setErrors(analysisProblems(analyzed, action));
      if (settings) {
        const proposed = sampleJson.generate(analyzed.requirements, settings);
        // Finish before selecting the new dataset, which invalidates pending input work.
        finishOperation("success");
        addDataset({ name: generatedDatasetFilename(source.name), kind: "generated", text: proposed });
        setJsonSettingsOpen(false);
      } else finishOperation("success");
    } catch (error) {
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return;
      if (error instanceof DOMException && error.name === "AbortError") { finishOperation("cancelled"); return; }
      if (error instanceof AnalysisFailure && error.analysis) {
        setAnalysis(error.analysis);
        const problems = analysisProblems(error.analysis, action);
        setErrors(error.analysis.errors.length ? problems : [...problems, ...error.diagnostics.map(item => ({ ...item, action }))]);
      } else {
        const problems = error instanceof AnalysisFailure ? error.diagnostics : diagnostic(`${action} failed`, error instanceof Error ? error.message : "The operation could not be completed.");
        setErrors(problems.map(item => ({ ...item, action })));
      }
      finishOperation("error");
      if (settings) {
        setJsonSettingsOpen(false);
        openDiagnostics("problems");
      }
    } finally {
      if (operationId.current === operation) {
        request.current = undefined;
        setPending(undefined);
      }
    }
  }

  async function submitJsonSettings() {
    if (pending || runningOperation.current) return;
    const started = performance.now();
    const operation = operationId.current;
    const expectedRevision = contentRevision.current;
    try {
      const settings = await jsonSettings.validateFields();
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return;
      if (!await confirmDiscard(datasetRecords.current.find(item => item.id === currentDatasetId.current), "Generate dataset?", "Replace the current dataset, discarding unsaved changes?", "Replace")) return;
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return;
      await analyze(settings);
    } catch {
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return;
      startOperation("Generate Dataset", started);
      setErrors(diagnostic("Invalid settings", "Items per collection must be a positive integer.").map(item => ({ ...item, action: "Generate Dataset" })));
      finishOperation("error");
    }
  }

  function clearTemplate() {
    const remaining = templateRecords.current.filter(item => item.id !== templateId);
    updateTemplates(remaining);
    selectTemplateEntry(undefined);
  }

  async function clearLocalData() {
    if (savingRef.current || clearing) return;
    const expectedRevision = contentRevision.current;
    await modal.confirm({
      title: "Clear Local Data?", okText: "Clear", cancelText: "Cancel", okButtonProps: { danger: true },
      content: "This removes local preferences, remembered file handles, and current workspace data, including generated documents. Unsaved workspace changes will be discarded. Files on disk will not be deleted.",
      onOk: async () => {
        if (contentRevision.current !== expectedRevision || savingRef.current) {
          void feedback.warning("The workspace changed. Cancel and review Clear Local Data again.");
          throw new Error("The workspace changed while confirmation was open.");
        }
        setClearing(true);
        setPersistenceReady(false);
        resetRevision.current += 1;
        commitContent();
        reopenDataId.current = undefined;
        reopenTemplateId.current = undefined;
        updateDatasets([{ ...emptyDataset }]);
        currentDatasetId.current = emptyDataset.id;
        currentJson.current = "";
        setDatasetId(emptyDataset.id);
        updateTemplates([]);
        currentTemplate.current = undefined;
        setTemplate(undefined);
        setTemplateId(undefined);
        setSelectedExample(undefined);
        setResult(undefined);
        setResultInputs(undefined);
        analysisCache.current = undefined;
        setAnalysis(undefined);
        setErrors([]);
        runningOperation.current = undefined;
        setOperationStatus({ outcome: "ready" });
        setJsonSettingsOpen(false);
        setShortcutsOpen(false);
        setProblemsOpen(false);
        setDiagnosticsTab("problems");
        setActive("template");
        setWordWrap(false);
        setDragging(false);
        try {
          const warning = await persistence.clear();
          setStorageWarning(warning);
          setWorkspaceKey(value => value + 1);
          let manifest = manifestRef.current;
          if (!manifest) {
            const response = await fetch("/examples.json", { cache: "no-cache" });
            if (!response.ok) throw new Error("The examples registry could not be loaded.");
            manifest = validateManifest(await response.json());
            manifestRef.current = manifest;
            setExamples(manifest.examples);
          }
          await loadExample(manifest.examples.find(example => example.id === manifest.defaultExample)!);
          if (!warning) void feedback.success("Local data cleared");
        } catch (error) { fileError(error, "The default example could not be reopened"); }
        finally { setClearing(false); setPersistenceReady(true); }
      }
    });
  }

  function closeResult() {
    setResult(undefined);
    setResultInputs(undefined);
    if (template) setActive("template");
  }

  function downloadBlob(blob: Blob, filename: string) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  const selectedTemplate = templates.find(item => item.id === templateId);
  const editorAvailable = Boolean(dataset && !dataset.restoreError);
  const modified = Boolean(dataset && dataset.text !== dataset.baseline);
  const openLabel = files.canOpen ? "Open…" : "Upload…";
  const platform = navigator.platform;
  const isMac = /Mac|iPhone|iPad/.test(platform);
  const commandKey = isMac ? "⌘ " : "Ctrl+";
  const saveAsShortcut = isMac ? "⇧ ⌘ S" : "Ctrl+Shift+S";
  const formatShortcut = /Linux/.test(platform) ? "Ctrl+Shift+I" : isMac ? "⇧ ⌥ F" : "Shift+Alt+F";
  const generateShortcut = isMac ? "⌘ Enter" : "Ctrl+Enter";
  const shortcutGroups = [
    { id: "workspace", label: "Workspace", items: [
      { action: "Generate document", shortcut: generateShortcut },
      { action: "Keyboard shortcuts", shortcut: isMac ? "⇧ /" : "Shift+/" },
      { action: "Collapse diagnostics", shortcut: isMac ? "⇧ ⎋" : "Shift+Escape" }
    ] },
    ...(editorAvailable ? [{ id: "data-editor", label: "Dataset editor", items: [
      { action: files.canOpen ? "Open dataset" : "Upload dataset", shortcut: `${commandKey}O` },
      { action: "Save dataset", shortcut: `${commandKey}S` },
      { action: "Save dataset as", shortcut: saveAsShortcut },
      { action: "Format dataset", shortcut: formatShortcut }
    ] }] : [])
  ];
  const exampleItems = (["docx", "xlsx"] as const).map(format => ({
    type: "group" as const,
    key: format,
    label: format.toUpperCase(),
    children: examples.filter(example => example.format === format).map(example => ({
      key: example.id,
      label: <FormatLabel format={format} title={example.title} />
    }))
  })).filter(group => group.children.length > 0);
  const resultCurrent = Boolean(result && resultInputs?.datasetId === datasetId && resultInputs?.json === json && resultInputs?.template === template);
  const statusLabel = operationStatus.outcome === "running"
    ? operationStatus.action === "Generate" ? "Generating..." : "Analyzing..."
    : operationStatus.outcome === "ready" ? "Ready"
      : operationStatus.outcome === "success" ? "Success"
        : operationStatus.outcome === "error" ? "Error" : "Cancelled";
  const statusIcon = operationStatus.outcome === "success" ? <CheckCircleFilled aria-hidden="true" style={{ color: token.colorSuccess }} />
    : operationStatus.outcome === "error" ? <CloseCircleFilled aria-hidden="true" style={{ color: token.colorError }} />
      : operationStatus.outcome === "cancelled" ? <StopOutlined aria-hidden="true" style={{ color: token.colorWarning }} />
        : undefined;
  const source = active === "result" ? result?.blob : template;
  const format = active === "result" ? result?.format : template && templateFormat(template.name);
  const label = active === "result" ? "Result" : "Template";
  const dataMenu = { items: [
    { key: "open", label: openLabel, extra: editorAvailable ? `${commandKey}O` : undefined, onClick: chooseDataFile },
    { key: "save", label: "Save", extra: editorAvailable ? `${commandKey}S` : undefined, disabled: !editorAvailable || saving, onClick: () => void saveData() },
    { key: "save-as", label: "Save As…", extra: editorAvailable ? saveAsShortcut : undefined, disabled: !editorAvailable || saving, onClick: () => void saveData(true) },
    { type: "divider" as const, key: "file-actions-end" },
    { key: "format", label: "Format", extra: editorAvailable ? formatShortcut : undefined, disabled: !editorAvailable, onClick: () => void formatJson() },
    { key: "wrap", label: "Word wrap", role: "menuitemcheckbox", "aria-checked": wordWrap, disabled: !editorAvailable, onClick: () => editor.current?.toggleWrap() },
    { key: "copy", label: "Copy", disabled: !editorAvailable, onClick: () => void copyJson() }
  ].map(item => "type" in item ? item : { ...item, icon: item.key === "wrap" && wordWrap ? <CheckOutlined aria-hidden="true" /> : <span aria-hidden="true" /> }) };
  const documentMenu = { items: [
    ...(active === "template" ? [{ key: "upload", label: openLabel, onClick: () => void openTemplate() }] : []),
    { key: "download", label: "Download", disabled: !source, onClick: () => {
      if (active === "result" && result) downloadBlob(result.blob, result.filename);
      else if (template) downloadBlob(template, template.name);
    } },
    { type: "divider" as const, key: "file-actions-end" },
    { key: "close", label: "Close", disabled: active === "template" ? !selectedTemplate : !source, onClick: active === "template" ? clearTemplate : closeResult }
  ].map(item => "type" in item ? item : { ...item, icon: <span aria-hidden="true" /> }) };
  const problemsLabel = <Space size="small">Problems{errors.length > 0 && <Badge count={errors.length} size="small" />}</Space>;

  return <ConfigProvider>
    {modalContext}
    {feedbackContext}
    <Modal title="Keyboard shortcuts" open={shortcutsOpen} onCancel={() => setShortcutsOpen(false)} footer={null} width={440} style={{ maxWidth: "calc(100vw - 32px)" }}>
      <Flex vertical gap="small">
        {shortcutGroups.map(group => <section key={group.id} aria-labelledby={`${group.id}-shortcuts-title`}>
          <Typography.Title level={5} id={`${group.id}-shortcuts-title`} style={{ marginTop: 0 }}>{group.label}</Typography.Title>
          <Flex vertical gap="small">
            {group.items.map(item => <Flex key={item.action} justify="space-between" align="center" gap="middle" role="group" aria-label={item.action}>
              <Typography.Text>{item.action}</Typography.Text>
              <ShortcutKeys value={item.shortcut} />
            </Flex>)}
          </Flex>
        </section>)}
      </Flex>
    </Modal>
    <Modal title="Generate Dataset" open={jsonSettingsOpen} okText="OK" footer={(_, { OkBtn }) => <OkBtn />} width={400} okButtonProps={{ loading: pending === "generate-json" }} onOk={() => void submitJsonSettings()} onCancel={() => { invalidateOperation(); setJsonSettingsOpen(false); }}>
      <ConfigProvider componentSize="small" theme={{ components: { Form: { itemMarginBottom: token.marginSM } } }}>
        <Form form={jsonSettings} layout="horizontal" labelCol={{ flex: "150px" }} wrapperCol={{ flex: "auto" }} labelAlign="start" colon={false} initialValues={{ itemsPerCollection: 10, values: "empty" }}>
          <Form.Item name="itemsPerCollection" label="Items per collection" rules={[{ validator: (_, value) => Number.isInteger(value) && value > 0 ? Promise.resolve() : Promise.reject(new Error("Enter a positive integer.")) }]}>
            <InputNumber changeOnBlur={false} />
          </Form.Item>
          <Form.Item name="values" label="Values"><Radio.Group options={[{ label: "Empty", value: "empty" }, { label: "Random", value: "random" }]} /></Form.Item>
        </Form>
      </ConfigProvider>
    </Modal>
    <div className="app">
      <Flex component="header" className="app-header" align="center" justify="space-between" gap="small" wrap>
        <span className="brand">monoreport</span>
        <Flex align="center" justify="flex-end" gap="small" wrap flex="auto" style={{ maxWidth: "100%" }} aria-label="Workspace actions">
          <Tooltip title="Explore examples to see how MonoReport works">
            <Dropdown {...dropdownProps} trigger={["click"]} menu={{ items: exampleItems, selectedKeys: [pendingExample ?? selectedExample ?? ""], onClick: ({ key }) => void chooseExample(key) }}>
              <Button aria-label="Examples" disabled={examples.length === 0} icon={<BulbOutlined aria-hidden="true" />}>Examples <DownOutlined aria-hidden="true" /></Button>
            </Dropdown>
          </Tooltip>
          <Flex align="center" gap="small">
            <Space.Compact>
              <Tooltip title={<Space size="small"><span>Generate document</span><ShortcutKeys value={generateShortcut} color={token.colorTextLightSolid} /></Space>}>
                <Button type="primary" aria-label="Generate document" loading={pending === "generate"} disabled={Boolean(pending) || Boolean(dataset?.restoreError) || Boolean(selectedTemplate?.restoreError)} onClick={() => void generate()}>Generate</Button>
              </Tooltip>
              <Dropdown {...dropdownProps} trigger={["click"]} menu={{ items: [
                { key: "analyze", label: "Analyze Template", disabled: !template || Boolean(pending), onClick: () => void analyze() },
                { key: "generate-json", label: "Generate Dataset...", disabled: !template || Boolean(pending), onClick: () => setJsonSettingsOpen(true) }
              ] }}>
                <Button type="primary" aria-label="Generate actions" icon={<DownOutlined />} disabled={Boolean(pending)} />
              </Dropdown>
            </Space.Compact>
            <Dropdown {...dropdownProps} trigger={["click"]} menu={{ items: [
              { key: "api", label: <a href="/docs/" target="_blank" rel="noopener noreferrer">API documentation</a> },
              { key: "shortcuts", label: "Keyboard shortcuts", extra: isMac ? "⇧ /" : "Shift+/", onClick: () => setShortcutsOpen(true) },
              { type: "divider" },
              { key: "clear-local", label: "Clear Local Data…", disabled: saving || clearing, onClick: () => void clearLocalData() }
            ] }}>
              <Button type="text" icon={<EllipsisOutlined />} aria-label="Application menu" />
            </Dropdown>
          </Flex>
        </Flex>
      </Flex>
      {storageWarning && <Alert type="warning" title={storageWarning} showIcon />}
      <main className="layout">
        <WorkspaceLayout key={workspaceKey} problemsOpen={problemsOpen} onProblemsOpenChange={changeProblemsOpen}
          data={<section className="source" aria-label="Input dataset">
            <Flex className="pane-header" align="center" gap="small">
              <Flex align="center" gap="middle" flex="auto" style={{ minWidth: 0 }}>
                <Typography.Text strong style={{ flexShrink: 0 }}>Dataset</Typography.Text>
                <Flex className="context-name" align="center" gap="small" wrap={false}>
                  <Filename name={dataset?.name} source={dataset?.kind === "example" ? "Example file" : dataset?.kind === "file" || dataset?.kind === "upload" ? "Local file" : undefined} />
                  {modified && <Badge status="default" color={token.colorTextSecondary} aria-label="Dataset modified" style={{ flex: "none" }} />}
                </Flex>
              </Flex>
              <Dropdown {...dropdownProps} trigger={["click"]} menu={dataMenu}><Button size="small" type="text" icon={<EllipsisOutlined />} aria-label="Dataset menu" /></Dropdown>
            </Flex>
            {dataset?.restoreError && <Alert type="warning" title={`${dataset.name}: ${dataset.restoreError}`} action={<Button size="small" onClick={() => void recoverData(dataset)}>{dataset.needsPermission && dataset.handle?.requestPermission ? "Grant access" : "Reopen…"}</Button>} />}
            <MonacoJsonEditor ref={editor} datasetId={dataset?.id} value={json} onChange={changeJson} wordWrap={wordWrap} onWordWrapChange={changeWordWrap} fileCommands={{ open: chooseDataFile, ...(editorAvailable ? { save: () => void saveData(), saveAs: () => void saveData(true) } : {}) }} />
          </section>}
          document={<section className="preview" aria-label="Preview">
            <Flex className="pane-header" align="center" gap="small" wrap>
              <Radio.Group size="small" value={active} onChange={event => setActive(event.target.value)} aria-label="Document view">
                <Radio.Button value="template">Template</Radio.Button>
                <Radio.Button value="result">Result</Radio.Button>
              </Radio.Group>
              <Flex className="context-name" align="center" gap="small" wrap={false}>
                <Filename name={active === "result" ? result?.filename : selectedTemplate?.name} source={active === "result" ? resultInputs?.templateSource : selectedTemplate?.exampleId ? "Example file" : "Local file"} />
              </Flex>
              {(pending || exampleLoading) && <Spin size="small" aria-label={exampleLoading ? "Loading example" : pending === "generate" ? "Generating document" : pending === "analyze" ? "Analyzing template" : "Generating Dataset"} />}
              {result && active === "result" && !resultCurrent && <Tooltip title="The Dataset or Template changed since this Result was generated."><Tag color="warning" variant="outlined" icon={<ExclamationCircleOutlined />} style={{ flex: "none" }}>Outdated</Tag></Tooltip>}
              <Dropdown {...dropdownProps} trigger={["click"]} menu={documentMenu}><Button size="small" type="text" icon={<EllipsisOutlined />} aria-label="Document menu" /></Dropdown>
            </Flex>
            {selectedTemplate?.restoreError && <Alert type="warning" title={`${selectedTemplate.name}: ${selectedTemplate.restoreError}`} action={<Button size="small" onClick={() => void recoverTemplate(selectedTemplate)}>{selectedTemplate.needsPermission && selectedTemplate.handle?.requestPermission ? "Grant access" : "Reopen…"}</Button>} />}
            <div className={`canvas ${format === "xlsx" ? "xlsx-view" : ""} ${dragging ? "drop-target" : ""}`}
              onDragEnter={event => { if ((source || active === "result") && event.dataTransfer.types.includes("Files")) { event.preventDefault(); setDragging(true); } }}
              onDragOver={event => { if (event.dataTransfer.types.includes("Files")) event.preventDefault(); }}
              onDragLeave={event => { if (!event.currentTarget.contains(event.relatedTarget as Node)) setDragging(false); }}
              onDrop={drop}>
              {source ? format === "xlsx" ? <XlsxPreview source={source} label={label} /> : <DocxPreview source={source} label={label} />
                : active === "result" ? <Empty description="Not generated" />
                : files.canOpen ? <Empty description="Select a DOCX or XLSX template"><Button onClick={() => void openTemplate()}>{openLabel}</Button></Empty>
                : <Upload.Dragger accept={templateAccept} showUploadList={false} onDrop={event => event.stopPropagation()} beforeUpload={file => { setDragging(false); acceptTemplate(file); return false; }}>
                    <Empty description="Select a DOCX or XLSX template"><Button icon={<UploadOutlined aria-hidden="true" />}>Upload…</Button></Empty>
                  </Upload.Dragger>}
            </div>
          </section>}
          problems={<ConfigProvider componentSize="small" theme={{ token: { fontSize: 12, fontSizeSM: 12, fontSizeLG: 12, fontSizeHeading5: 12 }, components: {
            Button: { contentFontSizeSM: 12 },
            Tabs: { titleFontSizeSM: 12, horizontalItemPaddingSM: `${token.paddingXXS}px ${token.paddingSM}px`, horizontalItemGutter: token.marginXS, horizontalMargin: "0" },
            Table: { cellFontSizeSM: 12, cellPaddingBlockSM: token.paddingXXS, cellPaddingInlineSM: token.paddingXS },
            Typography: { titleMarginTop: token.marginXS, titleMarginBottom: token.marginXXS }
          } }}>
          <section className="problems" aria-label="Diagnostics" style={{ fontSize: 12 }}>
            <Tabs size="small" styles={{ root: { height: "100%", overflow: "hidden" }, header: { background: token.colorFillQuaternary, paddingBlock: 0 }, body: { height: "100%" }, content: { height: "100%", overflow: "auto" } }}
              activeKey={diagnosticsTab} onChange={key => setDiagnosticsTab(key as DiagnosticsTab)} tabBarExtraContent={<Tooltip title="Hide diagnostics"><Button size="small" type="text" icon={<CloseOutlined />} aria-label="Hide diagnostics" onClick={() => changeProblemsOpen(false)} /></Tooltip>} items={[{
              key: "problems", label: problemsLabel, children: <>
                {errors.length > 0 && <Flex role="toolbar" aria-label="Problems actions" align="center" gap={token.paddingXXS} style={{ marginTop: token.paddingXXS, marginBottom: token.marginXXS }}>
                  <Tooltip title="Clear"><Button size="small" type="text" icon={<DeleteOutlined />} aria-label="Clear" onClick={() => setErrors([])} /></Tooltip>
                  <Tooltip title="Copy"><Button size="small" type="text" icon={<CopyOutlined />} aria-label="Copy" onClick={() => void copyProblems()} /></Tooltip>
                </Flex>}
                {errors.length ? <div>{errors.map((error, index) => {
                  const severity = error.severity ?? "Error";
                  const Icon = severity === "Warning" ? ExclamationCircleFilled : CloseCircleFilled;
                  const code = error.code ?? error.title;
                  const title = error.title === code ? undefined : error.title;
                  return <div className="issue" key={`${code}-${index}`}>
                    <div className="issue-content">
                      <div className="issue-heading"><div role="img" aria-label={severity} className="issue-icon" style={{ color: severity === "Warning" ? token.colorWarning : token.colorError }}><Icon aria-hidden="true" /></div><strong>{code}</strong>{title && <span>{title}</span>}</div>
                      <p>{[error.field, error.fieldCode, error.dataPath, error.part, error.location, error.formatter, error.removedCharacterCount === undefined ? undefined : `removed ${error.removedCharacterCount}`, error.traceId === undefined ? undefined : `trace ${error.traceId}`, error.limitReached ? "error limit reached" : undefined, error.detail].filter(Boolean).join(" · ")}</p>
                    </div>
                  </div>;
                })}</div> : <Empty image={Empty.PRESENTED_IMAGE_SIMPLE} description="No problems" />}
              </>
            }, { key: "analysis", label: "Analysis", children: <AnalysisView analysis={analysis} /> }]} />
          </section>
          </ConfigProvider>}
        />
      </main>
      <input ref={fileInput} type="file" hidden accept={templateAccept} onChange={selected} />
      <input ref={dataInput} type="file" hidden accept=".json,application/json" aria-label="Upload dataset" onChange={event => { const file = event.target.files?.[0]; const replaceId = reopenDataId.current; reopenDataId.current = undefined; event.target.value = ""; if (file) void openData(file, replaceId); }} />
      <ConfigProvider componentSize="small" theme={{ token: { fontSize: 12, fontSizeSM: 12, fontSizeLG: 12 }, components: { Button: { contentFontSizeSM: 12 } } }}>
      <Flex component="footer" align="center" gap="small" wrap style={{ fontSize: 12 }}>
        <Button size="small" type="text" style={{ paddingInlineStart: 0 }} onClick={() => openDiagnostics("problems")} aria-label={`Workspace status: ${statusLabel}`}><Space size="small"><span role="status" aria-label={statusLabel}><Flex align="center" gap={token.paddingXXS}><span style={{ display: "inline-flex", alignItems: "center" }}>{operationStatus.outcome === "running" ? <Spin size="small" /> : statusIcon}</span><span>{statusLabel}</span></Flex></span>{operationStatus.outcome !== "running" && operationStatus.duration !== undefined && <Typography.Text type="secondary">{operationStatus.duration.toFixed(0)} ms</Typography.Text>}</Space></Button>
        {!problemsOpen && <>
        <Button size="small" type="text" aria-expanded={problemsOpen && diagnosticsTab === "problems"} onClick={() => openDiagnostics("problems")}>{problemsLabel}</Button>
        <Button size="small" type="text" aria-expanded={problemsOpen && diagnosticsTab === "analysis"} onClick={() => openDiagnostics("analysis")}>Analysis</Button>
        </>}
        <span className="project-links"><a href={repositoryUrl} target="_blank" rel="noreferrer" title="MonoReport on GitHub" style={{ color: token.colorTextTertiary }}><GithubFilled aria-hidden="true" />MonoReport</a>{__APP_VERSION__ === "dev" ? <span style={{ color: token.colorTextTertiary }}>dev</span> : <a href={`${repositoryUrl}/releases/tag/v${__APP_VERSION__}`} target="_blank" rel="noreferrer" style={{ color: token.colorTextTertiary }}>v{__APP_VERSION__}</a>}</span>
      </Flex>
      </ConfigProvider>
    </div>
  </ConfigProvider>;
}
