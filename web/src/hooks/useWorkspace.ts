import { message, Modal, theme } from "antd";
import { useEffect, useRef, useState } from "react";
import { type JsonEditorCommands } from "../components/MonacoJsonEditor";
import { LocalFiles, type FileAccess } from "../services/localFiles";
import { WorkspacePersistence, workspaceStorage, type WorkspaceContext } from "../services/workspacePersistence";
import { WorkspaceRestorer } from "../services/workspaceRestorer";
import { useTemplateFileActions } from "./useTemplateFileActions";
import { useDatasetFileActions } from "./useDatasetFileActions";
import { useDocumentActions } from "./useDocumentActions";
import { useWorkspaceShortcuts } from "./useWorkspaceShortcuts";
import { ExampleCatalog, filenameFromPath, type ExampleDefinition, type ExampleManifest } from "../services/examples";
import type { Dataset, TemplateEntry, Tab, DiagnosticsTab, ExampleSelection, Problem, Diagnostic } from "../types";
const emptyDataset: Dataset = { id: "initial", name: "Untitled.json", kind: "generated", text: "", baseline: "" };

const wordWrapPreference = workspaceStorage.wordWrap;
function diagnostic(title: string, detail: string): Diagnostic[] {
  return [{ title, detail }];
}

function restoredDiagnosticsOpen() {
  try { return localStorage.getItem(workspaceStorage.diagnosticsCollapsed) === "false"; }
  catch { return false; }
}

export function useWorkspace() {
  const { token } = theme.useToken();
  const [templateEntry, setTemplateEntry] = useState<TemplateEntry>();
  const templateRecord = useRef<TemplateEntry | undefined>(undefined);
  const template = templateEntry?.file;
  const [dataset, setDataset] = useState<Dataset>({ id: "initial", name: "Untitled.json", kind: "generated", text: "{}", baseline: "{}" });
  const datasetRecord = useRef(dataset);
  const datasetId = dataset.id;
  const json = dataset?.text ?? "";
  const [files] = useState(() => new LocalFiles(window as FileAccess));
  const [persistence] = useState(() => new WorkspacePersistence({
    getItem: key => localStorage.getItem(key), setItem: (key, value) => localStorage.setItem(key, value), removeItem: key => localStorage.removeItem(key)
  }, window.indexedDB, files));
  const [exampleCatalog] = useState(() => new ExampleCatalog((...args) => fetch(...args)));
  const [restorer] = useState(() => new WorkspaceRestorer(files));
  const [persistenceReady, setPersistenceReady] = useState(false);
  const [storageWarning, setStorageWarning] = useState<string>();
  const [workspaceKey, setWorkspaceKey] = useState(0);
  const resetRevision = useRef(0);
  const [clearing, setClearing] = useState(false);
  const manifestRef = useRef<ExampleManifest | undefined>(undefined);
  const editor = useRef<JsonEditorCommands>(null);
  const [wordWrap, setWordWrap] = useState(() => {
    try { return localStorage.getItem(wordWrapPreference) === "true"; }
    catch { return false; }
  });
  const nextId = useRef(0);
  const [errors, setErrors] = useState<Problem[]>([]);
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  const [diagnosticsTab, setDiagnosticsTab] = useState<DiagnosticsTab>("problems");
  const generateAction = useRef<() => void>(() => undefined);
  const [active, setActive] = useState<Tab>("template");
  const [problemsOpen, setProblemsOpen] = useState(restoredDiagnosticsOpen);
  const [modal, modalContext] = Modal.useModal();
  const [feedback, feedbackContext] = message.useMessage();
  const [examples, setExamples] = useState<ExampleDefinition[]>([]);
  const [selectedExample, setSelectedExample] = useState<ExampleSelection>();
  const [pendingExample, setPendingExample] = useState<string>();
  const [exampleLoading, setExampleLoading] = useState(true);
  const contentRevision = useRef(0);
  const exampleLoadId = useRef(0);
  const currentTemplate = useRef<File | undefined>(undefined);
  const currentJson = useRef(json);
  const { result, resultInputs, analysis, clearAnalysis, closeResult, resetDocumentActions, disposeDocumentActions, jsonSettingsOpen, setJsonSettingsOpen, jsonSettings, operationStatus, pending, invalidateOperation, generate, analyze, submitJsonSettings } = useDocumentActions({ template, datasetId, json, selectedTemplate: templateEntry, currentTemplate, contentRevision, datasetRecord, setErrors, openDiagnostics, setActive, addDataset, confirmDiscard });
  generateAction.current = () => void generate();
  const { openData, chooseDataFile, saveData, recoverData, saving, savingRef, dataInput, reopenDataId } = useDatasetFileActions({ files, resetRevision, contentRevision, datasetRecord, confirmDiscard, addDataset, setDatasetRecord, setErrors, fileError, downloadBlob, modal, feedback });
  const { openTemplate, recoverTemplate, acceptTemplate, selected, drop, clearTemplate, dragging, setDragging, fileInput, reopenTemplateId } = useTemplateFileActions({ files, resetRevision, contentRevision, templateRecord, nextId, selectTemplateEntry, setErrors, fileError, openDiagnostics });

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
    void persistence.save({ datasets: [dataset], templates: templateEntry ? [templateEntry] : [], datasetId, templateId: templateEntry?.id, selectedExample }).then(warning => { if (active && warning) setStorageWarning(current => current ?? warning); });
    return () => { active = false; };
  }, [persistenceReady, clearing, dataset, templateEntry, datasetId, selectedExample]);

  function addDataset(record: Omit<Dataset, "id">) {
    const entry = { ...record, id: `data-${++nextId.current}` };
    commitContent();
    setDatasetRecord(entry);
  }

  function setDatasetRecord(record: Dataset) {
    datasetRecord.current = record;
    currentJson.current = record.text;
    setDataset(record);
  }

  function selectTemplateEntry(entry: TemplateEntry | undefined) {
    commitContent();
    if (currentTemplate.current !== entry?.file) {
      clearAnalysis();
    }
    currentTemplate.current = entry?.file;
    templateRecord.current = entry;
    setTemplateEntry(entry);
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

  async function restoreWorkspace(context: WorkspaceContext, catalog: ExampleDefinition[], loadId: number, expectedRevision: number) {
    const { dataset: restoredData, template: restoredTemplate } = await restorer.restore(context, catalog);
    if (exampleLoadId.current !== loadId || contentRevision.current !== expectedRevision) return;
    const dataset = restoredData ?? { ...emptyDataset };
    nextId.current = Math.max(nextId.current, Number(dataset.id.split("-").at(-1)) || 0, Number(restoredTemplate?.id.split("-").at(-1)) || 0);
    setDatasetRecord(dataset);
    templateRecord.current = restoredTemplate;
    currentTemplate.current = restoredTemplate?.file;
    setTemplateEntry(restoredTemplate);
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
        const manifest = await exampleCatalog.loadManifest();
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
      disposeDocumentActions();
      exampleLoadId.current += 1;
    };
  }, []);

  function openDiagnostics(tab: DiagnosticsTab) {
    setDiagnosticsTab(tab);
    changeProblemsOpen(true);
  }

  useWorkspaceShortcuts({ generateAction, jsonSettingsOpen, shortcutsOpen, problemsOpen, saveData, changeProblemsOpen, setShortcutsOpen });

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
      const { file, data } = await exampleCatalog.loadAssets(example);
      if (exampleLoadId.current !== loadId || contentRevision.current !== expectedRevision) return;
      addDataset({ name: filenameFromPath(example.dataPath), kind: "example", text: data, baseline: data, exampleId: example.id });
      const entry = { id: `template-${++nextId.current}`, name: file.name, file, exampleId: example.id };
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
    if (id === selectedExample && templateEntry?.exampleId === id && dataset?.kind === "example" && dataset.exampleId === id && dataset.text === dataset.baseline) {
      setActive("template");
      if (exampleLoading) {
        exampleLoadId.current += 1;
        setPendingExample(undefined);
        setExampleLoading(false);
      }
      return;
    }
    const expectedRevision = contentRevision.current;
      const record = datasetRecord.current;
    if (!await confirmDiscard(record, "Open example?", "Replace the current template and dataset with the selected example, discarding unsaved changes?", "Replace")) return;
    if (contentRevision.current !== expectedRevision) return;
    setPendingExample(id);
    void loadExample(example);
  }

  function changeJson(value: string) {
    if (value === currentJson.current) return;
    commitContent();
    currentJson.current = value;
    const updated = { ...datasetRecord.current, text: value };
    datasetRecord.current = updated;
    setDataset(updated);
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
        datasetRecord.current = { ...emptyDataset };
        setDatasetRecord({ ...emptyDataset });
        templateRecord.current = undefined;
        setTemplateEntry(undefined);
        currentTemplate.current = undefined;
        setSelectedExample(undefined);
        resetDocumentActions();
        setErrors([]);
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
            manifest = await exampleCatalog.loadManifest();
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

  function downloadBlob(blob: Blob, filename: string) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
  }

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
  const resultCurrent = Boolean(result && resultInputs?.datasetId === datasetId && resultInputs?.json === json && resultInputs?.template === template);

  return {
    storageWarning,
    dialogs: {
      contexts: { modalContext, feedbackContext },
      shortcuts: { shortcutsOpen, setShortcutsOpen, shortcutGroups },
      generateDataset: { jsonSettingsOpen, pending, submitJsonSettings, invalidateOperation, setJsonSettingsOpen, token, jsonSettings }
    },
    layout: { workspaceKey, problemsOpen, onProblemsOpenChange: changeProblemsOpen },
    header: { examples, pendingExample, selectedExample, chooseExample, generateShortcut, token, pending, dataset, selectedTemplate: templateEntry, generate, template, analyze, setJsonSettingsOpen, isMac, setShortcutsOpen, saving, clearing, clearLocalData },
    datasetPanel: { dataset, modified, token, recoverData, editor, json, changeJson, wordWrap, changeWordWrap, chooseDataFile, editorAvailable, saveData, openLabel, commandKey, saveAsShortcut, formatShortcut, saving, formatJson, copyJson, dataInput, reopenDataId, openData },
    documentPanel: { active, setActive, result, selectedTemplate: templateEntry, resultInputs, pending, exampleLoading, resultCurrent, dragging, setDragging, recoverTemplate, drop, files, openTemplate, openLabel, acceptTemplate, template, downloadBlob, clearTemplate, closeResult, fileInput, selected },
    diagnosticsPanel: { token, diagnosticsTab, setDiagnosticsTab, changeProblemsOpen, errors, setErrors, copyProblems, analysis },
    footer: { token, operationStatus, errors, openDiagnostics, problemsOpen, diagnosticsTab }
  };
}

export type Workspace = ReturnType<typeof useWorkspace>;
