import { Form } from "antd";
import { useRef, useState, type Dispatch, type RefObject, type SetStateAction } from "react";
import { analyzeTemplate, AnalysisFailure } from "../services/analysis";
import { generateDocument, GenerationFailure } from "../services/generation";
import { SampleJson, type SampleJsonSettings } from "../services/sampleJson";
import type { Diagnostic, GeneratedDocument, TemplateAnalysis, Dataset, DiagnosticsTab, Problem, Tab, TemplateEntry } from "../types";

type Operation = "generate" | "analyze" | "generate-json";
export type OperationAction = "Generate" | "Analyze" | "Generate Dataset";
type OperationOutcome = "running" | "success" | "error" | "cancelled" | "ready";
interface OperationStatus { action?: OperationAction; outcome: OperationOutcome; started?: number; duration?: number; }

function useDocumentActionLifecycle() {
  const [operationStatus, setOperationStatus] = useState<OperationStatus>({ outcome: "ready" });
  const runningOperation = useRef<OperationStatus | undefined>(undefined);
  const [pending, setPending] = useState<Operation>();
  const operationId = useRef(0);
  const request = useRef<AbortController | undefined>(undefined);

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

  function invalidateOperation() {
    finishOperation("cancelled");
    operationId.current += 1;
    request.current?.abort();
    request.current = undefined;
    setPending(undefined);
  }

  return { operationStatus, setOperationStatus, runningOperation, pending, setPending, operationId, request, startOperation, finishOperation, invalidateOperation };
}

function diagnostic(title: string, detail: string): Diagnostic[] {
  return [{ title, detail }];
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

interface Props {
  template: File | undefined;
  datasetId: string | undefined;
  json: string;
  selectedTemplate: TemplateEntry | undefined;
  currentTemplate: RefObject<File | undefined>;
  contentRevision: RefObject<number>;
  datasetRecord: RefObject<Dataset>;
  setErrors: Dispatch<SetStateAction<Problem[]>>;
  openDiagnostics(tab: DiagnosticsTab): void;
  setActive: Dispatch<SetStateAction<Tab>>;
  addDataset(record: Omit<Dataset, "id">): void;
  confirmDiscard(record: Dataset | undefined, title: string, content: string, okText: string): Promise<boolean>;
}

export function useDocumentActions({ template, datasetId, json, selectedTemplate, currentTemplate, contentRevision, datasetRecord, setErrors, openDiagnostics, setActive, addDataset, confirmDiscard }: Props) {
  const [result, setResult] = useState<GeneratedDocument>();
  const [resultInputs, setResultInputs] = useState<{ datasetId: string | undefined; json: string; template: File; templateSource: "Example file" | "Local file" }>();
  const [analysis, setAnalysis] = useState<TemplateAnalysis>();
  const analysisCache = useRef<{ source: File; analysis: TemplateAnalysis } | undefined>(undefined);
  const [sampleJson] = useState(() => new SampleJson(Math.random));
  const [jsonSettingsOpen, setJsonSettingsOpen] = useState(false);
  const [jsonSettings] = Form.useForm<SampleJsonSettings>();
  const { operationStatus, setOperationStatus, runningOperation, pending, setPending, operationId, request, startOperation, finishOperation, invalidateOperation } = useDocumentActionLifecycle();

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
      if (!await confirmDiscard(datasetRecord.current, "Generate dataset?", "Replace the current dataset, discarding unsaved changes?", "Replace")) return;
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return;
      await analyze(settings);
    } catch {
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return;
      startOperation("Generate Dataset", started);
      setErrors(diagnostic("Invalid settings", "Items per collection must be a positive integer.").map(item => ({ ...item, action: "Generate Dataset" })));
      finishOperation("error");
    }
  }

  function clearAnalysis() {
    analysisCache.current = undefined;
    setAnalysis(undefined);
  }

  function closeResult() {
    setResult(undefined);
    setResultInputs(undefined);
    if (template) setActive("template");
  }

  function resetDocumentActions() {
    setResult(undefined);
    setResultInputs(undefined);
    clearAnalysis();
    runningOperation.current = undefined;
    setOperationStatus({ outcome: "ready" });
    setJsonSettingsOpen(false);
  }

  function disposeDocumentActions() {
    operationId.current += 1;
    request.current?.abort();
  }

  return { result, resultInputs, analysis, clearAnalysis, closeResult, resetDocumentActions, disposeDocumentActions, jsonSettingsOpen, setJsonSettingsOpen, jsonSettings, operationStatus, pending, invalidateOperation, generate, analyze, submitJsonSettings };
}
