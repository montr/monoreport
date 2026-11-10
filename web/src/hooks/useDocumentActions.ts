import { Form } from "antd";
import { useRef, useState, type Dispatch, type RefObject, type SetStateAction } from "react";
import { analyzeTemplate, AnalysisFailure } from "../services/analysis";
import { generateDocument, GenerationFailure } from "../services/generation";
import { requestSampleJson, SampleJsonFailure, type SampleJsonSettings } from "../services/sampleJsonRequest";
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
  selectedTemplate: TemplateEntry | undefined;
  currentTemplate: RefObject<File | undefined>;
  contentRevision: RefObject<number>;
  datasetRecord: RefObject<Dataset>;
  setErrors: Dispatch<SetStateAction<Problem[]>>;
  openDiagnostics(tab: DiagnosticsTab): void;
  setActive: Dispatch<SetStateAction<Tab>>;
  addDataset(record: Omit<Dataset, "id">): void;
  onSettingsRevision: RefObject<number>;
  refreshTemplate(owned: boolean, isCurrent: () => boolean): Promise<void>;
  reconcileInputs(): Promise<void>;
}

export function useDocumentActions({ template, selectedTemplate, currentTemplate, contentRevision, datasetRecord, setErrors, openDiagnostics, setActive, addDataset, onSettingsRevision, refreshTemplate, reconcileInputs }: Props) {
  const [result, setResult] = useState<GeneratedDocument>();
  const [resultInputs, setResultInputs] = useState<{ datasetId: string | undefined; json: string; template: File; templateSource: "Example file" | "Local file" }>();
  const [analysis, setAnalysis] = useState<TemplateAnalysis>();
  const [analysisWireResult, setAnalysisWireResult] = useState<unknown>();
  const [sampleCandidate, setSampleCandidate] = useState<{ json: string; warnings: Diagnostic[]; template: File; datasetId: string | undefined; revision: number; settingsRevision: number; requestId: number }>();
  const [sampleFailure, setSampleFailure] = useState<Diagnostic[]>();
  const [jsonSettingsOpen, setJsonSettingsOpenState] = useState(false);
  const dialogSessionRevision = useRef(0);
  const [jsonSettings] = Form.useForm<SampleJsonSettings>();
  const { operationStatus, setOperationStatus, runningOperation, pending, setPending, operationId, request, startOperation, finishOperation, invalidateOperation } = useDocumentActionLifecycle();

  const automaticOperation = useRef(false);
  const lastAttempt = useRef<{ revision: number; readFailure?: boolean } | undefined>(undefined);

  function invalidateGeneration() {
    if (runningOperation.current?.action === "Generate") invalidateSample();
  }

  function cancelAutomatic() {
    if (automaticOperation.current && runningOperation.current?.action === "Generate") invalidateSample();
  }

  function retryAutomatic(readRecovered = false) {
    if (readRecovered) {
      if (lastAttempt.current?.readFailure) lastAttempt.current = undefined;
      return;
    }
    if (runningOperation.current?.action === "Generate") return;
    const data = datasetRecord.current;
    const matches = result && resultInputs?.template === currentTemplate.current && resultInputs?.datasetId === data.id && resultInputs.json === data.text;
    lastAttempt.current = matches ? { revision: contentRevision.current } : undefined;
  }

  function setJsonSettingsOpen(open: boolean) {
    dialogSessionRevision.current += 1;
    setJsonSettingsOpenState(open);
  }

  function captureSampleValidation() {
    const expectedOperation = operationId.current;
    const expectedContentRevision = contentRevision.current;
    const expectedSettingsRevision = onSettingsRevision.current;
    const expectedTemplate = currentTemplate.current;
    const expectedDatasetId = datasetRecord.current.id;
    const expectedDialogSession = dialogSessionRevision.current;
    return () => operationId.current === expectedOperation &&
      contentRevision.current === expectedContentRevision &&
      onSettingsRevision.current === expectedSettingsRevision &&
      currentTemplate.current === expectedTemplate &&
      datasetRecord.current.id === expectedDatasetId &&
      dialogSessionRevision.current === expectedDialogSession;
  }

  async function generate(automatic = false) {
    if (automatic && lastAttempt.current?.revision === contentRevision.current) return "skipped" as const;
    if (pending || runningOperation.current || jsonSettingsOpen) return "skipped" as const;
    invalidateOperation();
    automaticOperation.current = automatic;
    startOperation("Generate");
    if (!template) {
      setErrors(diagnostic("Template is required", "Select a DOCX or XLSX file before generating a document.").map(item => ({ ...item, action: "Generate" })));
      finishOperation("error");
      openDiagnostics("problems");
      return "error" as const;
    }
    const operation = operationId.current;
    let expectedRevision = contentRevision.current;
    lastAttempt.current = { revision: expectedRevision, readFailure: true };
    const controller = new AbortController();
    request.current = controller;
    setErrors([]);
    setPending("generate");
    try {
      await refreshTemplate(true, () => operationId.current === operation);
      if (operationId.current !== operation) return "obsolete" as const;
      expectedRevision = contentRevision.current;
      const source = currentTemplate.current;
      const data = { ...datasetRecord.current };
      if (!source) return "obsolete" as const;
      lastAttempt.current = { revision: expectedRevision };
      const generated = await generateDocument(source, data.text, controller.signal);
      await reconcileInputs();
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return "obsolete" as const;
      setResult({ ...generated, filename: resultFilename(source.name) });
      const warnings = (generated.warnings ?? []).map(item => ({ ...item, action: "Generate" as const, severity: "Warning" as const }));
      setErrors(warnings);
      setResultInputs({ datasetId: data.id, json: data.text, template: source, templateSource: selectedTemplate?.exampleId ? "Example file" : "Local file" });
      if (!automatic || !result) setActive("result");
      finishOperation("success");
      if (warnings.length) openDiagnostics("problems");
      return "success" as const;
    } catch (error) {
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return "obsolete" as const;
      try { await reconcileInputs(); } catch { /* Source feedback is owned by the read path. */ }
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return "obsolete" as const;
      if (error instanceof DOMException && error.name === "AbortError") {
        lastAttempt.current = undefined;
        finishOperation("cancelled");
        return "cancelled" as const;
      }
      setErrors((error instanceof GenerationFailure ? error.diagnostics : diagnostic("Generation failed", "The document could not be generated.")).map(item => ({ ...item, action: "Generate" })));
      finishOperation("error");
      openDiagnostics("problems");
      return "error" as const;
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

  async function analyze() {
    if (pending || runningOperation.current || !template || jsonSettingsOpen) return;
    invalidateOperation();
    const action: OperationAction = "Analyze";
    startOperation(action);
    const operation = operationId.current;
    const expectedRevision = contentRevision.current;
    const source = template;
    const controller = new AbortController();
    request.current = controller;
    setErrors([]);
    setPending("analyze");
    openDiagnostics("analysis");
    try {
      const analyzed = await analyzeTemplate(source, controller.signal);
      if (operationId.current !== operation || contentRevision.current !== expectedRevision || currentTemplate.current !== source) return;
      setAnalysis(analyzed.analysis);
      setAnalysisWireResult(analyzed.wireResult);
      setErrors(analysisProblems(analyzed.analysis, action));
      finishOperation("success");
    } catch (error) {
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return;
      if (error instanceof DOMException && error.name === "AbortError") { finishOperation("cancelled"); return; }
      if (error instanceof AnalysisFailure && error.analysis) {
        setAnalysis(error.analysis);
        setAnalysisWireResult(error.wireResult);
        const problems = analysisProblems(error.analysis, action);
        setErrors(error.analysis.errors.length ? problems : [...problems, ...error.diagnostics.map(item => ({ ...item, action }))]);
      } else {
        const problems = error instanceof AnalysisFailure ? error.diagnostics : diagnostic(`${action} failed`, error instanceof Error ? error.message : "The operation could not be completed.");
        setErrors(problems.map(item => ({ ...item, action })));
      }
      finishOperation("error");
      openDiagnostics("problems");
    } finally {
      if (operationId.current === operation) {
        request.current = undefined;
        setPending(undefined);
      }
    }
  }

  async function submitJsonSettings(settings: SampleJsonSettings) {
    if (pending || runningOperation.current || !template) return;
    invalidateOperation();
    const started = performance.now();
    startOperation("Generate Dataset", started);
    const operation = operationId.current;
    const expectedRevision = contentRevision.current;
    const source = template;
    const selectedDatasetId = datasetRecord.current.id;
    const settingsRevision = onSettingsRevision.current;
    const controller = new AbortController();
    request.current = controller;
    setSampleCandidate(undefined);
    setSampleFailure(undefined);
    setErrors([]);
    setPending("generate-json");
    try {
      const response = await requestSampleJson(source, settings, controller.signal);
      if (operationId.current !== operation || contentRevision.current !== expectedRevision || currentTemplate.current !== source || datasetRecord.current.id !== selectedDatasetId || onSettingsRevision.current !== settingsRevision) return;
      const serialized = JSON.stringify(response.data, null, 2);
      setSampleCandidate({ json: serialized, warnings: response.warnings, template: source, datasetId: selectedDatasetId, revision: expectedRevision, settingsRevision, requestId: operation });
      setSampleFailure(undefined);
      setErrors(response.warnings.map(item => ({ ...item, action: "Generate Dataset", severity: "Warning" })));
      finishOperation("success");
    } catch (error) {
      if (operationId.current !== operation || contentRevision.current !== expectedRevision) return;
      if (error instanceof DOMException && error.name === "AbortError") { finishOperation("cancelled"); return; }
      const problems = error instanceof SampleJsonFailure ? error.diagnostics : diagnostic("Sample generation failed", "The sample could not be generated.");
      setSampleFailure(problems);
      setErrors(problems.map(item => ({ ...item, action: "Generate Dataset" })));
      finishOperation("error");
      openDiagnostics("problems");
    } finally {
      if (operationId.current === operation) { request.current = undefined; setPending(undefined); }
    }
  }

  function acceptSample() {
    const candidate = sampleCandidate;
    if (!candidate || operationId.current !== candidate.requestId || contentRevision.current !== candidate.revision || currentTemplate.current !== candidate.template || datasetRecord.current.id !== candidate.datasetId || onSettingsRevision.current !== candidate.settingsRevision) return;
    addDataset({ name: generatedDatasetFilename(candidate.template.name), kind: "generated", text: candidate.json });
    setErrors(current => current.filter(item => item.action !== "Generate Dataset"));
    setSampleCandidate(undefined);
    setSampleFailure(undefined);
    setJsonSettingsOpen(false);
  }

  function invalidateSample(preserveOperation = false) {
    setSampleCandidate(undefined);
    setSampleFailure(undefined);
    setErrors(current => current.filter(item => item.action !== "Generate Dataset"));
    if (!preserveOperation) {
      if (runningOperation.current?.action === "Generate") lastAttempt.current = undefined;
      invalidateOperation();
    }
  }

  function clearAnalysis() {
    setAnalysis(undefined);
    setAnalysisWireResult(undefined);
  }

  function closeResult() {
    cancelAutomatic();
    lastAttempt.current = { revision: contentRevision.current };
    setResult(undefined);
    setResultInputs(undefined);
    if (template) setActive("template");
  }

  function resetDocumentActions() {
    setResult(undefined);
    setResultInputs(undefined);
    setSampleCandidate(undefined);
    setSampleFailure(undefined);
    clearAnalysis();
    runningOperation.current = undefined;
    setOperationStatus({ outcome: "ready" });
    setJsonSettingsOpen(false);
  }

  function disposeDocumentActions() {
    operationId.current += 1;
    request.current?.abort();
  }

  return { invalidateGeneration, cancelAutomatic, retryAutomatic, result, resultInputs, analysis, analysisWireResult, sampleCandidate, sampleFailure, acceptSample, clearAnalysis, closeResult, resetDocumentActions, disposeDocumentActions, jsonSettingsOpen, setJsonSettingsOpen, captureSampleValidation, jsonSettings, operationStatus, pending, invalidateOperation: invalidateSample, generate, analyze, submitJsonSettings };
}
