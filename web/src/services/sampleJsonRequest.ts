import type { Diagnostic } from "../types";

export interface SampleJsonSettings {
  itemsPerCollection: number;
  values: "filled" | "empty";
}

export interface SampleJsonResponse {
  data: unknown;
  warnings: Diagnostic[];
}

export class SampleJsonFailure extends Error {
  constructor(public readonly diagnostics: Diagnostic[], public readonly response?: unknown) { super("Sample JSON generation failed"); }
}

export async function requestSampleJson(file: File, settings: SampleJsonSettings, signal: AbortSignal): Promise<SampleJsonResponse> {
  const body = new FormData();
  body.append("template", file);
  body.append("itemsPerCollection", String(settings.itemsPerCollection));
  body.append("values", settings.values);
  let response: Response;
  try { response = await fetch("/api/documents/sample", { method: "POST", body, signal }); }
  catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") throw error;
    throw new SampleJsonFailure([{ title: "Server unavailable", detail: "The sample service could not be reached." }]);
  }
  let payload: unknown;
  try { payload = await response.json(); } catch { payload = undefined; }
  if (!response.ok) throw new SampleJsonFailure(problemDiagnostics(payload), payload);
  if (!payload || typeof payload !== "object" || !("data" in payload) || !Array.isArray((payload as { warnings?: unknown }).warnings)) {
    throw new SampleJsonFailure([{ title: "Unexpected response", detail: "The server returned an invalid sample." }]);
  }
  const sample = payload as { data: unknown; warnings: unknown[] };
  const warnings = sample.warnings.map(parseDiagnostic);
  if (!warnings.every(Boolean)) throw new SampleJsonFailure([{ title: "Unexpected response", detail: "The server returned invalid sample warnings." }]);
  return { data: sample.data, warnings: warnings as Diagnostic[] };
}

function parseDiagnostic(value: unknown): Diagnostic | undefined {
  if (!value || typeof value !== "object") return undefined;
  const item = value as Record<string, unknown>;
  const detail = typeof item.detail === "string" ? item.detail : typeof item.message === "string" ? item.message : undefined;
  if (!detail) return undefined;
  return {
    title: typeof item.title === "string" ? item.title : typeof item.code === "string" ? item.code : "Sample warning",
    detail,
    code: typeof item.code === "string" ? item.code : undefined,
    field: typeof item.field === "string" ? item.field : undefined,
    part: typeof item.part === "string" ? item.part : undefined,
    location: typeof item.location === "string" ? item.location : undefined,
    formatter: typeof item.formatter === "string" ? item.formatter : undefined,
    dataPath: typeof item.dataPath === "string" ? item.dataPath : undefined,
    limitReached: typeof item.limitReached === "boolean" ? item.limitReached : undefined
  };
}

function problemDiagnostics(value: unknown): Diagnostic[] {
  if (!value || typeof value !== "object") return [{ title: "Sample generation failed", detail: "The sample could not be generated." }];
  const item = value as Record<string, unknown>;
  const detail = typeof item.detail === "string" ? item.detail : "The sample could not be generated.";
  const extensions = item.extensions && typeof item.extensions === "object"
    ? item.extensions as Record<string, unknown>
    : item;
  const source = Array.isArray(extensions.diagnostics) ? extensions.diagnostics : Array.isArray(extensions.errors) ? extensions.errors : [];
  const diagnostics = source.map(parseDiagnostic).filter((entry): entry is Diagnostic => Boolean(entry));
  const code = typeof extensions.code === "string" ? extensions.code : undefined;
  const duplicateIndex = code ? diagnostics.findIndex(diagnostic => diagnostic.code === code) : -1;
  if (duplicateIndex >= 0) {
    diagnostics[duplicateIndex] = {
      ...diagnostics[duplicateIndex],
      field: diagnostics[duplicateIndex].field ?? (typeof extensions.field === "string" ? extensions.field : undefined),
      traceId: diagnostics[duplicateIndex].traceId ?? (typeof extensions.traceId === "string" ? extensions.traceId : undefined),
      limitReached: diagnostics[duplicateIndex].limitReached ?? (typeof extensions.limitReached === "boolean" ? extensions.limitReached : undefined)
    };
    return diagnostics;
  }
  return [{ title: typeof item.title === "string" ? item.title : "Sample generation failed", detail,
    code, field: typeof extensions.field === "string" ? extensions.field : undefined,
    traceId: typeof extensions.traceId === "string" ? extensions.traceId : undefined,
    limitReached: typeof extensions.limitReached === "boolean" ? extensions.limitReached : undefined }, ...diagnostics];
}
