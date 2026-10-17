import type { Diagnostic, TemplateAnalysis, TemplateOccurrence, TemplateRequirement } from "../types";
import { normalizeProblem } from "./generation";

export class AnalysisFailure extends Error {
  constructor(public readonly diagnostics: Diagnostic[]) {
    super("Template analysis failed");
  }
}

function text(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function diagnostic(value: unknown): Diagnostic | undefined {
  if (!value || typeof value !== "object") return undefined;
  const item = value as Record<string, unknown>;
  const detail = text(item.message) ?? text(item.detail);
  if (!detail) return undefined;
  return {
    title: text(item.title) ?? text(item.code) ?? "Template issue",
    detail,
    code: text(item.code),
    field: text(item.field),
    part: text(item.part),
    location: text(item.location),
    formatter: text(item.formatter),
    dataPath: text(item.dataPath)
  };
}

function diagnostics(value: unknown): Diagnostic[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const parsed = value.map(diagnostic);
  return parsed.every(Boolean) ? parsed as Diagnostic[] : undefined;
}

function requirement(value: unknown): TemplateRequirement | undefined {
  if (!value || typeof value !== "object") return undefined;
  const item = value as Record<string, unknown>;
  if (typeof item.name !== "string" || !item.name ||
      (item.type !== "object" && item.type !== "array" && item.type !== "value" && item.type !== "boolean")) return undefined;
  if (item.fields === undefined) return { name: item.name, type: item.type };
  if (!Array.isArray(item.fields) || item.fields.length === 0) return undefined;
  const children = item.fields.map(requirement);
  if (!children.every(Boolean)) return undefined;
  return { name: item.name, type: item.type, fields: children as TemplateRequirement[] };
}

function occurrence(value: unknown): TemplateOccurrence | undefined {
  if (!value || typeof value !== "object") return undefined;
  const item = value as Record<string, unknown>;
  if ((item.kind !== "field" && item.kind !== "for" && item.kind !== "if") ||
      typeof item.expression !== "string" || typeof item.path !== "string" ||
      typeof item.part !== "string" || typeof item.location !== "string") return undefined;
  return { kind: item.kind, expression: item.expression, path: item.path, part: item.part, location: item.location };
}

function parseAnalysis(value: unknown): TemplateAnalysis | undefined {
  if (!value || typeof value !== "object") return undefined;
  const item = value as Record<string, unknown>;
  const errors = diagnostics(item.errors);
  const warnings = diagnostics(item.warnings);
  if (typeof item.valid !== "boolean" || !errors || !warnings ||
      !Array.isArray(item.requirements) || !Array.isArray(item.occurrences)) return undefined;
  const requirements = item.requirements.map(requirement);
  const occurrences = item.occurrences.map(occurrence);
  if (!requirements.every(Boolean) || !occurrences.every(Boolean)) return undefined;
  return {
    valid: item.valid,
    errors,
    warnings,
    requirements: requirements as TemplateRequirement[],
    occurrences: occurrences as TemplateOccurrence[]
  };
}

async function responseJson(response: Response): Promise<unknown> {
  try { return await response.json(); } catch { return undefined; }
}

export async function analyzeTemplate(file: File, signal?: AbortSignal): Promise<TemplateAnalysis> {
  const body = new FormData();
  body.append("template", file);
  let response: Response;
  try {
    response = await fetch("/api/documents/analyze", { method: "POST", body, signal });
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") throw error;
    throw new AnalysisFailure([{ title: "Server unavailable", detail: "The analysis service could not be reached." }]);
  }
  const payload = await responseJson(response);
  if (!response.ok) throw new AnalysisFailure(normalizeProblem(payload, response.status));
  const analysis = parseAnalysis(payload);
  if (!analysis) throw new AnalysisFailure([{ title: "Unexpected response", detail: "The server returned an invalid template analysis." }]);
  if (!analysis.valid) {
    throw new AnalysisFailure(analysis.errors.length ? analysis.errors : [{ title: "Invalid template", detail: "The template could not be analyzed safely." }]);
  }
  return analysis;
}
