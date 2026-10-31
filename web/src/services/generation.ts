import type { Diagnostic, GeneratedDocument } from "../types";
import { documentFallbackFilename, documentFormatFromMime } from "./documentFormat";

export class GenerationFailure extends Error {
  constructor(public readonly diagnostics: Diagnostic[]) {
    super("Document generation failed");
  }
}

function text(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function issue(value: unknown, fallbackTitle: string): Diagnostic | undefined {
  if (!value || typeof value !== "object") return undefined;
  const item = value as Record<string, unknown>;
  const detail = text(item.detail) ?? text(item.message);
  if (!detail) return undefined;
  const part = text(item.part);
  const location = text(item.location);
  const formatter = text(item.formatter);
  const dataPath = text(item.dataPath);
  return {
    title: text(item.title) ?? text(item.code) ?? fallbackTitle,
    detail,
    code: text(item.code),
    field: text(item.field),
    ...(part ? { part } : {}),
    ...(location ? { location } : {}),
    ...(formatter ? { formatter } : {}),
    ...(dataPath ? { dataPath } : {})
  };
}

function collection(value: unknown, fallbackTitle: string): Diagnostic[] {
  if (Array.isArray(value)) return value.map(item => issue(item, fallbackTitle)).filter((item): item is Diagnostic => Boolean(item));
  if (!value || typeof value !== "object") return [];
  return Object.entries(value as Record<string, unknown>).flatMap(([field, entries]) => {
    const values = Array.isArray(entries) ? entries : [entries];
    return values.flatMap(entry => {
      if (typeof entry === "string" && entry.trim()) return [{ title: fallbackTitle, detail: entry.trim(), field }];
      const parsed = issue(entry, fallbackTitle);
      return parsed ? [{ ...parsed, field: parsed.field ?? field }] : [];
    });
  });
}

export function normalizeProblem(value: unknown, status?: number): Diagnostic[] {
  if (!value || typeof value !== "object") return [{ title: "Request failed", detail: status ? `The server returned HTTP ${status}.` : "The server returned an unexpected response." }];
  const problem = value as Record<string, unknown>;
  const issues = collection(problem.errors, "Validation error");
  if (issues.length) return issues;
  const parsed = issue(problem, "Request failed");
  return parsed ? [parsed] : [{ title: "Request failed", detail: status ? `The server returned HTTP ${status}.` : "The server returned an unexpected response." }];
}

async function failureDiagnostics(response: Response): Promise<Diagnostic[]> {
  try {
    return normalizeProblem(await response.json(), response.status);
  } catch {
    return normalizeProblem(undefined, response.status);
  }
}

function safeFilename(disposition: string | null, format: GeneratedDocument["format"]): string {
  const fallback = documentFallbackFilename(format);
  if (!disposition) return fallback;
  const encoded = disposition.match(/filename\*=UTF-8''([^;]+)/i)?.[1];
  const quoted = disposition.match(/filename="([^"]+)"/i)?.[1];
  const plain = disposition.match(/filename=([^;]+)/i)?.[1]?.trim();
  let candidate = encoded ?? quoted ?? plain;
  if (encoded) {
    try { candidate = decodeURIComponent(encoded); } catch { return fallback; }
  }
  if (!candidate || /[\\/\x00-\x1f\x7f]/.test(candidate)) return fallback;
  candidate = candidate.trim();
  return candidate.toLowerCase().endsWith(`.${format}`) && candidate.length <= 180 ? candidate : fallback;
}

export async function generateDocument(file: File, json: string, signal?: AbortSignal): Promise<GeneratedDocument> {
  const body = new FormData();
  body.append("template", file);
  body.append("data", json);
  let response: Response;
  try {
    response = await fetch("/api/documents/generate", { method: "POST", body, signal });
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") throw error;
    throw new GenerationFailure([{ title: "Server unavailable", detail: "The generation service could not be reached." }]);
  }
  if (!response.ok) throw new GenerationFailure(await failureDiagnostics(response));
  const blob = await response.blob();
  const mediaType = response.headers.get("Content-Type") || blob.type;
  const format = documentFormatFromMime(mediaType);
  if (!format || blob.size === 0) {
    throw new GenerationFailure([{ title: "Unexpected response", detail: "The server did not return a supported DOCX or XLSX document." }]);
  }
  return { blob, filename: safeFilename(response.headers.get("Content-Disposition"), format), format };
}
