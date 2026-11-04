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
  const fieldCode = text(item.fieldCode);
  const traceId = text(item.traceId);
  const removedCharacterCount = typeof item.removedCharacterCount === "number" && Number.isInteger(item.removedCharacterCount) && item.removedCharacterCount >= 0
    ? item.removedCharacterCount : undefined;
  return {
    title: text(item.title) ?? text(item.code) ?? fallbackTitle,
    detail,
    code: text(item.code),
    field: text(item.field),
    ...(part ? { part } : {}),
    ...(location ? { location } : {}),
    ...(formatter ? { formatter } : {}),
    ...(dataPath ? { dataPath } : {}),
    ...(fieldCode ? { fieldCode } : {}),
    ...(removedCharacterCount === undefined ? {} : { removedCharacterCount }),
    ...(traceId ? { traceId } : {}),
    ...(typeof item.limitReached === "boolean" ? { limitReached: item.limitReached } : {})
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
  if (issues.length) return addProblemContext(issues, problem);
  const parsed = issue(problem, "Request failed");
  return parsed ? [parsed] : addProblemContext([{ title: "Request failed", detail: status ? `The server returned HTTP ${status}.` : "The server returned an unexpected response." }], problem);
}

function addProblemContext(issues: Diagnostic[], problem: Record<string, unknown>): Diagnostic[] {
  const traceId = text(problem.traceId);
  const limitReached = typeof problem.limitReached === "boolean" ? problem.limitReached : undefined;
  return issues.map(item => ({ ...item, ...(traceId ? { traceId } : {}), ...(limitReached === undefined ? {} : { limitReached }) }));
}

async function failureDiagnostics(response: Response): Promise<Diagnostic[]> {
  try {
    return normalizeProblem(await response.json(), response.status);
  } catch {
    return normalizeProblem(undefined, response.status);
  }
}

function bytes(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function findBytes(source: Uint8Array, target: Uint8Array, start: number): number {
  for (let index = start; index <= source.length - target.length; index++) {
    let matched = true;
    for (let offset = 0; offset < target.length; offset++)
      if (source[index + offset] !== target[offset]) { matched = false; break; }
    if (matched) return index;
  }
  return -1;
}

function parseMultipart(contentType: string, body: Uint8Array): { document: Uint8Array; documentType: string; filename: string; diagnostics: unknown } {
  const match = /^multipart\/mixed\s*;\s*boundary=(?:"([A-Za-z0-9_-]{1,70})"|([A-Za-z0-9_-]{1,70}))\s*$/i.exec(contentType);
  const boundary = match?.[1] ?? match?.[2];
  if (!boundary || match?.[0] !== contentType) throw new Error("Invalid multipart boundary.");
  const marker = bytes(`--${boundary}`);
  let cursor = 0;
  const parts: { headers: string; content: Uint8Array }[] = [];
  for (;;) {
    if (findBytes(body, marker, cursor) !== cursor) throw new Error("Invalid multipart framing.");
    cursor += marker.length;
    if (body[cursor] === 45 && body[cursor + 1] === 45) {
      cursor += 2;
      if (body[cursor] === 13 && body[cursor + 1] === 10) cursor += 2;
      if (cursor !== body.length || parts.length !== 2) throw new Error("Unexpected multipart parts.");
      break;
    }
    if (body[cursor] !== 13 || body[cursor + 1] !== 10) throw new Error("Invalid multipart boundary line.");
    cursor += 2;
    const headerEnd = findBytes(body, bytes("\r\n\r\n"), cursor);
    if (headerEnd < 0) throw new Error("Missing multipart headers.");
    const headerBytes = body.subarray(cursor, headerEnd);
    if (headerBytes.some(byte => byte > 0x7f)) throw new Error("Invalid multipart headers.");
    const headers = new TextDecoder("utf-8", { fatal: true }).decode(headerBytes);
    if (headers.split("\r\n").length !== 2) throw new Error("Unexpected multipart headers.");
    const nextBoundary = findBytes(body, bytes(`\r\n--${boundary}`), headerEnd + 4);
    if (nextBoundary < 0) throw new Error("Missing multipart closing boundary.");
    parts.push({ headers, content: body.slice(headerEnd + 4, nextBoundary) });
    cursor = nextBoundary + 2;
    if (parts.length > 2) throw new Error("Unexpected multipart parts.");
  }

  const documentType = parts[0].headers.match(/^Content-Type:\s*([^\r\n]+)$/im)?.[1]?.trim();
  const diagnosticsType = parts[1].headers.match(/^Content-Type:\s*([^\r\n]+)$/im)?.[1]?.trim();
  if (!documentType || documentType.includes(";") || diagnosticsType?.toLowerCase() !== "application/json")
    throw new Error("Unexpected multipart part types.");
  const documentDisposition = parts[0].headers.match(/^Content-Disposition:\s*attachment;\s*filename="([^"]+)"\s*$/im)?.[1];
  const diagnosticsDisposition = parts[1].headers.match(/^Content-Disposition:\s*attachment;\s*filename="diagnostics\.json"\s*$/im)?.[0];
  if (!documentDisposition || !diagnosticsDisposition) throw new Error("Unexpected multipart attachments.");
  let diagnostics: unknown;
  try { diagnostics = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(parts[1].content)); }
  catch { throw new Error("Invalid diagnostics JSON."); }
  return { document: parts[0].content, documentType, filename: documentDisposition, diagnostics };
}

function multipartWarnings(value: unknown, format: GeneratedDocument["format"], filename: string): Diagnostic[] {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("Invalid diagnostics metadata.");
  const metadata = value as Record<string, unknown>;
  if (metadata.format !== format || metadata.fileName !== filename || !Array.isArray(metadata.warnings))
    throw new Error("Invalid diagnostics metadata.");
  return metadata.warnings.map(item => {
    if (!item || typeof item !== "object" || Array.isArray(item)) throw new Error("Invalid warning metadata.");
    const warning = item as Record<string, unknown>;
    const message = text(warning.message);
    const field = text(warning.field);
    if (!message || !field) throw new Error("Invalid warning metadata.");
    return {
      title: text(warning.code) ?? "Generation warning",
      detail: message,
      ...(text(warning.code) ? { code: text(warning.code) } : {}),
      field,
      ...(text(warning.fieldCode) ? { fieldCode: text(warning.fieldCode) } : {}),
      ...(text(warning.part) ? { part: text(warning.part) } : {}),
      ...(text(warning.location) ? { location: text(warning.location) } : {}),
      ...(text(warning.dataPath) ? { dataPath: text(warning.dataPath) } : {}),
      ...(typeof warning.removedCharacterCount === "number" && Number.isInteger(warning.removedCharacterCount) && warning.removedCharacterCount > 0
        ? { removedCharacterCount: warning.removedCharacterCount } : {})
    };
  });
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
    response = await fetch("/api/documents/generate", { method: "POST", body, signal, headers: { Accept: "multipart/mixed" } });
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") throw error;
    throw new GenerationFailure([{ title: "Server unavailable", detail: "The generation service could not be reached." }]);
  }
  if (!response.ok) throw new GenerationFailure(await failureDiagnostics(response));
  try {
    const contentType = response.headers.get("Content-Type") ?? "";
    const parsed = parseMultipart(contentType, new Uint8Array(await response.arrayBuffer()));
    const format = documentFormatFromMime(parsed.documentType);
    const expectedFilename = format ? documentFallbackFilename(format) : "";
    if (!format || parsed.document.length === 0) throw new Error("Unsupported document attachment.");
    const filename = safeFilename(`attachment; filename="${parsed.filename}"`, format);
    if (filename !== parsed.filename || filename !== expectedFilename) throw new Error("Invalid document filename.");
    const warnings = multipartWarnings(parsed.diagnostics, format, filename);
    const exactDocumentBytes = new Uint8Array(parsed.document.length);
    exactDocumentBytes.set(parsed.document);
    return { blob: new Blob([exactDocumentBytes.buffer], { type: parsed.documentType }), filename, format, warnings };
  } catch {
    throw new GenerationFailure([{ title: "Unexpected response", detail: "The server returned an invalid document and diagnostics response." }]);
  }
}
