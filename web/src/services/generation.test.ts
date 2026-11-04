import { describe, expect, it, vi } from "vitest";
import { generateDocument, GenerationFailure, normalizeProblem } from "./generation";

const mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const xlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

async function success(blob: Blob, contentType: string, filename = contentType === xlsxMime ? "document.xlsx" : "document.docx", warnings: unknown[] = []) {
  const boundary = "test-boundary";
  const diagnostics = new TextEncoder().encode(JSON.stringify({ format: contentType === xlsxMime ? "xlsx" : "docx", fileName: filename, warnings }));
  const encoder = new TextEncoder();
  const documentHeader = encoder.encode(`--${boundary}\r\nContent-Type: ${contentType}\r\nContent-Disposition: attachment; filename="${filename}"\r\n\r\n`);
  const diagnosticHeader = encoder.encode(`\r\n--${boundary}\r\nContent-Type: application/json\r\nContent-Disposition: attachment; filename="diagnostics.json"\r\n\r\n`);
  const closing = encoder.encode(`\r\n--${boundary}--\r\n`);
  const all = new Uint8Array(documentHeader.length + blob.size + diagnosticHeader.length + diagnostics.length + closing.length);
  let offset = 0;
  for (const part of [documentHeader, new Uint8Array(await blob.arrayBuffer()), diagnosticHeader, diagnostics, closing]) {
    all.set(part, offset);
    offset += part.length;
  }
  return new Response(all.buffer, { status: 200, headers: { "Content-Type": `multipart/mixed; boundary="${boundary}"` } });
}

describe("generation client", () => {
  it("sends the exact multipart contract and returns actual DOCX bytes with a safe filename", async () => {
    const bytes = new Uint8Array([80, 75, 3, 4]);
    const responseBlob = new Blob([bytes], { type: mime });
    const fetchMock = vi.fn().mockResolvedValue(success(responseBlob, mime));
    vi.stubGlobal("fetch", fetchMock);
    const template = new File([bytes], "template.docx", { type: mime });

    const result = await generateDocument(template, "{\"value\":1}");

    expect(result.filename).toBe("document.docx");
    expect(result.format).toBe("docx");
    expect(new Uint8Array(await result.blob.arrayBuffer())).toEqual(new Uint8Array(await responseBlob.arrayBuffer()));
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe("/api/documents/generate");
    expect(init.method).toBe("POST");
    expect(init.headers).toEqual({ Accept: "multipart/mixed" });
    const body = init.body as FormData;
    expect(body.get("template")).toBe(template);
    expect(body.get("data")).toBe("{\"value\":1}");
    expect(body.has("culture")).toBe(false);
  });

  it("selects XLSX from the response MIME despite a DOCX upload and preserves its exact bytes", async () => {
    const responseBlob = new Blob([new Uint8Array([80, 75, 7, 8])], { type: xlsxMime });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(success(responseBlob, xlsxMime)));

    const result = await generateDocument(new File([new Uint8Array([1])], "template.docx", { type: mime }), "{}");

    expect(result.filename).toBe("document.xlsx");
    expect(result.format).toBe("xlsx");
    expect(new Uint8Array(await result.blob.arrayBuffer())).toEqual(new Uint8Array(await responseBlob.arrayBuffer()));
    expect(result.warnings).toEqual([]);
  });

  it("preserves API Problem Details and validation-style collections without inventing issues", () => {
    expect(normalizeProblem({ title: "Document generation failed", detail: "Data is incompatible.", code: "generation.invalid_data", field: "Items" }, 422)).toEqual([
      { title: "Document generation failed", detail: "Data is incompatible.", code: "generation.invalid_data", field: "Items" }
    ]);
    const errors = normalizeProblem({ errors: [
      { code: "one", message: "First", dataPath: "Items[0].Amount", location: "/word/document.xml, paragraph 1" },
      { code: "two", detail: "Second" }
    ] }, 422);
    expect(errors).toHaveLength(2);
    expect(errors[0]).toMatchObject({ dataPath: "Items[0].Amount", location: "/word/document.xml, paragraph 1" });
    expect(normalizeProblem({ traceId: "trace-1", limitReached: true, errors: [{ code: "one", message: "First" }] })[0])
      .toMatchObject({ traceId: "trace-1", limitReached: true });
  });

  it("uses safe fallbacks for malformed problems and unsafe filenames", async () => {
    expect(normalizeProblem("<html>failure</html>", 500)).toEqual([{ title: "Request failed", detail: "The server returned HTTP 500." }]);
    const responseBlob = new Blob([new Uint8Array([1])], { type: xlsxMime });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(success(responseBlob, xlsxMime)));
    const result = await generateDocument(new File([new Uint8Array([1])], "template.docx"), "{}");
    expect(result.filename).toBe("document.xlsx");
    expect(result.format).toBe("xlsx");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(success(responseBlob, xlsxMime)));
    expect((await generateDocument(new File([], "template.xlsx"), "{}")).filename).toBe("document.xlsx");
  });

  it("preserves ordered warning metadata and exact binary bytes from multipart", async () => {
    const bytes = new Uint8Array([0, 80, 75, 13, 10, 255, 128, 0, 13, 10]);
    const warnings = [
      { message: "Removed 2 prohibited XML 1.0 characters.", field: "{Name}", fieldCode: "{Name}", location: "A1", part: "/xl/worksheets/sheet1.xml", dataPath: "$.Name", code: "generation.xml_characters_removed", removedCharacterCount: 2 },
      { message: "Second warning.", field: "{Other}", location: "B2", dataPath: "$.Other" }
    ];
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(success(new Blob([bytes.buffer]), xlsxMime, "document.xlsx", warnings)));

    const result = await generateDocument(new File([], "template.docx"), "{}");

    expect(new Uint8Array(await result.blob.arrayBuffer())).toEqual(bytes);
    expect(result.warnings).toEqual([
      { title: "generation.xml_characters_removed", detail: warnings[0].message, code: "generation.xml_characters_removed", field: "{Name}", fieldCode: "{Name}", location: "A1", part: "/xl/worksheets/sheet1.xml", dataPath: "$.Name", removedCharacterCount: 2 },
      { title: "Generation warning", detail: warnings[1].message, field: "{Other}", location: "B2", dataPath: "$.Other" }
    ]);
  });

  it("rejects truncated and malformed multipart responses", async () => {
    const template = new File([], "template.docx");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(new Response("--broken", { status: 200, headers: { "Content-Type": "multipart/mixed; boundary=broken" } })));
    await expect(generateDocument(template, "{}")).rejects.toMatchObject<Partial<GenerationFailure>>({ diagnostics: [{ title: "Unexpected response", detail: expect.any(String) }] });
  });

  it("maps unavailable transport and invalid success responses to useful diagnostics", async () => {
    vi.stubGlobal("fetch", vi.fn().mockRejectedValue(new TypeError("network secret")));
    await expect(generateDocument(new File([new Uint8Array([1])], "template.docx"), "{}")).rejects.toMatchObject<Partial<GenerationFailure>>({
      diagnostics: [{ title: "Server unavailable", detail: "The generation service could not be reached." }]
    });
    const template = new File([new Uint8Array([1])], "template.docx");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(success(new Blob(["html"], { type: "text/html" }), "text/html")));
    await expect(generateDocument(template, "{}")).rejects.toBeInstanceOf(GenerationFailure);
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(success(new Blob([]), mime)));
    await expect(generateDocument(template, "{}")).rejects.toBeInstanceOf(GenerationFailure);
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({ ok: true, status: 200, blob: async () => new Blob(["bytes"]), headers: new Headers() } as Response));
    await expect(generateDocument(template, "{}")).rejects.toBeInstanceOf(GenerationFailure);
  });
});
