import { describe, expect, it, vi } from "vitest";
import { generateDocument, GenerationFailure, normalizeProblem } from "./generation";

const mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const xlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function success(blob: Blob, contentType: string, disposition: string | null = "attachment; filename=document.docx") {
  return { ok: true, status: 200, blob: async () => blob, headers: new Headers({ "Content-Type": contentType, ...(disposition ? { "Content-Disposition": disposition } : {}) }) } as Response;
}

describe("generation client", () => {
  it("sends the exact multipart contract and returns actual DOCX bytes with a safe filename", async () => {
    const bytes = new Uint8Array([80, 75, 3, 4]);
    const responseBlob = new Blob([bytes], { type: mime });
    const fetchMock = vi.fn().mockResolvedValue(success(responseBlob, ` ${mime.toUpperCase()}; charset=binary`));
    vi.stubGlobal("fetch", fetchMock);
    const template = new File([bytes], "template.docx", { type: mime });

    const result = await generateDocument(template, "{\"value\":1}");

    expect(result.filename).toBe("document.docx");
    expect(result.format).toBe("docx");
    expect(result.blob).toBe(responseBlob);
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe("/api/documents/generate");
    expect(init.method).toBe("POST");
    expect(init.headers).toBeUndefined();
    const body = init.body as FormData;
    expect(body.get("template")).toBe(template);
    expect(body.get("data")).toBe("{\"value\":1}");
    expect(body.has("culture")).toBe(false);
  });

  it("selects XLSX from the response MIME despite a DOCX upload and preserves its exact bytes", async () => {
    const responseBlob = new Blob([new Uint8Array([80, 75, 7, 8])], { type: xlsxMime });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(success(responseBlob, `${xlsxMime}; charset=binary`, "attachment; filename*=UTF-8''r%C3%A9sultat.xlsx")));

    const result = await generateDocument(new File([new Uint8Array([1])], "template.docx", { type: mime }), "{}");

    expect(result).toEqual({ blob: responseBlob, filename: "résultat.xlsx", format: "xlsx" });
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
  });

  it("uses safe fallbacks for malformed problems and unsafe filenames", async () => {
    expect(normalizeProblem("<html>failure</html>", 500)).toEqual([{ title: "Request failed", detail: "The server returned HTTP 500." }]);
    const responseBlob = new Blob([new Uint8Array([1])], { type: xlsxMime });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(success(responseBlob, xlsxMime, "attachment; filename=../secret.xlsx")));
    const result = await generateDocument(new File([new Uint8Array([1])], "template.docx"), "{}");
    expect(result.filename).toBe("document.xlsx");
    expect(result.format).toBe("xlsx");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(success(responseBlob, xlsxMime, "attachment; filename=wrong.docx")));
    expect((await generateDocument(new File([], "template.xlsx"), "{}")).filename).toBe("document.xlsx");
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
