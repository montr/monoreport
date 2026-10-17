import { describe, expect, it, vi } from "vitest";
import { AnalysisFailure, analyzeTemplate } from "./analysis";

const valid = {
  valid: true,
  errors: [],
  warnings: [],
  requirements: [{ name: "Name", type: "value" }],
  occurrences: [{ kind: "field", expression: "{Name}", path: "$.Name", part: "/word/document.xml", location: "paragraph 1" }]
};

describe("analysis client", () => {
  it("posts the original file and accepts the dependency and occurrence contract", async () => {
    const fetchMock = vi.fn().mockResolvedValue(new Response(JSON.stringify(valid), { status: 200, headers: { "Content-Type": "application/json" } }));
    vi.stubGlobal("fetch", fetchMock);
    const template = new File([new Uint8Array([80, 75])], "template.docx");

    await expect(analyzeTemplate(template)).resolves.toEqual(valid);
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe("/api/documents/analyze");
    expect(init.method).toBe("POST");
    expect((init.body as FormData).get("template")).toBe(template);
  });

  it("preserves location-bearing diagnostics", async () => {
    const payload = { valid: false, errors: [{ code: "template.invalid_structure", message: "Invalid block", part: "/word/document.xml", location: "paragraph 2" }], warnings: [], requirements: [], occurrences: [] };
    vi.stubGlobal("fetch", vi.fn().mockResolvedValueOnce(new Response(JSON.stringify(payload), { status: 200 })));
    await expect(analyzeTemplate(new File([], "bad.docx"))).rejects.toMatchObject<Partial<AnalysisFailure>>({
      diagnostics: [{ title: "template.invalid_structure", detail: "Invalid block", part: "/word/document.xml", location: "paragraph 2" }]
    });
  });

  it("rejects HTTP, transport, and malformed successful responses", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValueOnce(new Response(JSON.stringify({ title: "Bad request", detail: "Missing template" }), { status: 400 })));
    await expect(analyzeTemplate(new File([], "bad.docx"))).rejects.toBeInstanceOf(AnalysisFailure);

    vi.stubGlobal("fetch", vi.fn().mockRejectedValueOnce(new TypeError("private network detail")));
    await expect(analyzeTemplate(new File([], "bad.docx"))).rejects.toMatchObject<Partial<AnalysisFailure>>({ diagnostics: [{ title: "Server unavailable", detail: "The analysis service could not be reached." }] });

    vi.stubGlobal("fetch", vi.fn().mockResolvedValueOnce(new Response(JSON.stringify({ ...valid, requirements: [{ name: "Name", type: "string" }] }), { status: 200 })));
    await expect(analyzeTemplate(new File([], "bad.docx"))).rejects.toMatchObject<Partial<AnalysisFailure>>({ diagnostics: [{ title: "Unexpected response", detail: "The server returned an invalid template analysis." }] });
  });
});
