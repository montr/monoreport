import { afterEach, describe, expect, it, vi } from "vitest";
import { requestSampleJson, SampleJsonFailure } from "./sampleJsonRequest";

afterEach(() => vi.unstubAllGlobals());

describe("sample JSON request", () => {
  it("posts settings and returns detached JSON data with ordered warnings", async () => {
    const file = new File(["template"], "sample.docx");
    const fetchMock = vi.fn().mockResolvedValue(new Response(JSON.stringify({ data: { Name: "Ada" }, warnings: [{ code: "sample.unknown_shape", message: "Unknown shape", location: "paragraph 2" }] }), { status: 200 }));
    vi.stubGlobal("fetch", fetchMock);
    const result = await requestSampleJson(file, { itemsPerCollection: 10, values: "filled" }, new AbortController().signal);
    expect(result).toEqual({ data: { Name: "Ada" }, warnings: [{ title: "sample.unknown_shape", detail: "Unknown shape", code: "sample.unknown_shape", location: "paragraph 2" }] });
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe("/api/documents/sample");
    expect((init.body as FormData).get("template")).toBe(file);
    expect((init.body as FormData).get("itemsPerCollection")).toBe("10");
    expect((init.body as FormData).get("values")).toBe("filled");
  });

  it("preserves actionable Problem Details and cancellation", async () => {
    const problem = {
      type: "urn:monoreport:error:sample.element_limit",
      title: "Sample generation failed",
      status: 422,
      detail: "Reduce itemsPerCollection.",
      code: "sample.element_limit",
      traceId: "trace-42",
      field: null,
      limitReached: false,
      errors: [{
        code: "sample.element_limit",
        message: "The 1,000,000 element limit was exceeded.",
        part: "/word/document.xml",
        location: "/word/document.xml, paragraph 3"
      }, {
        code: "sample.template_issue",
        message: "A separate template limitation.",
        part: "/word/header1.xml",
        location: "/word/header1.xml, paragraph 2"
      }]
    };
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(new Response(JSON.stringify(problem), { status: 422 })));
    await expect(requestSampleJson(new File([], "sample.docx"), { itemsPerCollection: 10, values: "empty" }, new AbortController().signal)).rejects.toMatchObject<Partial<SampleJsonFailure>>({
      diagnostics: [
        { title: "sample.element_limit", code: "sample.element_limit", detail: "The 1,000,000 element limit was exceeded.", location: "/word/document.xml, paragraph 3", traceId: "trace-42", limitReached: false },
        { title: "sample.template_issue", code: "sample.template_issue", detail: "A separate template limitation.", location: "/word/header1.xml, paragraph 2" }
      ],
      response: problem
    });
    const controller = new AbortController();
    controller.abort();
    vi.stubGlobal("fetch", vi.fn().mockRejectedValue(new DOMException("Aborted", "AbortError")));
    await expect(requestSampleJson(new File([], "sample.docx"), { itemsPerCollection: 10, values: "empty" }, controller.signal)).rejects.toMatchObject({ name: "AbortError" });
  });
});
