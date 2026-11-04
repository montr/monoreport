import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { render, screen } from "@testing-library/react";
import { renderAsync } from "docx-preview";
import { describe, expect, it } from "vitest";
import { DocxPreview } from "./DocxPreview";

describe("DOCX preview with the real renderer", () => {
  it("skips embedded HTML while rendering the surrounding document text", async () => {
    const source = new File([new Uint8Array(readFileSync(resolve(import.meta.dirname, "../test/fixtures/embedded-html.docx")))], "embedded-html.docx");
    // Verify the fixture reaches the real renderer's HTML path without the preview's options.
    const control = document.createElement("div");
    await renderAsync(source, control);
    const frame = control.querySelector("iframe");
    expect(frame).not.toBeNull();
    expect(frame!.srcdoc).toContain("<script>");
    expect(frame!.srcdoc).toContain("Embedded HTML fixture");

    render(<DocxPreview source={source} label="Template" />);

    expect(await screen.findByText("Ordinary text after embedded HTML.")).toBeInTheDocument();
    const preview = screen.getByTestId("template-preview");
    expect(preview).toHaveTextContent("Ordinary text before embedded HTML.");
    expect(preview.querySelector("iframe")).toBeNull();
    expect(preview).not.toHaveTextContent("Embedded HTML fixture");
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
  });

  it("renders an ordinary DOCX sample", async () => {
    const source = new File([new Uint8Array(readFileSync(resolve(import.meta.dirname, "../../../Samples/sample.template.docx")))], "sample.template.docx");

    render(<DocxPreview source={source} label="Result" />);

    expect(await screen.findByText("Construction Project Cost Estimate")).toBeInTheDocument();
    expect(screen.getByTestId("result-preview")).toHaveTextContent("Materials estimate");
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
  });
});
