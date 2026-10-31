import { act, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { DocxPreview } from "./DocxPreview";

const renderAsync = vi.fn();
vi.mock("docx-preview", () => ({ renderAsync: (...args: unknown[]) => renderAsync(...args) }));

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

describe("DOCX preview boundary", () => {
  beforeEach(() => {
    renderAsync.mockReset();
  });

  it("reports a render failure without throwing away the supplied document", async () => {
    renderAsync.mockRejectedValueOnce(new Error("render detail"));
    const source = new Blob([new Uint8Array([1, 2, 3])]);
    render(<DocxPreview source={source} label="Result" />);
    expect(await screen.findByText("Result preview could not be rendered.")).toBeInTheDocument();
    expect(renderAsync).toHaveBeenCalledWith(source, expect.any(HTMLElement), undefined, expect.any(Object));
  });

  it("renders the supplied source directly", async () => {
    renderAsync.mockResolvedValue(undefined);
    const source = new Blob([new Uint8Array([1, 2, 3])]);
    render(<DocxPreview source={source} label="Template" />);

    await waitFor(() => expect(renderAsync).toHaveBeenCalledTimes(1));
    expect(renderAsync.mock.calls[0][0]).toBe(source);
  });

  it("ignores an obsolete render failure after source replacement", async () => {
    const firstRender = deferred<void>();
    const secondRender = deferred<void>();
    renderAsync.mockReturnValueOnce(firstRender.promise).mockReturnValueOnce(secondRender.promise);
    const first = new Blob([new Uint8Array([1])]);
    const second = new Blob([new Uint8Array([2])]);
    const view = render(<DocxPreview source={first} label="Template" />);
    view.rerender(<DocxPreview source={second} label="Template" />);
    await waitFor(() => expect(renderAsync).toHaveBeenCalledTimes(2));

    await act(async () => firstRender.reject(new Error("obsolete render failure")));
    expect(screen.queryByText("Template preview could not be rendered.")).not.toBeInTheDocument();
    expect(renderAsync.mock.calls[0][0]).toBe(first);
    expect(renderAsync.mock.calls[1][0]).toBe(second);
    await act(async () => secondRender.resolve());
  });

  it("isolates a replacement from an obsolete render host", async () => {
    renderAsync.mockReturnValue(new Promise(() => undefined));
    const first = new Blob([new Uint8Array([1])]);
    const second = new Blob([new Uint8Array([2])]);
    const view = render(<DocxPreview source={first} label="Template" />);
    await waitFor(() => expect(renderAsync).toHaveBeenCalledTimes(1));
    const firstHost = renderAsync.mock.calls.at(-1)?.[1] as HTMLElement;

    view.rerender(<DocxPreview source={second} label="Template" />);
    await waitFor(() => expect(renderAsync).toHaveBeenCalledTimes(2));

    const secondHost = renderAsync.mock.calls.at(-1)?.[1] as HTMLElement;
    expect(secondHost).not.toBe(firstHost);
    expect(firstHost.isConnected).toBe(false);
    expect(secondHost.isConnected).toBe(true);
  });
});
