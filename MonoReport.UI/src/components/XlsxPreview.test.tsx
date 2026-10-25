import { act, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { XlsxPreview } from "./XlsxPreview";

const { create, load, destroy } = vi.hoisted(() => ({ create: vi.fn(), load: vi.fn(), destroy: vi.fn() }));
vi.mock("@silurus/ooxml/xlsx", () => ({
  XlsxViewer: class {
    constructor(...args: unknown[]) { create(...args); }
    load = load;
    destroy = destroy;
  }
}));

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}
function source(bytes = new ArrayBuffer(3)) {
  const blob = new Blob([bytes]);
  Object.defineProperty(blob, "arrayBuffer", { value: vi.fn().mockResolvedValue(bytes) });
  return blob;
}

describe("XLSX preview boundary", () => {
  beforeEach(() => { create.mockReset(); load.mockReset().mockResolvedValue(undefined); destroy.mockReset(); });

  it("uses neutral styling while loading and error styling only after a preview failure", async () => {
    const pending = deferred<void>();
    load.mockReturnValueOnce(pending.promise);
    render(<XlsxPreview source={source()} label="Template" />);

    const loading = await screen.findByRole("status");
    expect(loading).toHaveTextContent("Loading template preview…");
    expect(loading).toHaveClass("preview-loading");
    expect(loading).not.toHaveClass("preview-warning");

    await act(async () => pending.reject(new Error("render failed")));
    const failure = await screen.findByRole("status");
    expect(failure).toHaveTextContent(/Template XLSX preview could not be rendered/);
    expect(failure).toHaveClass("preview-warning");
    expect(failure).not.toHaveClass("preview-loading");
  });

  it("loads private source bytes with read-only worker settings and destroys owned resources on unmount", async () => {
    const bytes = new ArrayBuffer(3);
    const blob = source(bytes);
    const view = render(<XlsxPreview source={blob} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledWith(bytes));
    expect(blob.arrayBuffer).toHaveBeenCalledOnce();
    expect(create).toHaveBeenCalledWith(expect.any(HTMLElement), expect.objectContaining({
      mode: "worker", hiddenSheetMode: "skip", useGoogleFonts: false, enableHyperlinks: false, resizable: false, workerTimeoutMs: 30_000
    }));
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
    const host = create.mock.calls[0][0] as HTMLElement;
    view.unmount();
    expect(destroy).toHaveBeenCalledOnce();
    expect(host.isConnected).toBe(false);
  });

  it("ignores an obsolete byte read and recovers on a new source", async () => {
    const read = deferred<ArrayBuffer>();
    const first = source();
    vi.mocked(first.arrayBuffer).mockReturnValue(read.promise);
    const view = render(<XlsxPreview source={first} label="Template" />);
    await waitFor(() => expect(first.arrayBuffer).toHaveBeenCalled());
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledOnce());
    await act(async () => read.reject(new Error("obsolete read")));
    expect(create).toHaveBeenCalledOnce();
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
  });

  it("isolates pending loads and late callbacks after replacement or format-switch unmount", async () => {
    const pending = deferred<void>();
    load.mockReturnValueOnce(pending.promise);
    const view = render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
    const firstHost = create.mock.calls[0][0] as HTMLElement;
    const onError = create.mock.calls[0][1].onError as () => void;
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledTimes(2));
    expect(firstHost.isConnected).toBe(false);
    expect(destroy).toHaveBeenCalledOnce();
    await act(async () => { onError(); pending.reject(new Error("obsolete load")); });
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
    const lateError = create.mock.calls[1][1].onError as () => void;
    view.unmount();
    act(() => lateError());
    expect(destroy).toHaveBeenCalledTimes(2);
  });

  it("ignores an obsolete successful load after a replacement source has rendered", async () => {
    const obsoleteLoad = deferred<void>();
    load.mockReturnValueOnce(obsoleteLoad.promise).mockResolvedValueOnce(undefined);
    const view = render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledOnce());
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledTimes(2));
    await waitFor(() => expect(screen.queryByRole("status")).not.toBeInTheDocument());

    await act(async () => obsoleteLoad.resolve());

    expect(screen.queryByRole("status")).not.toBeInTheDocument();
    expect(destroy).toHaveBeenCalledOnce();
    expect(create.mock.calls[1][0].isConnected).toBe(true);
  });

  it("disposes immediately on background/resource errors and recovers after a failed load", async () => {
    load.mockRejectedValueOnce(new Error("resource-limit: private document details"));
    const view = render(<XlsxPreview source={source()} label="Template" />);
    expect(await screen.findByText(/Template XLSX preview could not be rendered/)).toBeInTheDocument();
    expect(screen.queryByText(/private document details/)).not.toBeInTheDocument();
    expect(destroy).toHaveBeenCalledOnce();
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledTimes(2));
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
    act(() => create.mock.calls[1][1].onError(new Error("background render")));
    expect(screen.getByText(/Template XLSX preview could not be rendered/)).toBeInTheDocument();
    expect(destroy).toHaveBeenCalledTimes(2);
    view.unmount();
    expect(destroy).toHaveBeenCalledTimes(2);
  });

  it("reports byte-read and runtime-construction failures as preview errors", async () => {
    const blob = source();
    vi.mocked(blob.arrayBuffer).mockRejectedValueOnce(new Error("unreadable source"));
    const view = render(<XlsxPreview source={blob} label="Template" />);
    expect(await screen.findByText(/Template XLSX preview could not be rendered/)).toBeInTheDocument();
    expect(create).not.toHaveBeenCalled();
    create.mockImplementationOnce(() => { throw new Error("runtime unavailable"); });
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
    expect(screen.getByText(/Template XLSX preview could not be rendered/)).toBeInTheDocument();
    expect(load).not.toHaveBeenCalled();
  });

  it("does not start a viewer after unmount during lazy import", async () => {
    const blob = source();
    const view = render(<XlsxPreview source={blob} label="Template" />);
    view.unmount();
    await act(async () => {});
    expect(blob.arrayBuffer).not.toHaveBeenCalled();
    expect(create).not.toHaveBeenCalled();
  });
});
