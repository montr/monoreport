import { act, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { CellAddress, XlsxRangeSelectionContext, XlsxSelectionContext, XlsxSelectionContextCell, XlsxViewerOptions } from "@silurus/ooxml/xlsx";
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

function cell(row: number, col: number, displayText: string, formula?: string): XlsxSelectionContextCell {
  return { address: { row, col }, displayText, valueType: "text", value: "stored value", ...(formula === undefined ? {} : { formula }) };
}

function context(activeCell: CellAddress, cells: readonly XlsxSelectionContextCell[], extra: Partial<XlsxRangeSelectionContext> = {}): XlsxRangeSelectionContext {
  return {
    format: "xlsx", kind: "range", sheetIndex: 0, sheetName: "First",
    selection: {
      areas: [{ kind: "cells", top: 1, left: 1, bottom: activeCell.row, right: activeCell.col }],
      activeAreaIndex: 0, activeCell, extensionAnchor: activeCell
    },
    coordinateCountUpperBound: activeCell.row * activeCell.col,
    cells, truncated: false, truncationReasons: [], maxCells: 1000, textCharacters: 0, maxTextCharacters: 65_536,
    ...extra
  };
}

function select(selection: XlsxSelectionContext | null, session = 0) {
  const options = create.mock.calls[session][1] as XlsxViewerOptions;
  act(() => options.onSelectionContextChange!(selection));
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

  it.each(["SUM(A1:A2)", "=SUM(A1:A2)", "==SUM(A1:A2)"])("shows the active cell's formula with one equals sign: %s", async formula => {
    render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledOnce());
    select(context({ row: 27, col: 27 }, [
      cell(1, 1, "Other cell", "WRONG()"),
      { ...cell(27, 27, "", formula), valueType: "empty", value: null }
    ]));

    expect(screen.getByLabelText("Template active cell")).toHaveTextContent("AA27");
    const field = screen.getByRole("textbox", { name: "Template formula or value" });
    expect(field).toHaveValue("=SUM(A1:A2)");
    expect(field).toHaveAttribute("readonly");
    expect(screen.queryByText("Partial")).not.toBeInTheDocument();
  });

  it("shows formatted values verbatim, preserving zero, booleans and multiline text", async () => {
    render(<XlsxPreview source={source()} label="Result" />);
    await waitFor(() => expect(load).toHaveBeenCalledOnce());
    for (const displayText of ["0", "FALSE", "1,280.00", "<b>Workbook text</b>\nSecond line"]) {
      select(context({ row: 1, col: 1 }, [cell(1, 1, displayText)]));
      expect(screen.getByRole("textbox", { name: "Result formula or value" })).toHaveValue(displayText);
    }
    expect(screen.queryByText("Workbook text", { selector: "b" })).not.toBeInTheDocument();
  });

  it("clears content for empty or merge-covered coordinates and resets on sheet context changes", async () => {
    render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledOnce());
    const field = screen.getByRole("textbox", { name: "Template formula or value" });
    select(context({ row: 1, col: 1 }, [cell(1, 1, "First sheet")]));
    select(null);
    expect(field).toHaveValue("");
    expect(screen.getByLabelText("Template active cell")).toBeEmptyDOMElement();
    select(context({ row: 1, col: 1 }, [cell(1, 1, "Second sheet")], { sheetIndex: 1, sheetName: "Second" }));
    expect(field).toHaveValue("Second sheet");
    select(context({ row: 1, col: 2 }, []));
    expect(screen.getByLabelText("Template active cell")).toHaveTextContent("B1");
    expect(field).toHaveValue("");
    select({
      format: "xlsx", kind: "element", sheetIndex: 1, sheetName: "Second", elementType: "shape", elementIndex: 0,
      anchor: { from: { row: 1, col: 1, offsetX: 0, offsetY: 0 }, to: { row: 2, col: 2, offsetX: 0, offsetY: 0 } },
      truncated: false, truncationReasons: [], textCharacters: 0, maxTextCharacters: 65_536
    });
    expect(field).toHaveValue("");
    expect(screen.getByLabelText("Template active cell")).toBeEmptyDOMElement();
  });

  it("leaves unresolved content blank with its address and preserves available text in truncated contexts", async () => {
    render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledOnce());
    const field = screen.getByRole("textbox", { name: "Template formula or value" });
    select(context({ row: 1, col: 1 }, [cell(1, 1, "Previous value")]));
    select(context({ row: 1000, col: 16384 }, [cell(1, 1, "Wrong cell")], { truncated: true, truncationReasons: ["cells"] }));
    expect(screen.getByLabelText("Template active cell")).toHaveTextContent("XFD1000");
    expect(field).toHaveValue("");
    expect(screen.queryByText("Unavailable for this selection")).not.toBeInTheDocument();
    expect(screen.queryByText("Partial")).not.toBeInTheDocument();
    select(context({ row: 1, col: 1 }, [cell(1, 1, "", "SUM(A1:")], { truncated: true, truncationReasons: ["text"] }));
    expect(screen.getByLabelText("Template active cell")).toHaveTextContent("A1");
    expect(field).toHaveValue("=SUM(A1:");
    expect(screen.queryByText("Partial")).not.toBeInTheDocument();
    select(context({ row: 1, col: 1 }, [cell(1, 1, "Complete")]));
    expect(field).toHaveValue("Complete");
    expect(screen.queryByText("Partial")).not.toBeInTheDocument();
  });

  it("clears replacement and failed sessions and ignores callbacks after replacement, failure and disposal", async () => {
    const view = render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledOnce());
    const selected = context({ row: 1, col: 1 }, [cell(1, 1, "Old session")]);
    select(selected);
    const field = screen.getByRole("textbox", { name: "Template formula or value" });
    expect(field).toHaveValue("Old session");

    view.rerender(<XlsxPreview source={source()} label="Template" />);
    expect(field).toHaveValue("");
    await waitFor(() => expect(load).toHaveBeenCalledTimes(2));
    select(context({ row: 1, col: 1 }, [cell(1, 1, "Current session")]), 1);
    select(selected, 0);
    expect(field).toHaveValue("Current session");
    act(() => (create.mock.calls[1][1] as XlsxViewerOptions).onError!(new Error("failed")));
    expect(field).toHaveValue("");
    select(selected, 1);
    expect(field).toHaveValue("");

    // A dependency can emit synchronously from destroy(); the session must already be invalid.
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledTimes(3));
    select(selected, 2);
    destroy.mockImplementationOnce(() => select(context({ row: 2, col: 2 }, [cell(2, 2, "Disposed")]), 2));
    view.unmount();
    select(selected, 2);
    expect(screen.queryByRole("group", { name: "Template formula bar" })).not.toBeInTheDocument();
  });
});
