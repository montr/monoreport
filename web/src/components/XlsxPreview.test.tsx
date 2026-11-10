import { act, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellAddress, DefinedName, Worksheet, XlsxRangeSelectionContext, XlsxSelectionContext, XlsxSelectionContextCell, XlsxViewerOptions } from "@silurus/ooxml/xlsx";
import { XlsxPreview } from "./XlsxPreview";

const { create, load, destroy, workbookDestroy } = vi.hoisted(() => ({
  create: vi.fn(), load: vi.fn(), destroy: vi.fn(), workbookDestroy: vi.fn()
}));
vi.mock("@silurus/ooxml/xlsx", () => ({
  XlsxViewer: { fromWorkbook: (host: HTMLElement, workbook: unknown, options: XlsxViewerOptions) => create(host, options, workbook) },
  XlsxWorkbook: { load }
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
function workbook(definitions: DefinedName[][] = [[], []]) {
  return {
    sheetNames: ["First", "Second"],
    isHidden: vi.fn(() => false),
    getWorksheet: vi.fn(async (index: number) => ({ definedNames: definitions[index] } as Worksheet)),
    destroy: vi.fn(workbookDestroy)
  };
}
function makeViewer(_host: HTMLElement, options: XlsxViewerOptions, book: ReturnType<typeof workbook>) {
  const viewer = {
    sheetIndex: 0,
    getScale: vi.fn(() => 1),
    setScale: vi.fn(),
    getViewportOffset: vi.fn(() => ({ x: 0, y: 0 })),
    setViewportOffset: vi.fn(async (_offset: { x: number; y: number }) => undefined),
    destroy: vi.fn(destroy),
    goToSheet: vi.fn(async (index: number) => {
      viewer.sheetIndex = index;
      options.onSelectionStateChange?.(null);
      options.onSelectionContextChange?.(null);
      options.onSheetChange?.(index, book.sheetNames.length);
    }),
    scrollToCell: vi.fn(async (_cell: string, _options?: unknown) => {}),
    setSelection: vi.fn((range: string) => {
      const first = /^([A-Z]+)([0-9]+)/.exec(range)!;
      const col = [...first[1]].reduce((column, letter) => column * 26 + letter.charCodeAt(0) - 64, 0);
      const selected = context({ row: Number(first[2]), col }, []);
      options.onSelectionStateChange?.(selected.selection);
      options.onSelectionContextChange?.(selected);
    })
  };
  queueMicrotask(() => { options.onSheetChange?.(0, book.sheetNames.length); options.onReady?.(book.sheetNames); });
  return viewer;
}
function sessionViewer(session = 0) { return create.mock.results[session].value as ReturnType<typeof makeViewer>; }
function options(session = 0) { return create.mock.calls[session][1] as XlsxViewerOptions; }

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
  act(() => options(session).onSelectionContextChange!(selection));
}
async function choose(name: string, label = "Template") {
  const user = userEvent.setup();
  await user.click(screen.getByRole("combobox", { name: `${label} name box` }));
  await user.click(await screen.findByTitle(name));
}

beforeEach(() => {
  vi.stubGlobal("ResizeObserver", class {
    observe() {}
    unobserve() {}
    disconnect() {}
  });
  create.mockReset().mockImplementation(makeViewer);
  load.mockReset().mockImplementation(async () => workbook());
  destroy.mockReset();
  workbookDestroy.mockReset();
});
afterEach(() => vi.unstubAllGlobals());

describe("XLSX preview boundary", () => {
  it.each(["reordered", "hidden", "removed"])("restores same-source sheet by name, scale and offsets after %s sheets", async change => {
    const first = workbook();
    const second = workbook();
    second.sheetNames = change === "removed" ? ["First"] : ["Second", "First"];
    if (change === "hidden") second.isHidden.mockImplementation((...args: unknown[]) => args[0] === 0);
    load.mockResolvedValueOnce(first).mockResolvedValueOnce(second);
    const view = render(<XlsxPreview source={source()} label="Template" sourceId="same" />);
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
    const old = sessionViewer();
    old.sheetIndex = 1;
    old.getScale.mockReturnValue(1.5);
    old.getViewportOffset.mockReturnValue({ x: 90, y: 120 });
    view.rerender(<XlsxPreview source={source()} label="Template" sourceId="same" />);
    await waitFor(() => expect(create).toHaveBeenCalledTimes(2));
    const next = sessionViewer(1);
    await waitFor(() => expect(next.setViewportOffset).toHaveBeenCalledWith({ x: 90, y: 120 }));
    expect(next.sheetIndex).toBe(change === "hidden" ? 1 : 0);
    expect(next.setScale).toHaveBeenCalledWith(1.5);
    expect(old.destroy.mock.invocationCallOrder[0]).toBeLessThan(first.destroy.mock.invocationCallOrder[0]);
  });

  it("resets view for unrelated sources and ignores late restoration after replacement", async () => {
    const view = render(<XlsxPreview source={source()} label="Template" sourceId="one" />);
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
    sessionViewer().getScale.mockReturnValue(2);
    view.rerender(<XlsxPreview source={source()} label="Template" sourceId="two" />);
    await waitFor(() => expect(create).toHaveBeenCalledTimes(2));
    expect(sessionViewer(1).setScale).not.toHaveBeenCalled();
    const pending = deferred<void>();
    sessionViewer(1).sheetIndex = 1;
    create.mockImplementationOnce((host, opts, book) => {
      const viewer = makeViewer(host, opts, book);
      viewer.goToSheet.mockImplementationOnce(() => pending.promise);
      return viewer;
    });
    view.rerender(<XlsxPreview source={source()} label="Template" sourceId="two" />);
    await waitFor(() => expect(create).toHaveBeenCalledTimes(3));
    await waitFor(() => expect(sessionViewer(2).goToSheet).toHaveBeenCalled());
    view.unmount();
    await act(async () => pending.resolve());
    expect(sessionViewer(2).setScale).not.toHaveBeenCalled();
  });

  it("reports loading followed by a preview failure", async () => {
    const pending = deferred<ReturnType<typeof workbook>>();
    load.mockReturnValueOnce(pending.promise);
    render(<XlsxPreview source={source()} label="Template" />);
    expect(await screen.findByRole("status")).toHaveTextContent("Loading template preview…");
    await act(async () => pending.reject(new Error("render failed")));
    expect(await screen.findByRole("status")).toHaveTextContent(/Template XLSX preview could not be rendered/);
    expect(create).not.toHaveBeenCalled();
  });

  it("loads bytes once, shares the workbook and destroys the viewer before its workbook", async () => {
    const bytes = new ArrayBuffer(3);
    const blob = source(bytes);
    const book = workbook();
    load.mockResolvedValueOnce(book);
    const view = render(<XlsxPreview source={blob} label="Template" />);
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
    expect(load).toHaveBeenCalledExactlyOnceWith(bytes, {
      mode: "worker", useGoogleFonts: false, workerTimeoutMs: 30_000
    });
    expect(blob.arrayBuffer).toHaveBeenCalledOnce();
    expect(create).toHaveBeenCalledWith(expect.any(HTMLElement), expect.objectContaining({
      hiddenSheetMode: "skip", enableHyperlinks: false, resizable: false
    }), book);
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
    const host = create.mock.calls[0][0] as HTMLElement;
    view.unmount();
    expect(destroy).toHaveBeenCalledOnce();
    expect(book.destroy).toHaveBeenCalledOnce();
    expect(destroy.mock.invocationCallOrder[0]).toBeLessThan(workbookDestroy.mock.invocationCallOrder[0]);
    expect(host.isConnected).toBe(false);
  });

  it("ignores an obsolete byte read and recovers on a new source", async () => {
    const read = deferred<ArrayBuffer>();
    const first = source();
    vi.mocked(first.arrayBuffer).mockReturnValue(read.promise);
    const view = render(<XlsxPreview source={first} label="Template" />);
    await waitFor(() => expect(first.arrayBuffer).toHaveBeenCalled());
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
    await act(async () => read.reject(new Error("obsolete read")));
    expect(load).toHaveBeenCalledOnce();
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
  });

  it("disposes a late loaded workbook without constructing an obsolete viewer", async () => {
    const pending = deferred<ReturnType<typeof workbook>>();
    load.mockReturnValueOnce(pending.promise);
    const view = render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledOnce());
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
    const obsolete = workbook();
    await act(async () => pending.resolve(obsolete));
    expect(obsolete.destroy).toHaveBeenCalledOnce();
    expect(create).toHaveBeenCalledOnce();
    expect(destroy).not.toHaveBeenCalled();
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
    view.unmount();
    expect(destroy).toHaveBeenCalledOnce();
    expect(workbookDestroy).toHaveBeenCalledTimes(2);
  });

  it("ignores obsolete load rejection after unmount", async () => {
    const pending = deferred<ReturnType<typeof workbook>>();
    load.mockReturnValueOnce(pending.promise);
    const view = render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(load).toHaveBeenCalledOnce());
    view.unmount();
    await act(async () => pending.reject(new Error("obsolete load")));
    expect(create).not.toHaveBeenCalled();
  });

  it("cleans up background errors and recovers without exposing document details", async () => {
    load.mockRejectedValueOnce(new Error("resource-limit: private document details"));
    const view = render(<XlsxPreview source={source()} label="Template" />);
    expect(await screen.findByText(/Template XLSX preview could not be rendered/)).toBeInTheDocument();
    expect(screen.queryByText(/private document details/)).not.toBeInTheDocument();
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
    act(() => options().onError!(new Error("background render")));
    expect(screen.getByText(/Template XLSX preview could not be rendered/)).toBeInTheDocument();
    expect(destroy).toHaveBeenCalledOnce();
    expect(workbookDestroy).toHaveBeenCalledOnce();
    view.unmount();
    expect(destroy).toHaveBeenCalledOnce();
    expect(workbookDestroy).toHaveBeenCalledOnce();
  });

  it("cleans up workbook ownership for byte-read and construction failures", async () => {
    const blob = source();
    vi.mocked(blob.arrayBuffer).mockRejectedValueOnce(new Error("unreadable source"));
    const view = render(<XlsxPreview source={blob} label="Template" />);
    expect(await screen.findByText(/Template XLSX preview could not be rendered/)).toBeInTheDocument();
    expect(load).not.toHaveBeenCalled();
    create.mockImplementationOnce(() => { throw new Error("runtime unavailable"); });
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
    expect(screen.getByText(/Template XLSX preview could not be rendered/)).toBeInTheDocument();
    expect(workbookDestroy).toHaveBeenCalledOnce();
  });

  it("releases a viewer returned after a synchronous construction error callback", async () => {
    create.mockImplementationOnce((host, config, book) => {
      config.onError(new Error("construction error"));
      return makeViewer(host, config, book);
    });
    const view = render(<XlsxPreview source={source()} label="Template" />);
    expect(await screen.findByText(/Template XLSX preview could not be rendered/)).toBeInTheDocument();
    expect(destroy).toHaveBeenCalledOnce();
    expect(workbookDestroy).toHaveBeenCalledOnce();
    expect(destroy.mock.invocationCallOrder[0]).toBeLessThan(workbookDestroy.mock.invocationCallOrder[0]);
    view.unmount();
    expect(destroy).toHaveBeenCalledOnce();
    expect(workbookDestroy).toHaveBeenCalledOnce();
  });

  it("does not load bytes after unmount during lazy import", async () => {
    const blob = source();
    const view = render(<XlsxPreview source={blob} label="Template" />);
    view.unmount();
    await act(async () => {});
    expect(blob.arrayBuffer).not.toHaveBeenCalled();
    expect(create).not.toHaveBeenCalled();
  });

  it.each(["SUM(A1:A2)", "=SUM(A1:A2)", "==SUM(A1:A2)"])("shows the active cell's formula with one equals sign: %s", async formula => {
    render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
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
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
    for (const displayText of ["0", "FALSE", "1,280.00", "<b>Workbook text</b>\nSecond line", "  Leading\tand  trailing  \n\n  spaces  "]) {
      select(context({ row: 1, col: 1 }, [cell(1, 1, displayText)]));
      expect(screen.getByRole("textbox", { name: "Result formula or value" })).toHaveValue(displayText);
    }
    expect(screen.queryByText("Workbook text", { selector: "b" })).not.toBeInTheDocument();
  });

  it("clears content for empty or merge-covered coordinates and resets on sheet context changes", async () => {
    render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
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
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
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
    await waitFor(() => expect(create).toHaveBeenCalledOnce());
    const selected = context({ row: 1, col: 1 }, [cell(1, 1, "Old session")]);
    select(selected);
    const field = screen.getByRole("textbox", { name: "Template formula or value" });
    expect(field).toHaveValue("Old session");

    view.rerender(<XlsxPreview source={source()} label="Template" />);
    expect(field).toHaveValue("");
    await waitFor(() => expect(create).toHaveBeenCalledTimes(2));
    select(context({ row: 1, col: 1 }, [cell(1, 1, "Current session")]), 1);
    select(selected, 0);
    expect(field).toHaveValue("Current session");
    act(() => (create.mock.calls[1][1] as XlsxViewerOptions).onError!(new Error("failed")));
    expect(field).toHaveValue("");
    select(selected, 1);
    expect(field).toHaveValue("");

    // A dependency can emit synchronously from destroy(); the session must already be invalid.
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(create).toHaveBeenCalledTimes(3));
    select(selected, 2);
    destroy.mockImplementationOnce(() => select(context({ row: 2, col: 2 }, [cell(2, 2, "Disposed")]), 2));
    view.unmount();
    select(selected, 2);
    expect(screen.queryByRole("group", { name: "Template formula bar" })).not.toBeInTheDocument();
  });
});

describe("XLSX Name Box navigation", () => {
  it.each(["Template", "Result"])("shows the active address and navigates existing names read-only in %s", async label => {
    const book = workbook([[
      { name: "_xlnm.Print_Area", formula: "First!$B$2:$D$12" },
      { name: "Totals", formula: "First!$G$25" },
      { name: "Dynamic", formula: "OFFSET(First!$A$1,0,0,2,2)" }
    ], []]);
    load.mockResolvedValueOnce(book);
    render(<XlsxPreview source={source()} label={label} />);
    const box = await screen.findByRole("combobox", { name: `${label} name box` });
    await waitFor(() => expect(box).not.toBeDisabled());
    expect(box).toHaveAttribute("readonly");
    select(context({ row: 3, col: 5 }, [cell(3, 5, "Original", "SUM(A1:A2)")]));
    expect(screen.getByLabelText(`${label} active cell`)).toHaveTextContent("E3");

    const user = userEvent.setup();
    await user.click(box);
    expect(await screen.findByText("Print_Area")).toBeInTheDocument();
    expect(screen.getByText("Totals")).toBeInTheDocument();
    expect(screen.queryByText("Dynamic")).not.toBeInTheDocument();
    expect(screen.queryByRole("option", { name: "E3" })).not.toBeInTheDocument();
    expect(sessionViewer().setSelection).not.toHaveBeenCalled();
    expect(screen.getByRole("textbox", { name: `${label} formula or value` })).toHaveValue("=SUM(A1:A2)");
    await user.click(screen.getByTitle("Print_Area"));

    await waitFor(() => expect(sessionViewer().setSelection).toHaveBeenCalledExactlyOnceWith("B2:D12"));
    expect(sessionViewer().goToSheet).not.toHaveBeenCalled();
    expect(sessionViewer().scrollToCell).toHaveBeenCalledExactlyOnceWith("B2", { align: "nearest" });
    expect(screen.getByLabelText(`${label} active cell`)).toHaveTextContent("B2");
    select(context({ row: 18, col: 8 }, [cell(18, 8, "Next value")]));
    expect(screen.getByLabelText(`${label} active cell`)).toHaveTextContent("H18");
    expect(screen.getByRole("textbox", { name: `${label} formula or value` })).toHaveValue("Next value");
    await choose("Print_Area", label);
    await waitFor(() => expect(sessionViewer().setSelection).toHaveBeenCalledTimes(2));
    expect(screen.getByLabelText(`${label} active cell`)).toHaveTextContent("B2");
    expect(load).toHaveBeenCalledOnce();
  });

  it("reads Template and Result Print Areas independently", async () => {
    load.mockResolvedValueOnce(workbook([[{ name: "_xlnm.Print_Area", formula: "First!$A$1:$B$5" }], []]));
    load.mockResolvedValueOnce(workbook([[{ name: "_xlnm.Print_Area", formula: "First!$C$2:$E$20" }], []]));
    const template = render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(screen.getByRole("combobox", { name: "Template name box" })).not.toBeDisabled());
    await choose("Print_Area", "Template");
    template.unmount();
    render(<XlsxPreview source={source()} label="Result" />);
    await waitFor(() => expect(screen.getByRole("combobox", { name: "Result name box" })).not.toBeDisabled());
    await choose("Print_Area", "Result");
    expect(sessionViewer(0).setSelection).toHaveBeenCalledWith("A1:B5");
    expect(sessionViewer(1).setSelection).toHaveBeenCalledWith("C2:E20");
    expect(load).toHaveBeenCalledTimes(2);
  });

  it("activates a visible target worksheet before scrolling and selecting the whole range", async () => {
    const book = workbook([
      [{ name: "OtherReport", formula: "'Second'!$C$5:$F$30" }],
      [{ name: "_xlnm.Print_Area", formula: "Second!$C$5:$F$30" }]
    ]);
    load.mockResolvedValueOnce(book);
    render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(screen.getByRole("combobox", { name: "Template name box" })).not.toBeDisabled());
    const activated = deferred<void>();
    const scrolled = deferred<void>();
    const viewer = sessionViewer();
    viewer.goToSheet.mockImplementationOnce(async () => {
      await activated.promise;
      viewer.sheetIndex = 1;
      options().onSelectionStateChange?.(null);
      options().onSheetChange?.(1, 2);
    });
    viewer.scrollToCell.mockReturnValueOnce(scrolled.promise);
    await choose("OtherReport");
    expect(viewer.goToSheet).toHaveBeenCalledExactlyOnceWith(1);
    expect(viewer.scrollToCell).not.toHaveBeenCalled();
    await act(async () => activated.resolve());
    expect(viewer.scrollToCell).toHaveBeenCalledWith("C5", { align: "nearest" });
    expect(viewer.setSelection).not.toHaveBeenCalled();
    await act(async () => scrolled.resolve());
    expect(viewer.setSelection).toHaveBeenCalledExactlyOnceWith("C5:F30");
    await choose("Print_Area");
    expect(viewer.setSelection).toHaveBeenCalledTimes(2);
    expect(viewer.goToSheet).toHaveBeenCalledOnce();
    expect(book.getWorksheet).toHaveBeenCalledWith(1);
  });

  it("ignores obsolete worksheet metadata and disables navigation when no supported names exist", async () => {
    const oldMetadata = deferred<Worksheet>();
    const book = workbook();
    book.getWorksheet.mockReturnValueOnce(oldMetadata.promise);
    load.mockResolvedValueOnce(book);
    render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(book.getWorksheet).toHaveBeenCalledWith(0));
    await act(async () => sessionViewer().goToSheet(1));
    await act(async () => oldMetadata.resolve({ definedNames: [{ name: "Stale", formula: "First!$A$1" }] } as Worksheet));
    expect(screen.getByRole("combobox", { name: "Template name box" })).toBeDisabled();
    expect(sessionViewer().setSelection).not.toHaveBeenCalled();
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
  });

  it("keeps the most recent navigation when scrolling completes out of order", async () => {
    load.mockResolvedValueOnce(workbook([[
      { name: "FirstRange", formula: "First!$B$2:$C$3" },
      { name: "LastRange", formula: "First!$G$10:$H$20" }
    ], []]));
    render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(screen.getByRole("combobox", { name: "Template name box" })).not.toBeDisabled());
    const pending = deferred<void>();
    sessionViewer().scrollToCell.mockReturnValueOnce(pending.promise);
    await choose("FirstRange");
    await choose("LastRange");
    expect(sessionViewer().setSelection).toHaveBeenCalledExactlyOnceWith("G10:H20");
    await act(async () => pending.resolve());
    expect(sessionViewer().setSelection).toHaveBeenCalledOnce();
    expect(screen.getByLabelText("Template active cell")).toHaveTextContent("G10");
  });

  it.each(["selection", "worksheet", "replacement"])("does not apply pending navigation after a %s change", async change => {
    load.mockResolvedValueOnce(workbook([[
      { name: "Report", formula: "First!$B$2:$C$3" }
    ], []]));
    const view = render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(screen.getByRole("combobox", { name: "Template name box" })).not.toBeDisabled());
    const viewer = sessionViewer();
    const pending = deferred<void>();
    viewer.scrollToCell.mockReturnValueOnce(pending.promise);
    await choose("Report");
    if (change === "selection") {
      act(() => viewer.setSelection("H10"));
      viewer.setSelection.mockClear();
    } else if (change === "worksheet") {
      await act(async () => viewer.goToSheet(1));
    } else {
      view.rerender(<XlsxPreview source={source()} label="Template" />);
      await waitFor(() => expect(create).toHaveBeenCalledTimes(2));
    }
    await act(async () => pending.resolve());
    expect(viewer.setSelection).not.toHaveBeenCalled();
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
    if (change === "selection") expect(screen.getByLabelText("Template active cell")).toHaveTextContent("H10");
  });

  it("ignores a metadata failure after the viewer has activated another worksheet", async () => {
    const oldMetadata = deferred<Worksheet>();
    const book = workbook();
    book.getWorksheet.mockReturnValueOnce(oldMetadata.promise);
    load.mockResolvedValueOnce(book);
    render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(book.getWorksheet).toHaveBeenCalledWith(0));
    // Native sheetIndex changes before onSheetChange reports the committed frame.
    sessionViewer().sheetIndex = 1;
    await act(async () => oldMetadata.reject(new Error("obsolete metadata")));
    expect(destroy).not.toHaveBeenCalled();
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
    act(() => options().onSheetChange?.(1, 2));
    await waitFor(() => expect(book.getWorksheet).toHaveBeenCalledWith(1));
  });

  it("does not scroll or select after replacement during worksheet activation", async () => {
    load.mockResolvedValueOnce(workbook([[{ name: "Other", formula: "Second!$B$2:$C$3" }], []]));
    const view = render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(screen.getByRole("combobox", { name: "Template name box" })).not.toBeDisabled());
    const viewer = sessionViewer();
    const pending = deferred<void>();
    viewer.goToSheet.mockReturnValueOnce(pending.promise);
    await choose("Other");
    expect(viewer.goToSheet).toHaveBeenCalledWith(1);
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(create).toHaveBeenCalledTimes(2));
    await act(async () => pending.resolve());
    expect(viewer.scrollToCell).not.toHaveBeenCalled();
    expect(viewer.setSelection).not.toHaveBeenCalled();
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
  });

  it("ignores metadata from a replaced document and releases navigation failures", async () => {
    const oldMetadata = deferred<Worksheet>();
    const oldBook = workbook();
    oldBook.getWorksheet.mockReturnValueOnce(oldMetadata.promise);
    load.mockResolvedValueOnce(oldBook);
    load.mockResolvedValueOnce(workbook([[{ name: "Current", formula: "First!$A$10" }], []]));
    const view = render(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(oldBook.getWorksheet).toHaveBeenCalled());
    view.rerender(<XlsxPreview source={source()} label="Template" />);
    await waitFor(() => expect(screen.getByRole("combobox", { name: "Template name box" })).not.toBeDisabled());
    await act(async () => oldMetadata.reject(new Error("obsolete worksheet")));
    expect(screen.queryByRole("status")).not.toBeInTheDocument();
    sessionViewer(1).scrollToCell.mockRejectedValueOnce(new Error("render failed"));
    await choose("Current");
    expect(await screen.findByRole("status")).toHaveTextContent(/could not be rendered/);
    expect(destroy).toHaveBeenCalledTimes(2);
    expect(workbookDestroy).toHaveBeenCalledTimes(2);
    expect(sessionViewer(1).setSelection).not.toHaveBeenCalled();
  });
});
