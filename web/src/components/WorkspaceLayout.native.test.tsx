import { act, fireEvent, render, screen, within } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { useState } from "react";
import { ConfigProvider } from "antd";
import { WorkspaceLayout } from "./WorkspaceLayout";

// Supply geometry for JSDOM while exercising the installed Splitter's own controls and callbacks.
const observers = new Map<ResizeObserverCallback, Set<Element>>();
class TestResizeObserver {
  constructor(private callback: ResizeObserverCallback) { observers.set(callback, new Set()); }
  observe(target: Element) { observers.get(this.callback)!.add(target); }
  unobserve(target: Element) { observers.get(this.callback)!.delete(target); }
  disconnect() { observers.get(this.callback)!.clear(); }
}

const desktopKey = "monoreport.playground.splitter.desktop";
const mobileKey = "monoreport.playground.splitter.mobile";
const desktopCollapsedKey = "monoreport.playground.splitter.desktop.collapsed";
const mobileCollapsedKey = "monoreport.playground.splitter.mobile.collapsed";
const diagnosticsSizeKey = "monoreport.playground.splitter.diagnostics";
const diagnosticsCollapsedKey = "monoreport.playground.splitter.diagnostics.collapsed";
const diagnosticsExtent = 900;

function Workspace() {
  const [problemsOpen, setProblemsOpen] = useState(() => localStorage.getItem(diagnosticsCollapsedKey) === "false");
  return <ConfigProvider theme={{ token: { motion: false } }}><>
    <button onClick={() => { setProblemsOpen(true); localStorage.setItem(diagnosticsCollapsedKey, "false"); }}>Open diagnostics</button>
    <WorkspaceLayout data={<span>Editor</span>} document={<span>Viewer</span>} problems={<span>Problems</span>}
      problemsOpen={problemsOpen} onProblemsOpenChange={open => { setProblemsOpen(open); localStorage.setItem(diagnosticsCollapsedKey, String(!open)); }} />
  </>
  </ConfigProvider>;
}

describe.each([
  { orientation: "desktop", mobile: false, preference: 60, pixels: 864, extent: 1440 },
  { orientation: "stacked", mobile: true, preference: 35, pixels: 315, extent: 900 }
])("native workspace panel restoration in $orientation layout", ({ mobile, preference, pixels, extent }) => {
  beforeEach(() => {
    vi.stubGlobal("matchMedia", vi.fn((media: string) => ({
      matches: mobile, media, onchange: null,
      addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {}, dispatchEvent: () => false
    })));
    vi.stubGlobal("ResizeObserver", TestResizeObserver);
    vi.spyOn(HTMLElement.prototype, "offsetWidth", "get").mockReturnValue(1440);
    vi.spyOn(HTMLElement.prototype, "offsetHeight", "get").mockReturnValue(900);
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({ width: 1440, height: 900 } as DOMRect);
    localStorage.setItem(desktopKey, "60");
    localStorage.setItem(mobileKey, "35");
  });

  afterEach(() => { localStorage.clear(); vi.unstubAllGlobals(); });

  it.each(["Data", "Document"])("preserves the saved split through native %s collapse/restore and still saves subsequent dragging", async panel => {
    render(<Workspace />);
    await act(async () => {
      observers.forEach((targets, callback) => {
        if (targets.size) callback([...targets].map(target => ({ target }) as ResizeObserverEntry), {} as ResizeObserver);
      });
    });
    const dataPanel = screen.getByText("Editor").closest<HTMLElement>(".ant-splitter-panel")!;
    const viewer = screen.getByText("Viewer");
    const documentPanel = viewer.closest(".ant-splitter-panel")!;
    const workspace = dataPanel.parentElement!;
    const saved = vi.spyOn(Storage.prototype, "setItem");
    expect(workspace).toHaveClass(`ant-splitter-${mobile ? "vertical" : "horizontal"}`);
    expect(parseFloat(dataPanel.style.flexBasis)).toBeCloseTo(pixels);

    const collapsedPanel = panel === "Data" ? dataPanel : documentPanel;
    fireEvent.click(within(workspace).getByRole("button", { name: panel === "Data" ? "Toggle start panel" : "Toggle end panel" }));
    expect(collapsedPanel).toHaveClass("ant-splitter-panel-hidden");
    expect(localStorage.getItem(mobile ? mobileCollapsedKey : desktopCollapsedKey)).toBe(panel === "Data" ? "[true,false]" : "[false,true]");
    expect(localStorage.getItem(mobile ? mobileKey : desktopKey)).toBe(String(preference));

    fireEvent.click(within(workspace).getByRole("button", { name: panel === "Data" ? "Toggle end panel" : "Toggle start panel" }));
    expect(collapsedPanel).not.toHaveClass("ant-splitter-panel-hidden");
    expect(localStorage.getItem(mobile ? mobileCollapsedKey : desktopCollapsedKey)).toBe("[false,false]");
    expect(documentPanel).not.toHaveClass("ant-splitter-panel-hidden");
    expect(screen.getByText("Viewer")).toBe(viewer);
    expect(parseFloat(dataPanel.style.flexBasis)).toBeCloseTo(pixels);
    expect(localStorage.getItem(desktopKey)).toBe("60");
    expect(localStorage.getItem(mobileKey)).toBe("35");
    expect(saved).toHaveBeenCalled();

    fireEvent.mouseDown(within(workspace).getByRole("separator"), { clientX: 0, clientY: 0 });
    const offset = extent * (90 - preference) / 100;
    fireEvent.mouseMove(window, { clientX: mobile ? 0 : offset, clientY: mobile ? offset : 0 });
    fireEvent.mouseUp(window);
    expect(Number(localStorage.getItem(mobile ? mobileKey : desktopKey))).toBeCloseTo(90);
    expect(parseFloat(dataPanel.style.flexBasis)).toBeCloseTo(extent * 0.9);
  });

  it("restores diagnostics height after reload and preserves it through native collapse and expansion", async () => {
    localStorage.setItem(diagnosticsSizeKey, "30");
    localStorage.setItem(diagnosticsCollapsedKey, "true");
    const view = render(<Workspace />);
    await act(async () => {
      observers.forEach((targets, callback) => {
        if (targets.size) callback([...targets].map(target => ({ target }) as ResizeObserverEntry), {} as ResizeObserver);
      });
    });
    const diagnostics = screen.getByText("Problems").closest<HTMLElement>(".ant-splitter-panel")!;
    const outer = diagnostics.parentElement!;
    expect(diagnostics).toHaveClass("ant-splitter-panel-hidden");
    fireEvent.click(within(outer).getAllByRole("button", { name: "Toggle start panel" }).at(-1)!);
    expect(diagnostics).not.toHaveClass("ant-splitter-panel-hidden");
    expect(parseFloat(diagnostics.style.flexBasis)).toBeCloseTo(diagnosticsExtent * 0.3);

    fireEvent.mouseDown(within(outer).getAllByRole("separator").at(-1)!, { clientX: 0, clientY: 0 });
    fireEvent.mouseMove(window, { clientX: 0, clientY: -diagnosticsExtent * 0.1 });
    fireEvent.mouseUp(window);
    expect(Number(localStorage.getItem(diagnosticsSizeKey))).toBeCloseTo(40);

    fireEvent.click(within(outer).getAllByRole("button", { name: "Toggle end panel" }).at(-1)!);
    expect(diagnostics).toHaveClass("ant-splitter-panel-hidden");
    expect(localStorage.getItem(diagnosticsSizeKey)).toBe("40");
    expect(localStorage.getItem(diagnosticsCollapsedKey)).toBe("true");
    view.unmount();
    render(<Workspace />);
    const restored = screen.getByText("Problems").closest<HTMLElement>(".ant-splitter-panel")!;
    expect(restored).toHaveClass("ant-splitter-panel-hidden");
    fireEvent.click(screen.getByRole("button", { name: "Open diagnostics" }));
    expect(restored).not.toHaveClass("ant-splitter-panel-hidden");
    expect(restored.style.flexBasis).toBe("40%");
  });

  it("restores the last main-panel size when a collapsed panel is expanded after reload", async () => {
    localStorage.setItem(mobile ? mobileKey : desktopKey, String(preference));
    localStorage.setItem(mobile ? mobileCollapsedKey : desktopCollapsedKey, "[false,true]");
    render(<Workspace />);
    await act(async () => {
      observers.forEach((targets, callback) => {
        if (targets.size) callback([...targets].map(target => ({ target }) as ResizeObserverEntry), {} as ResizeObserver);
      });
    });
    const viewer = screen.getByText("Viewer");
    const documentPanel = viewer.closest<HTMLElement>(".ant-splitter-panel")!;
    expect(documentPanel).toHaveClass("ant-splitter-panel-hidden");
    const main = documentPanel.parentElement!;
    fireEvent.click(within(main).getByRole("button", { name: "Toggle start panel" }));
    expect(documentPanel).not.toHaveClass("ant-splitter-panel-hidden");
    const dataPanel = screen.getByText("Editor").closest<HTMLElement>(".ant-splitter-panel")!;
    expect(parseFloat(dataPanel.style.flexBasis)).toBeCloseTo(pixels);
    expect(localStorage.getItem(mobile ? mobileKey : desktopKey)).toBe(String(preference));
    expect(localStorage.getItem(mobile ? mobileCollapsedKey : desktopCollapsedKey)).toBe("[false,false]");
  });
});
