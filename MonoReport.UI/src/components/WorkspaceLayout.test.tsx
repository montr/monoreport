import { act, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { ReactElement, ReactNode } from "react";
import { WorkspaceLayout } from "./WorkspaceLayout";

const native = vi.hoisted(() => ({ splitters: [] as Array<{ orientation: string; onResize(sizes: number[]): void; onResizeStart?(): void; onResizeEnd?(sizes: number[]): void; onCollapse?(collapsed: boolean[], sizes: number[]): void }> }));
// Exercise application integration through the native component's public callbacks.
vi.mock("antd", () => {
  function Splitter({ children, ...props }: { children: ReactNode; orientation: string; onResize(sizes: number[]): void; onResizeStart?(): void; onResizeEnd?(sizes: number[]): void; onCollapse?(collapsed: boolean[], sizes: number[]): void }) {
    const index = (children as ReactElement<{ children: ReactElement }>[])[0].props.children.type === Splitter ? 0 : 1;
    native.splitters[index] = props;
    return <div data-testid={`splitter-${index}`} data-orientation={props.orientation}>{children}</div>;
  }
  Splitter.Panel = ({ children, size }: { children: ReactNode; size: string | number }) =>
    <div data-size={size}>{children}</div>;
  return { Splitter };
});

const desktopKey = "monoreport.playground.splitter.desktop";
const mobileKey = "monoreport.playground.splitter.mobile";
const desktopCollapsedKey = "monoreport.playground.splitter.desktop.collapsed";
const mobileCollapsedKey = "monoreport.playground.splitter.mobile.collapsed";
const props = () => ({
  data: <span>Editor</span>, document: <span>Viewer</span>, problems: <span>Problems</span>,
  problemsOpen: false, onProblemsOpenChange: vi.fn()
});

function viewport() {
  let mobile = false;
  let listener: (() => void) | undefined;
  vi.stubGlobal("matchMedia", vi.fn(() => ({
    get matches() { return mobile; },
    addEventListener: (_event: string, callback: () => void) => { listener = callback; },
    removeEventListener: () => undefined
  })));
  return (value: boolean) => { mobile = value; act(() => listener?.()); };
}

function panel(text: string) { return screen.getByText(text).parentElement!; }

describe("workspace layout", () => {
  afterEach(() => { localStorage.clear(); native.splitters.length = 0; vi.unstubAllGlobals(); });

  it("restores and saves independent native desktop and stacked panel sizes across remounts", () => {
    const setMobile = viewport();
    const view = render(<WorkspaceLayout {...props()} />);
    expect(native.splitters[1].orientation).toBe("horizontal");
    expect(panel("Editor")).toHaveAttribute("data-size", "45%");
    act(() => native.splitters[1].onResize([600, 400]));
    act(() => native.splitters[1].onResizeEnd?.([600, 400]));
    expect(localStorage.getItem(desktopKey)).toBe("60");
    setMobile(true);
    expect(native.splitters[1].orientation).toBe("vertical");
    expect(panel("Editor")).toHaveAttribute("data-size", "50%");
    act(() => native.splitters[1].onResize([350, 650]));
    act(() => native.splitters[1].onResizeEnd?.([350, 650]));
    expect(localStorage.getItem(mobileKey)).toBe("35");
    setMobile(false);
    expect(panel("Editor")).toHaveAttribute("data-size", "60%");
    view.unmount();
    native.splitters.length = 0;
    render(<WorkspaceLayout {...props()} />);
    expect(panel("Editor")).toHaveAttribute("data-size", "60%");
    setMobile(true);
    expect(panel("Editor")).toHaveAttribute("data-size", "35%");
  });

  it("restores valid preferences beyond the former custom panel limits", () => {
    const setMobile = viewport();
    localStorage.setItem(desktopKey, "99");
    localStorage.setItem(mobileKey, "1");
    render(<WorkspaceLayout {...props()} />);
    expect(panel("Editor")).toHaveAttribute("data-size", "99%");
    setMobile(true);
    expect(panel("Editor")).toHaveAttribute("data-size", "1%");
  });

  it("persists collapsed main panels separately from their last usable split size", () => {
    const setMobile = viewport();
    const view = render(<WorkspaceLayout {...props()} />);
    act(() => native.splitters[1].onResize([600, 400]));
    act(() => native.splitters[1].onResizeEnd?.([600, 400]));
    act(() => native.splitters[1].onCollapse?.([false, true], [1000, 0]));
    expect(localStorage.getItem(desktopCollapsedKey)).toBe("[false,true]");
    expect(localStorage.getItem(desktopKey)).toBe("60");
    expect(panel("Viewer")).toHaveAttribute("data-size", "0%");
    view.unmount();
    native.splitters.length = 0;
    render(<WorkspaceLayout {...props()} />);
    expect(panel("Viewer")).toHaveAttribute("data-size", "0%");
    act(() => native.splitters[1].onResize([500, 500]));
    act(() => native.splitters[1].onResizeEnd?.([500, 500]));
    expect(localStorage.getItem(desktopKey)).toBe("60");
    act(() => native.splitters[1].onCollapse?.([false, false], [550, 450]));
    expect(panel("Viewer")).toHaveAttribute("data-size", "40%");
    expect(localStorage.getItem(desktopKey)).toBe("60");
    setMobile(true);
    expect(localStorage.getItem(mobileCollapsedKey)).toBeNull();
  });

  it("persists and restores the lower diagnostics panel height while preserving it when collapsed", () => {
    const callbacks = { ...props(), problemsOpen: true };
    const view = render(<WorkspaceLayout {...callbacks} />);
    act(() => native.splitters[0].onResize([650, 350]));
    act(() => native.splitters[0].onResizeEnd?.([650, 350]));
    expect(localStorage.getItem("monoreport.playground.splitter.diagnostics")).toBe("35");
    act(() => native.splitters[0].onCollapse?.([false, true], [1000, 0]));
    expect(callbacks.onProblemsOpenChange).toHaveBeenCalledWith(false);
    expect(localStorage.getItem("monoreport.playground.splitter.diagnostics")).toBe("35");
    view.unmount();
    native.splitters.length = 0;
    render(<WorkspaceLayout {...callbacks} problemsOpen />);
    expect(panel("Problems")).toHaveAttribute("data-size", "35%");
  });
});
