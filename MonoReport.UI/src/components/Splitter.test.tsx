import { act, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Splitter } from "./Splitter";

describe("responsive splitter", () => {
  afterEach(() => {
    localStorage.clear();
    vi.unstubAllGlobals();
  });

  it("keeps CSS dimensions and accessible values synchronized across breakpoints", () => {
    let mobile = false;
    let listener: (() => void) | undefined;
    vi.stubGlobal("matchMedia", vi.fn(() => ({
      get matches() { return mobile; },
      media: "(max-width: 850px)",
      onchange: null,
      addListener: () => undefined,
      removeListener: () => undefined,
      addEventListener: (_event: string, callback: () => void) => { listener = callback; },
      removeEventListener: () => undefined,
      dispatchEvent: () => false
    })));

    render(<main className="layout"><Splitter /></main>);
    const splitter = screen.getByRole("separator");
    expect(splitter).toHaveAttribute("aria-orientation", "vertical");
    expect(splitter).toHaveAttribute("aria-valuenow", "45");
    fireEvent.keyDown(splitter, { key: "ArrowRight" });
    expect(splitter).toHaveAttribute("aria-valuenow", "48");
    expect(document.documentElement.style.getPropertyValue("--left-width")).toBe("48%");

    mobile = true;
    act(() => listener?.());
    expect(splitter).toHaveAttribute("aria-orientation", "horizontal");
    expect(splitter).toHaveAttribute("aria-valuenow", "50");
    expect(document.documentElement.style.getPropertyValue("--source-height")).toBe("50%");
    fireEvent.keyDown(splitter, { key: "ArrowDown" });
    expect(splitter).toHaveAttribute("aria-valuenow", "53");
    expect(document.documentElement.style.getPropertyValue("--source-height")).toBe("53%");

    mobile = false;
    act(() => listener?.());
    expect(splitter).toHaveAttribute("aria-orientation", "vertical");
    expect(splitter).toHaveAttribute("aria-valuenow", "48");
    expect(document.documentElement.style.getPropertyValue("--left-width")).toBe("48%");
  });

  it("restores independent desktop and mobile preferences after remount and breakpoint changes", () => {
    let mobile = false;
    let listener: (() => void) | undefined;
    vi.stubGlobal("matchMedia", vi.fn(() => ({
      get matches() { return mobile; },
      media: "(max-width: 850px)",
      onchange: null,
      addListener: () => undefined,
      removeListener: () => undefined,
      addEventListener: (_event: string, callback: () => void) => { listener = callback; },
      removeEventListener: () => undefined,
      dispatchEvent: () => false
    })));

    const view = render(<main className="layout"><Splitter /></main>);
    const splitter = screen.getByRole("separator");
    fireEvent.keyDown(splitter, { key: "ArrowRight" });
    expect(localStorage.getItem("monoreport.playground.splitter.desktop")).toBe("48");
    view.unmount();

    const reloaded = render(<main className="layout"><Splitter /></main>);
    expect(screen.getByRole("separator")).toHaveAttribute("aria-valuenow", "48");
    mobile = true;
    act(() => listener?.());
    expect(screen.getByRole("separator")).toHaveAttribute("aria-valuenow", "50");
    fireEvent.keyDown(screen.getByRole("separator"), { key: "ArrowDown" });
    expect(localStorage.getItem("monoreport.playground.splitter.mobile")).toBe("53");
    mobile = false;
    act(() => listener?.());
    expect(screen.getByRole("separator")).toHaveAttribute("aria-valuenow", "48");
    mobile = true;
    act(() => listener?.());
    expect(screen.getByRole("separator")).toHaveAttribute("aria-valuenow", "53");
    reloaded.unmount();

    render(<main className="layout"><Splitter /></main>);
    expect(screen.getByRole("separator")).toHaveAttribute("aria-valuenow", "53");
  });

  it("clamps a restored preference to the current panel minimums during viewport resizing", () => {
    let listener: (() => void) | undefined;
    vi.stubGlobal("matchMedia", vi.fn(() => ({
      matches: false,
      media: "(max-width: 850px)",
      onchange: null,
      addListener: () => undefined,
      removeListener: () => undefined,
      addEventListener: (_event: string, callback: () => void) => { listener = callback; },
      removeEventListener: () => undefined,
      dispatchEvent: () => false
    })));
    localStorage.setItem("monoreport.playground.splitter.desktop", "70");

    render(<main className="layout"><Splitter /></main>);
    const layout = document.querySelector(".layout")!;
    let width = 900;
    vi.spyOn(layout, "getBoundingClientRect").mockImplementation(() => ({
      width,
      height: 700,
      top: 0,
      right: width,
      bottom: 700,
      left: 0,
      x: 0,
      y: 0,
      toJSON: () => undefined
    }));

    fireEvent(window, new Event("resize"));
    const splitter = screen.getByRole("separator");
    expect(splitter).toHaveAttribute("aria-valuenow", "59");
    expect(localStorage.getItem("monoreport.playground.splitter.desktop")).toBe("70");
    width = 1400;
    fireEvent(window, new Event("resize"));
    expect(splitter).toHaveAttribute("aria-valuenow", "70");
    expect(document.documentElement.style.getPropertyValue("--left-width")).toBe("70%");
    act(() => listener?.());
  });
});
