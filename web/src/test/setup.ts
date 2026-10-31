import "@testing-library/jest-dom/vitest";
import { vi } from "vitest";

Object.defineProperty(window, "matchMedia", {
  writable: true,
  value: (query: string) => ({
    matches: false,
    media: query,
    onchange: null,
    addListener: () => undefined,
    removeListener: () => undefined,
    addEventListener: () => undefined,
    removeEventListener: () => undefined,
    dispatchEvent: () => false
  })
});

Object.defineProperty(navigator, "clipboard", { value: { writeText: vi.fn() }, configurable: true });
Object.defineProperty(URL, "createObjectURL", { value: vi.fn(() => "blob:test"), configurable: true });
Object.defineProperty(URL, "revokeObjectURL", { value: vi.fn(), configurable: true });
vi.stubGlobal("__APP_VERSION__", "dev");
