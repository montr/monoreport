import { act, renderHook } from "@testing-library/react";
import type { DragEvent } from "react";
import { describe, expect, it, vi } from "vitest";
import { LocalFiles, type LocalFileHandle, type NativeFileDropItem } from "../services/localFiles";
import type { TemplateEntry } from "../types";
import { useTemplateFileActions } from "./useTemplateFileActions";

function nativeHandle(name = "report.docx"): LocalFileHandle & { kind: "file" } {
  return { kind: "file", name, getFile: vi.fn(), createWritable: vi.fn(), isSameEntry: vi.fn() };
}
function setup() {
  const templateRecord = { current: undefined as TemplateEntry | undefined };
  const resetRevision = { current: 0 };
  const selectTemplateEntry = vi.fn((entry: TemplateEntry | undefined) => { templateRecord.current = entry; });
  const retainTemplateHandle = vi.fn((entry: TemplateEntry, handle: LocalFileHandle) => { templateRecord.current = { ...entry, handle }; });
  const setErrors = vi.fn();
  const fileError = vi.fn();
  const files = new LocalFiles({});
  const hook = renderHook(() => useTemplateFileActions({
    files, resetRevision, contentRevision: { current: 0 }, templateRecord, nextId: { current: 0 },
    selectTemplateEntry, retainTemplateHandle, setErrors, fileError, openDiagnostics: vi.fn()
  }));
  function drop(file: File, item: NativeFileDropItem) {
    hook.result.current.drop({ preventDefault: vi.fn(), dataTransfer: { files: [file], items: [item] } } as unknown as DragEvent);
  }
  return { ...hook, drop, templateRecord, resetRevision, selectTemplateEntry, retainTemplateHandle, setErrors, fileError };
}

describe("Template selection and native drop", () => {
  it("starts native acquisition synchronously and attaches access to stable content without selecting twice", async () => {
    const hook = setup();
    const file = new File(["selected bytes"], "report.docx");
    const handle = nativeHandle();
    let finish!: (value: typeof handle) => void;
    let insideDrop = true;
    const item = { kind: "file", getAsFileSystemHandle: vi.fn(() => {
      expect(insideDrop).toBe(true);
      return new Promise<typeof handle>(resolve => { finish = resolve; });
    }) };
    await act(async () => { hook.drop(file, item); insideDrop = false; });
    const accepted = hook.templateRecord.current!;
    expect(accepted.file).not.toBe(file);
    expect(await accepted.file!.text()).toBe("selected bytes");
    await act(async () => { finish(handle); });
    expect(hook.retainTemplateHandle).toHaveBeenCalledExactlyOnceWith(accepted, handle);
    expect(hook.selectTemplateEntry).toHaveBeenCalledOnce();
    expect(handle.getFile).not.toHaveBeenCalled();
  });

  it.each(["unsupported", "null", "directory", "mismatch", "throw", "reject"])("preserves File fallback for %s", async outcome => {
    const hook = setup();
    const file = new File(["private contents"], "report.docx");
    const item: NativeFileDropItem = { kind: "file" };
    if (outcome !== "unsupported") item.getAsFileSystemHandle = () => {
      if (outcome === "throw") throw new Error("Unavailable");
      if (outcome === "reject") return Promise.reject(new Error("Unavailable"));
      if (outcome === "directory") return Promise.resolve({ kind: "directory", name: file.name });
      if (outcome === "mismatch") return Promise.resolve(nativeHandle("other.docx"));
      return Promise.resolve(null);
    };
    await act(async () => { hook.drop(file, item); });
    expect(await hook.templateRecord.current?.file?.text()).toBe("private contents");
    expect(hook.templateRecord.current?.handle).toBeUndefined();
    expect(hook.retainTemplateHandle).not.toHaveBeenCalled();
  });

  it.each(["unmount", "reset", "close", "replace"])("ignores late native access after %s", async action => {
    const hook = setup();
    let finish!: (value: ReturnType<typeof nativeHandle>) => void;
    await act(async () => { hook.drop(new File([], "report.docx"), { kind: "file", getAsFileSystemHandle: () => new Promise(resolve => { finish = resolve; }) }); });
    if (action === "unmount") hook.unmount();
    else if (action === "reset") hook.resetRevision.current += 1;
    else if (action === "close") act(() => hook.result.current.clearTemplate());
    else await act(async () => { await hook.result.current.acceptTemplate(new File([], "new.xlsx")); });
    const current = hook.templateRecord.current;
    await act(async () => { finish(nativeHandle()); });
    expect(hook.retainTemplateHandle).not.toHaveBeenCalled();
    expect(hook.templateRecord.current).toBe(current);
  });

  it("rejects obsolete File reads and retains the previous selection on read failure", async () => {
    const hook = setup();
    const slow = new File([], "slow.docx");
    let finish!: (value: ArrayBuffer) => void;
    vi.spyOn(slow, "arrayBuffer").mockImplementation(() => new Promise(resolve => { finish = resolve; }));
    act(() => { void hook.result.current.acceptTemplate(slow); });
    await act(async () => { await hook.result.current.acceptTemplate(new File(["new"], "new.xlsx")); });
    const current = hook.templateRecord.current;
    await act(async () => { finish(new ArrayBuffer(0)); });
    expect(hook.templateRecord.current).toBe(current);
    const broken = new File([], "broken.docx");
    vi.spyOn(broken, "arrayBuffer").mockRejectedValue(new Error("Changed"));
    await act(async () => { await hook.result.current.acceptTemplate(broken); });
    expect(hook.templateRecord.current).toBe(current);
    expect(hook.fileError).toHaveBeenCalled();
  });
});
