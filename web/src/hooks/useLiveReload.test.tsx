import { act, renderHook } from "@testing-library/react";
import { useState } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { LocalFiles, type LocalFileHandle } from "../services/localFiles";
import type { Dataset, TemplateEntry } from "../types";
import { useLiveReload } from "./useLiveReload";

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: Error) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}
function source(name: string, initial: string) {
  let contents = initial;
  const handle: LocalFileHandle = {
    name, getFile: vi.fn(async () => new File([contents], name, { lastModified: 42 })),
    isSameEntry: vi.fn(), createWritable: vi.fn()
  };
  return { handle, change: (value: string) => { contents = value; } };
}
async function setup() {
  const files = new LocalFiles({});
  const templateSource = source("report.docx", "one");
  const jsonSource = source("data.json", '{"v":1}');
  const templateRecord = { current: { id: "template", name: "report.docx", file: await files.readTemplateSnapshot(templateSource.handle), handle: templateSource.handle } as TemplateEntry | undefined };
  const datasetRecord = { current: { id: "data", name: "data.json", kind: "file", text: '{"v":1}', baseline: '{"v":1}', handle: jsonSource.handle, lastModified: 42 } as Dataset };
  const contentRevision = { current: 0 };
  const savingRef = { current: false };
  const saveRevision = { current: 0 };
  const generate = vi.fn(async () => undefined);
  const cancelAutomatic = vi.fn();
  const invalidateGeneration = vi.fn();
  const retryAutomatic = vi.fn();
  let replaceTemplate!: (entry: TemplateEntry) => void;
  let edit!: (text: string) => void;
  const hook = renderHook(() => {
    const [template, setTemplate] = useState(templateRecord.current);
    const [dataset, setDataset] = useState(datasetRecord.current);
    replaceTemplate = entry => { templateRecord.current = entry; setTemplate(entry); };
    edit = text => { contentRevision.current++; datasetRecord.current = { ...datasetRecord.current, text }; setDataset(datasetRecord.current); };
    return useLiveReload({
      files, template, templateRecord, dataset, datasetRecord, contentRevision, saving: savingRef.current, savingRef, saveRevision,
      busy: false, generate, cancelAutomatic, invalidateGeneration, retryAutomatic,
      applyTemplate: file => { contentRevision.current++; templateRecord.current = { ...templateRecord.current!, file }; setTemplate(templateRecord.current); },
      applyDataset: value => { if (value.text !== datasetRecord.current.text) contentRevision.current++; datasetRecord.current = value; setDataset(value); }
    });
  });
  return { ...hook, files, templateSource, jsonSource, templateRecord, datasetRecord, contentRevision, savingRef, saveRevision, generate, cancelAutomatic, invalidateGeneration, retryAutomatic,
    replace: (entry: TemplateEntry) => replaceTemplate(entry), edit: (text: string) => edit(text) };
}
async function tick(ms = 1000) { await act(async () => { await vi.advanceTimersByTimeAsync(ms); }); }
async function enable(hook: Awaited<ReturnType<typeof setup>>) { await act(async () => hook.result.current.setLiveReloadEnabled(true)); }

describe("Live Reload retained source lifecycle", () => {
  beforeEach(() => { vi.useFakeTimers(); });
  afterEach(() => { vi.useRealTimers(); });

  it("compares full content with equal metadata, preserves no-op identity and refreshes both sources independently", async () => {
    const hook = await setup();
    const initial = hook.templateRecord.current!.file;
    await enable(hook);
    await tick();
    expect(hook.templateRecord.current!.file).toBe(initial);
    expect(hook.contentRevision.current).toBe(0);
    expect(hook.generate).toHaveBeenCalledOnce();
    hook.templateSource.change("two");
    hook.jsonSource.change('{"v":2}');
    await tick();
    expect(await hook.templateRecord.current!.file!.text()).toBe("two");
    expect(hook.datasetRecord.current.text).toBe('{"v":2}');
    expect(hook.contentRevision.current).toBe(2);
    await tick(500);
    expect(hook.generate).toHaveBeenCalledTimes(2);
    expect(hook.invalidateGeneration).not.toHaveBeenCalled();
    expect(hook.templateSource.handle.createWritable).not.toHaveBeenCalled();
  });

  it("preserves dirty editor authority, reconciles matching editor content and excludes invalid disk JSON", async () => {
    const hook = await setup();
    act(() => hook.edit('{"v":3}'));
    hook.jsonSource.change('{"v":2}');
    await enable(hook);
    expect(hook.datasetRecord.current.text).toBe('{"v":3}');
    expect(hook.result.current.datasetStatus?.diskText).toBe('{"v":2}');
    const revision = hook.contentRevision.current;
    hook.jsonSource.change('{"v":3}');
    await tick();
    expect(hook.datasetRecord.current.baseline).toBe('{"v":3}');
    expect(hook.contentRevision.current).toBe(revision);
    hook.jsonSource.change("{");
    await tick();
    expect(hook.result.current.datasetStatus?.error).toContain("invalid");
    expect(hook.datasetRecord.current.text).toBe('{"v":3}');
    hook.jsonSource.change('{"v":4}');
    await tick();
    expect(hook.datasetRecord.current.text).toBe('{"v":4}');
  });

  it("skips overlap, keeps JSON active during slow Template reads and ignores obsolete same-handle samples", async () => {
    const hook = await setup();
    const pending = deferred<File>();
    vi.mocked(hook.templateSource.handle.getFile).mockReturnValueOnce(pending.promise);
    await enable(hook);
    const calls = vi.mocked(hook.templateSource.handle.getFile).mock.calls.length;
    hook.jsonSource.change('{"v":2}');
    await tick(3000);
    expect(hook.templateSource.handle.getFile).toHaveBeenCalledTimes(calls);
    expect(hook.datasetRecord.current.text).toBe('{"v":2}');
    act(() => hook.replace({ ...hook.templateRecord.current!, id: "replacement" }));
    hook.templateSource.change("current");
    await act(async () => pending.resolve(new File(["obsolete"], "report.docx")));
    expect(await hook.templateRecord.current!.file!.text()).toBe("current");
    expect(hook.templateSource.handle.getFile).toHaveBeenCalledTimes(calls + 1);
  });

  it("starts a different selected handle without waiting for an obsolete hung read", async () => {
    const hook = await setup();
    const pending = deferred<File>();
    vi.mocked(hook.templateSource.handle.getFile).mockReturnValueOnce(pending.promise);
    await enable(hook);
    const replacement = source("replacement.docx", "replacement");
    await act(async () => hook.replace({ id: "replacement", name: replacement.handle.name, handle: replacement.handle, file: hook.templateRecord.current!.file }));
    expect(await hook.templateRecord.current!.file!.text()).toBe("replacement");
    expect(replacement.handle.getFile).toHaveBeenCalledOnce();
    await act(async () => pending.resolve(new File(["obsolete"], "report.docx")));
    expect(await hook.templateRecord.current!.file!.text()).toBe("replacement");
  });

  it("recovers on later ticks after failure without an immediate retry loop", async () => {
    const hook = await setup();
    vi.mocked(hook.templateSource.handle.getFile).mockRejectedValueOnce(new DOMException("Busy", "NotReadableError"));
    await enable(hook);
    expect(hook.result.current.templateStatus?.error).toContain("unavailable");
    expect(hook.invalidateGeneration).toHaveBeenCalledOnce();
    expect(hook.templateSource.handle.getFile).toHaveBeenCalledTimes(2);
    await tick();
    expect(hook.result.current.templateStatus?.error).toBeUndefined();
    expect(hook.retryAutomatic).toHaveBeenCalledWith(true);
  });

  it("releases read-failed attempts after a same-content page-return reconciliation", async () => {
    const hook = await setup();
    vi.mocked(hook.templateSource.handle.getFile).mockRejectedValueOnce(new Error("Busy"));
    await enable(hook);
    hook.retryAutomatic.mockClear();
    await act(async () => window.dispatchEvent(new Event("pageshow")));
    expect(hook.retryAutomatic).toHaveBeenCalledWith(true);
    expect(hook.result.current.templateStatus?.error).toBeUndefined();
  });

  it("shares a pending sample with explicit preparation without a competing content commit", async () => {
    const hook = await setup();
    const pending = deferred<File>();
    vi.mocked(hook.templateSource.handle.getFile).mockReturnValueOnce(pending.promise);
    await enable(hook);
    let preparation!: Promise<void>;
    act(() => { preparation = hook.result.current.refreshTemplate(true); });
    await act(async () => { pending.resolve(new File(["prepared"], "report.docx")); await preparation; });
    expect(hook.contentRevision.current).toBe(1);
    expect(await hook.templateRecord.current!.file!.text()).toBe("prepared");
    expect(hook.templateSource.handle.getFile).toHaveBeenCalledTimes(2);
  });

  it("discards pre-return samples and coalesces page-return reconciliation without overlap", async () => {
    const hook = await setup();
    const pending = deferred<File>();
    vi.mocked(hook.templateSource.handle.getFile).mockReturnValueOnce(pending.promise);
    await enable(hook);
    hook.templateSource.change("returned");
    await act(async () => {
      Object.defineProperty(document, "visibilityState", { configurable: true, value: "visible" });
      document.dispatchEvent(new Event("visibilitychange"));
      window.dispatchEvent(new Event("pageshow"));
    });
    await act(async () => pending.resolve(new File(["old"], "report.docx")));
    expect(await hook.templateRecord.current!.file!.text()).toBe("returned");
    expect(hook.templateSource.handle.getFile).toHaveBeenCalledTimes(3);
  });

  it("ignores pre-Save JSON reads even when Save has already finished", async () => {
    const hook = await setup();
    const pending = deferred<File>();
    vi.mocked(hook.jsonSource.handle.getFile).mockReturnValueOnce(pending.promise);
    await enable(hook);
    hook.saveRevision.current += 2;
    await act(async () => pending.resolve(new File(['{"old":true}'], "data.json")));
    expect(hook.datasetRecord.current.text).toBe('{"v":1}');
  });

  it("waits for obsolete calls on rapid disable/re-enable and reconciles current bytes", async () => {
    const hook = await setup();
    const pending = deferred<File>();
    vi.mocked(hook.templateSource.handle.getFile).mockReturnValueOnce(pending.promise);
    await enable(hook);
    act(() => hook.result.current.setLiveReloadEnabled(false));
    await enable(hook);
    expect(hook.templateSource.handle.getFile).toHaveBeenCalledTimes(2);
    hook.templateSource.change("current");
    await act(async () => pending.resolve(new File(["obsolete"], "report.docx")));
    expect(await hook.templateRecord.current!.file!.text()).toBe("current");
    expect(hook.templateSource.handle.getFile).toHaveBeenCalledTimes(3);
  });

  it("lets explicit preparation finish when automatic monitoring is disabled", async () => {
    const hook = await setup();
    await enable(hook);
    const pending = deferred<File>();
    vi.mocked(hook.templateSource.handle.getFile).mockReturnValueOnce(pending.promise);
    let preparation!: Promise<void>;
    act(() => { preparation = hook.result.current.refreshTemplate(true); });
    act(() => hook.result.current.setLiveReloadEnabled(false));
    await act(async () => { pending.resolve(new File(["explicit"], "report.docx")); await preparation; });
    expect(await hook.templateRecord.current!.file!.text()).toBe("explicit");
  });

  it.each(["disable", "unmount"])("cleans timers/listeners and ignores pending samples on %s", async action => {
    const hook = await setup();
    const pending = deferred<File>();
    vi.mocked(hook.templateSource.handle.getFile).mockReturnValueOnce(pending.promise);
    await enable(hook);
    const accepted = hook.templateRecord.current!.file;
    if (action === "disable") act(() => hook.result.current.setLiveReloadEnabled(false));
    else hook.unmount();
    await act(async () => pending.resolve(new File(["late"], "report.docx")));
    const calls = vi.mocked(hook.templateSource.handle.getFile).mock.calls.length;
    await tick(5000);
    window.dispatchEvent(new Event("pageshow"));
    expect(hook.templateSource.handle.getFile).toHaveBeenCalledTimes(calls);
    expect(hook.templateRecord.current!.file).toBe(accepted);
  });
});
