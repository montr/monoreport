import { beforeEach, describe, expect, it, vi } from "vitest";
import { LocalFiles, type LocalFileHandle } from "./localFiles";
import { WorkspacePersistence, workspaceStorage, type WorkspaceContext } from "./workspacePersistence";
import { testIndexedDb } from "../test/indexedDb";

function handle(name: string) {
  return { name, getFile: vi.fn(), createWritable: vi.fn(), isSameEntry: vi.fn(async () => false) } satisfies LocalFileHandle;
}
function context(): WorkspaceContext { return { datasets: [], templates: [] }; }

describe("Workspace persistence", () => {
  beforeEach(() => localStorage.clear());

  it("distinguishes a first visit, an intentionally empty context, and unreadable metadata", async () => {
    const db = testIndexedDb();
    const service = new WorkspacePersistence(localStorage, db.factory, new LocalFiles({}));
    expect(await service.load()).toEqual({ warning: undefined });
    await service.save(context());
    expect((await service.load()).context).toEqual(context());
    localStorage.setItem(workspaceStorage.context, "broken");
    const failed = await service.load();
    expect(failed.context).toEqual(context());
    expect(failed.warning).toBeTruthy();
  });

  it("stores only bounded identity-deduplicated handles and explicit metadata, keeping the active older entries", async () => {
    const db = testIndexedDb();
    const service = new WorkspacePersistence(localStorage, db.factory, new LocalFiles({}));
    const datasets = Array.from({ length: 13 }, (_, index) => ({ id: `data-${index}`, name: `${index}.json`, kind: "file" as const, handle: handle(`${index}.json`), text: "PRIVATE JSON", baseline: "PRIVATE BASELINE" }));
    const templates = Array.from({ length: 12 }, (_, index) => ({ id: `template-${index}`, name: `${index}.docx`, handle: handle(`${index}.docx`), file: new File(["PRIVATE BYTES"], "template.docx") }));
    const alias = { ...datasets[12], id: "alias", handle: handle("alias.json") };
    datasets[12].handle.isSameEntry.mockImplementation(async (other?: LocalFileHandle) => other === alias.handle);
    alias.handle.isSameEntry.mockImplementation(async (other?: LocalFileHandle) => other === datasets[12].handle);
    await service.save({ datasets: [...datasets, alias], templates, datasetId: "data-0", templateId: "template-0", selectedExample: "custom" });
    const metadata = localStorage.getItem(workspaceStorage.context)!;
    expect(metadata).not.toMatch(/PRIVATE|baseline|text|file":/);
    const restored = (await service.load()).context!;
    expect(restored.datasets).toHaveLength(10);
    expect(restored.templates).toHaveLength(10);
    expect(restored.datasetId).toBe("data-0");
    expect(restored.datasets.some(item => item.id === "data-0")).toBe(true);
    expect(restored.templates.some(item => item.id === "template-0")).toBe(true);
    expect(restored.datasets.filter(item => item.id === "alias" || item.id === "data-12")).toHaveLength(1);
    const saved = db.databases.get(workspaceStorage.database)!.get(workspaceStorage.store)!.get(workspaceStorage.record) as { revision: string; handles: Array<{ key: string; handle: LocalFileHandle }> };
    expect(Object.keys(saved).sort()).toEqual(["handles", "metadata", "revision"]);
    expect(saved.handles).toHaveLength(20);
    expect(saved.handles.every(item => Object.keys(item).sort().join(",") === "handle,key")).toBe(true);
    expect(datasets).toHaveLength(13);
  });

  it("does not write again for content-only changes and does persist handle changes and closed entries", async () => {
    const db = testIndexedDb();
    const service = new WorkspacePersistence(localStorage, db.factory, new LocalFiles({}));
    const record = { id: "one", name: "one.json", kind: "file" as const, handle: handle("one.json"), text: "before" };
    await service.save({ datasets: [record], templates: [], datasetId: record.id });
    const writes = db.calls.filter(call => call.action === "put").length;
    const edited = { ...record, text: "after" };
    await service.save({ datasets: [edited], templates: [], datasetId: record.id });
    expect(db.calls.filter(call => call.action === "put")).toHaveLength(writes);
    const replacement = handle("one.json");
    await service.save({ datasets: [{ ...record, handle: replacement }], templates: [], datasetId: record.id });
    expect((await service.load()).context!.datasets[0].handle).toBe(replacement);
    await service.save(context());
    expect((await service.load()).context!.datasets).toEqual([]);
  });

  it("preserves source names without handles when IndexedDB is unsupported", async () => {
    const service = new WorkspacePersistence(localStorage, undefined, new LocalFiles({}));
    expect((await service.load()).warning).toMatch(/unavailable/);
    await service.save({ datasets: [{ id: "one", name: "one.json", kind: "file", handle: handle("one.json") }], templates: [] });
    const loaded = await service.load();
    expect(loaded.context!.datasets[0]).toMatchObject({ id: "one", name: "one.json", handle: undefined });
    expect(loaded.warning).toBeTruthy();
  });

  it("remembers fallback context without requiring the secure-context randomUUID API", async () => {
    const random = vi.spyOn(crypto, "randomUUID").mockImplementation(() => { throw new Error("Secure context required"); });
    const service = new WorkspacePersistence(localStorage, undefined, new LocalFiles({}));
    await service.save({ datasets: [], templates: [], selectedExample: "example" });
    expect((await service.load()).context!.selectedExample).toBe("example");
    expect(random).not.toHaveBeenCalled();
    random.mockRestore();
  });

  it("restores the IndexedDB context when local preference storage is unavailable", async () => {
    const db = testIndexedDb();
    const service = new WorkspacePersistence(localStorage, db.factory, new LocalFiles({}));
    const file = handle("one.json");
    await service.save({ datasets: [{ id: "one", name: file.name, kind: "file", handle: file }], templates: [], datasetId: "one" });
    const blocked = { getItem: () => { throw new Error("Blocked"); }, setItem: () => { throw new Error("Blocked"); }, removeItem: () => { throw new Error("Blocked"); } };
    const loaded = await new WorkspacePersistence(blocked, db.factory, new LocalFiles({})).load();
    expect(loaded.context!.datasetId).toBe("one");
    expect(loaded.context!.datasets[0].handle).toBe(file);
    expect(loaded.warning).toMatch(/preference storage is unavailable/);
  });

  it("never associates newer metadata with stale handles after cloning or transaction failure", async () => {
    const db = testIndexedDb();
    const service = new WorkspacePersistence(localStorage, db.factory, new LocalFiles({}));
    const original = handle("old.json");
    await service.save({ datasets: [{ id: "same-id", name: "old.json", kind: "file", handle: original }], templates: [] });
    db.failure.write = true;
    expect(await service.save({ datasets: [{ id: "same-id", name: "new.json", kind: "file", handle: handle("new.json") }], templates: [] })).toMatch(/could not be stored/);
    const loaded = await service.load();
    expect(loaded.context!.datasets[0]).toMatchObject({ name: "new.json", handle: undefined });
    expect(original.getFile).not.toHaveBeenCalled();
  });

  it("reports unavailable storage while leaving the session and unrelated storage usable", async () => {
    const db = testIndexedDb();
    const storage = { getItem: vi.fn(() => { throw new Error("Blocked"); }), setItem: vi.fn(() => { throw new Error("Quota"); }), removeItem: vi.fn(() => { throw new Error("Blocked"); }) };
    const service = new WorkspacePersistence(storage, db.factory, new LocalFiles({}));
    expect((await service.load()).warning).toBeTruthy();
    expect(await service.save(context())).toMatch(/could not be stored/);
    expect(await service.clear()).toMatch(/could not be removed/);
  });

  it("clears precisely owned keys and its handle store after pending writes, preserving other databases/stores and disk files", async () => {
    const db = testIndexedDb();
    const service = new WorkspacePersistence(localStorage, db.factory, new LocalFiles({}));
    const file = handle("one.json");
    localStorage.setItem("unrelated", "keep");
    for (const key of [workspaceStorage.desktop, workspaceStorage.mobile, workspaceStorage.desktopCollapsed, workspaceStorage.mobileCollapsed, workspaceStorage.diagnosticsSize, workspaceStorage.diagnosticsCollapsed, workspaceStorage.wordWrap]) localStorage.setItem(key, "value");
    db.databases.set("other-app", new Map([["data", new Map([["one", "keep"]])]]));
    db.databases.set(workspaceStorage.database, new Map([[workspaceStorage.store, new Map()], ["unrelated-store", new Map([["key", "keep"]])]]));
    const saving = service.save({ datasets: [{ id: "one", name: file.name, kind: "file", handle: file }], templates: [] });
    const clearing = service.clear();
    await saving;
    expect(await clearing).toBeUndefined();
    for (const key of [workspaceStorage.desktop, workspaceStorage.mobile, workspaceStorage.desktopCollapsed, workspaceStorage.mobileCollapsed, workspaceStorage.diagnosticsSize, workspaceStorage.diagnosticsCollapsed, workspaceStorage.wordWrap, workspaceStorage.context]) expect(localStorage.getItem(key)).toBeNull();
    expect(localStorage.getItem("unrelated")).toBe("keep");
    expect(db.databases.get("other-app")!.get("data")!.get("one")).toBe("keep");
    expect(db.databases.get(workspaceStorage.database)!.get(workspaceStorage.store)!.size).toBe(0);
    expect(db.databases.get(workspaceStorage.database)!.get("unrelated-store")!.get("key")).toBe("keep");
    expect(file.getFile).not.toHaveBeenCalled();
    expect(file.createWritable).not.toHaveBeenCalled();
  });

  it("reports partial clearing rather than claiming failed handle deletion succeeded", async () => {
    const db = testIndexedDb();
    db.failure.clear = true;
    const service = new WorkspacePersistence(localStorage, db.factory, new LocalFiles({}));
    localStorage.setItem(workspaceStorage.context, "old");
    expect(await service.clear()).toMatch(/could not be checked or cleared/);
    expect(localStorage.getItem(workspaceStorage.context)).toBeNull();
  });
});
