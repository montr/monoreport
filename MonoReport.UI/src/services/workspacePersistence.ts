import type { LocalFileHandle, LocalFiles } from "./localFiles";

export const workspaceStorage = {
  desktop: "monoreport.playground.splitter.desktop",
  mobile: "monoreport.playground.splitter.mobile",
  desktopCollapsed: "monoreport.playground.splitter.desktop.collapsed",
  mobileCollapsed: "monoreport.playground.splitter.mobile.collapsed",
  diagnosticsSize: "monoreport.playground.splitter.diagnostics",
  diagnosticsCollapsed: "monoreport.playground.splitter.diagnostics.collapsed",
  wordWrap: "monoreport.playground.editor.wordWrap",
  context: "monoreport.playground.workspace",
  database: "monoreport.playground.workspace",
  store: "handles",
  record: "current"
} as const;

export interface DataReference {
  id: string;
  name: string;
  kind: "example" | "generated" | "upload" | "file";
  exampleId?: string;
  handle?: LocalFileHandle;
}
export interface TemplateReference { id: string; name: string; exampleId?: string; handle?: LocalFileHandle; }
export interface WorkspaceContext {
  datasets: DataReference[];
  templates: TemplateReference[];
  datasetId?: string;
  templateId?: string;
  selectedExample?: string;
}
interface MetadataReference extends Omit<DataReference, "handle"> { hasHandle: boolean; }
interface Metadata {
  version: 1;
  revision: string;
  datasets: MetadataReference[];
  templates: Array<Omit<TemplateReference, "handle"> & { hasHandle: boolean }>;
  datasetId?: string;
  templateId?: string;
  selectedExample?: string;
}
interface HandleSnapshot { revision: string; metadata: Metadata; handles: Array<{ key: string; handle: LocalFileHandle }>; }
const unavailable = "File-handle restoration is unavailable in this browser. Examples and file names can be remembered; local files require Reopen after a reload.";

export class WorkspacePersistence {
  private pending: Promise<string | undefined> = Promise.resolve(undefined);
  private previous?: WorkspaceContext;

  constructor(private readonly storage: Pick<Storage, "getItem" | "setItem" | "removeItem">, private readonly indexedDb: IDBFactory | undefined, private readonly files: LocalFiles) {}

  private async handles<T>(mode: IDBTransactionMode, operation: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
    if (!this.indexedDb) throw new Error(unavailable);
    const database = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = this.indexedDb!.open(workspaceStorage.database, 1);
      let blocked = false;
      request.onupgradeneeded = () => { request.result.createObjectStore(workspaceStorage.store); };
      request.onsuccess = () => { if (blocked) request.result.close(); else resolve(request.result); };
      request.onerror = () => reject(request.error ?? new Error("File-handle storage could not be opened."));
      request.onblocked = () => { blocked = true; reject(new Error("File-handle storage is blocked by another tab. Close it and retry.")); };
    });
    return new Promise<T>((resolve, reject) => {
      let result: T;
      try {
        const transaction = database.transaction(workspaceStorage.store, mode);
        transaction.oncomplete = () => { database.close(); resolve(result); };
        transaction.onabort = transaction.onerror = () => { database.close(); reject(transaction.error ?? new Error("File-handle storage failed.")); };
        const request = operation(transaction.objectStore(workspaceStorage.store));
        request.onsuccess = () => { result = request.result; };
      } catch (error) { database.close(); reject(error); }
    });
  }

  async load(): Promise<{ context?: WorkspaceContext; warning?: string }> {
    let raw: string | null = null;
    let warning = this.indexedDb ? undefined : unavailable;
    try { raw = this.storage.getItem(workspaceStorage.context); }
    catch { warning = "Workspace preference storage is unavailable; only accessible IndexedDB references can be restored."; }
    let snapshot: HandleSnapshot | undefined;
    if (this.indexedDb) {
      try { snapshot = await this.handles("readonly", store => store.get(workspaceStorage.record)); }
      catch { warning = "Stored workspace references could not be read. Reopen affected files; this session remains usable."; }
    }
    try {
      const metadata = (raw ? JSON.parse(raw) : snapshot?.metadata) as Metadata | undefined;
      if (!metadata) return { context: warning && this.indexedDb ? { datasets: [], templates: [] } : undefined, warning };
      const validReference = (item: TemplateReference) => item && typeof item.id === "string" && typeof item.name === "string" && (item.exampleId === undefined || typeof item.exampleId === "string");
      if (metadata.version !== 1 || typeof metadata.revision !== "string" || !Array.isArray(metadata.datasets) || !Array.isArray(metadata.templates) || !metadata.datasets.every(item => validReference(item) && ["example", "generated", "upload", "file"].includes(item.kind)) || !metadata.templates.every(validReference) || [metadata.datasetId, metadata.templateId, metadata.selectedExample].some(value => value !== undefined && typeof value !== "string")) throw new Error("The remembered workspace is invalid. Reopen your files to continue.");
      // Metadata must never resolve to handles left over from a failed or interrupted save.
      const handles = snapshot?.revision === metadata.revision && Array.isArray(snapshot.handles) ? snapshot.handles : [];
      return { warning, context: {
        datasets: metadata.datasets.slice(-10).map(item => ({ id: item.id, name: item.name, kind: item.kind, exampleId: item.exampleId, handle: handles.find(handle => handle.key === `data:${item.id}`)?.handle })),
        templates: metadata.templates.slice(-10).map(item => ({ id: item.id, name: item.name, exampleId: item.exampleId, handle: handles.find(handle => handle.key === `template:${item.id}`)?.handle })),
        datasetId: metadata.datasetId, templateId: metadata.templateId, selectedExample: metadata.selectedExample
      } };
    } catch (error) {
      return { context: { datasets: [], templates: [] }, warning: error instanceof Error ? error.message : "Workspace references could not be read. Reopen your files to continue." };
    }
  }

  private async bounded<T extends TemplateReference>(records: T[], activeId: string | undefined) {
    const active = records.find(item => item.id === activeId);
    const ordered = [...records.filter(item => item !== active), ...(active ? [active] : [])].reverse();
    const kept: T[] = [];
    for (const item of ordered) {
      let duplicate = false;
      for (const other of kept) {
        try { if (await this.files.sameFile(item.handle, other.handle)) { duplicate = true; break; } }
        catch { /* An inaccessible identity check does not close a live file. */ }
      }
      if (!duplicate && kept.length < 10) kept.push(item);
    }
    return kept.reverse();
  }

  save(context: WorkspaceContext): Promise<string | undefined> {
    // Explicit projection prevents structurally compatible live records from persisting content.
    const captured: WorkspaceContext = {
      datasets: context.datasets.map(item => ({ id: item.id, name: item.name, kind: item.kind, exampleId: item.exampleId, handle: item.handle })),
      templates: context.templates.map(item => ({ id: item.id, name: item.name, exampleId: item.exampleId, handle: item.handle })),
      datasetId: context.datasetId, templateId: context.templateId, selectedExample: context.selectedExample
    };
    const same = <T extends TemplateReference>(a: T[], b: T[]) => a.length === b.length && a.every((item, index) => Object.keys(item).every(key => item[key as keyof T] === b[index][key as keyof T]));
    if (this.previous && same(captured.datasets, this.previous.datasets) && same(captured.templates, this.previous.templates) && captured.datasetId === this.previous.datasetId && captured.templateId === this.previous.templateId && captured.selectedExample === this.previous.selectedExample) return this.pending;
    this.previous = captured;
    this.pending = this.pending.catch(() => undefined).then(async () => {
      const datasets = await this.bounded(captured.datasets, captured.datasetId);
      const templates = await this.bounded(captured.templates, captured.templateId);
      const revision = Array.from(crypto.getRandomValues(new Uint32Array(4))).join("-");
      const metadata: Metadata = { version: 1, revision, datasets: datasets.map(({ handle, ...item }) => ({ ...item, hasHandle: Boolean(handle) })), templates: templates.map(({ handle, ...item }) => ({ ...item, hasHandle: Boolean(handle) })), datasetId: captured.datasetId, templateId: captured.templateId, selectedExample: captured.selectedExample };
      let warning: string | undefined;
      try {
        if (this.indexedDb) await this.handles("readwrite", store => store.put({ revision, metadata, handles: [
          ...datasets.filter(item => item.handle).map(item => ({ key: `data:${item.id}`, handle: item.handle! })),
          ...templates.filter(item => item.handle).map(item => ({ key: `template:${item.id}`, handle: item.handle! }))
        ] } satisfies HandleSnapshot, workspaceStorage.record));
        else warning = unavailable;
      } catch { warning = "File handles could not be stored. File names are remembered; Reopen will be required after a reload."; }
      try { this.storage.setItem(workspaceStorage.context, JSON.stringify(metadata)); }
      catch { warning = "Workspace references could not be stored. This session remains usable, but reopening after a reload is not guaranteed."; }
      return warning;
    });
    return this.pending;
  }

  clear(): Promise<string | undefined> {
    this.previous = undefined;
    this.pending = this.pending.catch(() => undefined).then(async () => {
      const failures: string[] = [];
      for (const key of [workspaceStorage.desktop, workspaceStorage.mobile, workspaceStorage.desktopCollapsed, workspaceStorage.mobileCollapsed, workspaceStorage.diagnosticsSize, workspaceStorage.diagnosticsCollapsed, workspaceStorage.wordWrap, workspaceStorage.context]) {
        try { this.storage.removeItem(key); } catch { failures.push("Some MonoReport preferences or workspace references could not be removed."); }
      }
      try { await this.handles("readwrite", store => store.clear()); }
      catch { failures.push("Stored file handles could not be checked or cleared. Retry when browser storage is available."); }
      return failures.length ? [...new Set(failures)].join(" ") : undefined;
    });
    return this.pending;
  }
}
