export interface LocalFileHandle {
  name: string;
  getFile(): Promise<File>;
  isSameEntry(other: LocalFileHandle): Promise<boolean>;
  queryPermission?(options: { mode: "read" }): Promise<PermissionState>;
  requestPermission?(options: { mode: "read" }): Promise<PermissionState>;
  createWritable(): Promise<{
    write(data: string): Promise<void>;
    close(): Promise<void>;
    abort(): Promise<void>;
  }>;
}

interface PickerOptions {
  multiple?: boolean;
  suggestedName?: string;
  types: { description: string; accept: Record<string, string[]> }[];
}

export interface LocalDirectoryHandle { name: string; kind: "directory"; }
export interface NativeFileDropItem {
  kind: string;
  getAsFileSystemHandle?(): Promise<(LocalFileHandle & { kind: "file" }) | LocalDirectoryHandle | null>;
}
export interface FileAccess {
  isSecureContext?: boolean;
  showOpenFilePicker?: (options: PickerOptions) => Promise<LocalFileHandle[]>;
  showSaveFilePicker?: (options: PickerOptions) => Promise<LocalFileHandle>;
}

export interface FileRead { file: File; bytes: Uint8Array<ArrayBuffer>; }

export class LocalFiles {
  private readonly templateBytes = new WeakMap<File, Uint8Array<ArrayBuffer>>();
  constructor(private readonly access: FileAccess) {}

  get canOpen() { return this.access.isSecureContext === true && typeof this.access.showOpenFilePicker === "function"; }
  get canSave() { return this.access.isSecureContext === true && typeof this.access.showSaveFilePicker === "function"; }
  async readFile(handle: LocalFileHandle): Promise<FileRead> {
    const file = await handle.getFile();
    return { file, bytes: new Uint8Array(await file.arrayBuffer()) };
  }

  acceptTemplateRead({ file, bytes }: FileRead, previous?: File) {
    const accepted = previous && this.templateBytes.get(previous);
    if (accepted && this.equalBytes(accepted, bytes)) return previous!;
    const snapshot = new File([bytes], file.name, { type: file.type, lastModified: file.lastModified });
    this.templateBytes.set(snapshot, bytes);
    return snapshot;
  }

  equalBytes(left: Uint8Array, right: Uint8Array) {
    return left.length === right.length && left.every((value, index) => value === right[index]);
  }

  async stableTemplate(file: File, previous?: File) {
    if (this.templateBytes.has(file)) return file;
    return this.acceptTemplateRead({ file, bytes: new Uint8Array(await file.arrayBuffer()) }, previous);
  }

  async readTemplateSnapshot(handle: LocalFileHandle, previous?: File) {
    return this.acceptTemplateRead(await this.readFile(handle), previous);
  }

  inspectDatasetRead(read: FileRead) {
    const text = new TextDecoder().decode(read.bytes);
    try { JSON.parse(text); return { text, valid: true }; }
    catch { return { text, valid: false }; }
  }

  async readDatasetSnapshot(handle: LocalFileHandle) {
    const file = await handle.getFile();
    return { text: await this.readJson(file), lastModified: file.lastModified, name: file.name, size: file.size };
  }

  async open(kind: "json" | "template") {
    const handles = await this.access.showOpenFilePicker!({
      multiple: false,
      types: kind === "json"
        ? [{ description: "JSON file", accept: { "application/json": [".json"] } }]
        : [{ description: "Document template", accept: {
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"],
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"]
        } }]
    });
    const handle = handles[0];
    return handle ? { handle, file: await handle.getFile() } : undefined;
  }

  async reopen(handle: LocalFileHandle, grant = false) {
    const check = grant ? handle.requestPermission : handle.queryPermission;
    if (!check) throw new Error("File permission checks are unavailable. Reopen this file with the file picker.");
    if (await check.call(handle, { mode: "read" }) !== "granted") throw new DOMException("File access requires your permission. Use Grant access or Reopen to continue.", "NotAllowedError");
    return handle.getFile();
  }

  async destination(name: string) {
    return this.access.showSaveFilePicker!({ suggestedName: name, types: [{ description: "JSON file", accept: { "application/json": [".json"] } }] });
  }

  async hasChanged(handle: LocalFileHandle, lastModified: number | undefined, baseline?: string) {
    if (lastModified === undefined) throw new Error("The file's previous modification time is unavailable. Reopen it or use Save As before saving.");
    const file = await handle.getFile();
    return file.lastModified !== lastModified || (baseline !== undefined && await file.text() !== baseline);
  }

  async write(handle: LocalFileHandle, text: string) {
    const writable = await handle.createWritable();
    try {
      await writable.write(text);
      await writable.close();
    } catch (error) {
      try { await writable.abort(); } catch { /* Preserve the original write failure. */ }
      throw error;
    }
    return (await handle.getFile()).lastModified;
  }

  async sameFile(left: LocalFileHandle | undefined, right: LocalFileHandle | undefined) {
    return Boolean(left && right && (left === right || await left.isSameEntry(right)));
  }

  async readJson(file: File) {
    const text = await file.text();
    JSON.parse(text);
    return text;
  }
}
