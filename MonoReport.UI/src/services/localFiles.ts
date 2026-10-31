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

export interface FileAccess {
  isSecureContext?: boolean;
  showOpenFilePicker?: (options: PickerOptions) => Promise<LocalFileHandle[]>;
  showSaveFilePicker?: (options: PickerOptions) => Promise<LocalFileHandle>;
}

export class LocalFiles {
  constructor(private readonly access: FileAccess) {}

  get canOpen() { return this.access.isSecureContext === true && typeof this.access.showOpenFilePicker === "function"; }
  get canSave() { return this.access.isSecureContext === true && typeof this.access.showSaveFilePicker === "function"; }

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

  async hasChanged(handle: LocalFileHandle, lastModified: number | undefined) {
    if (lastModified === undefined) throw new Error("The file's previous modification time is unavailable. Reopen it or use Save As before saving.");
    return (await handle.getFile()).lastModified !== lastModified;
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
