import { describe, expect, it, vi } from "vitest";
import { LocalFiles, type LocalFileHandle } from "./localFiles";

function handle(name = "data.json", text = '{ "a": 1 }\r\n') {
  const writable = { write: vi.fn(async (_text: string) => undefined), close: vi.fn(async () => undefined), abort: vi.fn(async () => undefined) };
  const file = new File([text], name);
  const target: LocalFileHandle = { name, getFile: vi.fn(async () => file), isSameEntry: vi.fn(async () => false), createWritable: vi.fn(async () => writable) };
  return { target, writable, file, text };
}

describe("Local files", () => {
  it("detects secure picker capabilities independently", () => {
    expect(new LocalFiles({}).canOpen).toBe(false);
    expect(new LocalFiles({ showOpenFilePicker: vi.fn() }).canOpen).toBe(false);
    const open = new LocalFiles({ isSecureContext: true, showOpenFilePicker: vi.fn() });
    expect(open.canOpen).toBe(true);
    expect(open.canSave).toBe(false);
    expect(new LocalFiles({ isSecureContext: true, showSaveFilePicker: vi.fn() }).canSave).toBe(true);
  });

  it("fully reads fresh retained handles and copies template bytes into an independent File", async () => {
    const files = new LocalFiles({});
    const { target } = handle();
    const disk = new File(["updated bytes"], "template.xlsx", { type: "application/octet-stream", lastModified: 42 });
    vi.mocked(target.getFile).mockResolvedValue(disk);
    const snapshot = await files.readTemplateSnapshot(target);
    expect(snapshot).not.toBe(disk);
    expect([snapshot.name, snapshot.type, snapshot.lastModified]).toEqual([disk.name, disk.type, 42]);
    expect(await snapshot.text()).toBe("updated bytes");
    vi.spyOn(disk, "arrayBuffer").mockRejectedValue(new DOMException("Replaced", "NotReadableError"));
    expect(await snapshot.text()).toBe("updated bytes");
    await expect(files.readTemplateSnapshot(target)).rejects.toThrow("Replaced");
    vi.mocked(target.getFile).mockResolvedValue(new File(['{"fresh":true}'], "data.json", { lastModified: 43 }));
    expect(await files.readDatasetSnapshot(target)).toEqual({ text: '{"fresh":true}', lastModified: 43, name: "data.json", size: 14 });
    vi.mocked(target.getFile).mockResolvedValue(new File(["{"], "data.json"));
    await expect(files.readDatasetSnapshot(target)).rejects.toBeInstanceOf(SyntaxError);
    expect(target.createWritable).not.toHaveBeenCalled();
  });

  it("reuses stable accepted bytes on unchanged full reads and detects equal-metadata changes", async () => {
    const files = new LocalFiles({});
    const { target } = handle("report.docx", "one");
    const accepted = await files.readTemplateSnapshot(target);
    expect(await files.stableTemplate(accepted)).toBe(accepted);
    expect(await files.readTemplateSnapshot(target, accepted)).toBe(accepted);
    vi.mocked(target.getFile).mockResolvedValue(new File(["two"], accepted.name, { lastModified: accepted.lastModified }));
    const changed = await files.readTemplateSnapshot(target, accepted);
    expect(changed).not.toBe(accepted);
    expect(await changed.text()).toBe("two");
    expect(await accepted.text()).toBe("one");
  });

  it.each(["json", "template"] as const)("opens %s with original file identity and a reusable handle", async kind => {
    const { target, file } = handle();
    const picker = vi.fn(async () => [target]);
    const files = new LocalFiles({ isSecureContext: true, showOpenFilePicker: picker });
    expect(await files.open(kind)).toEqual({ file, handle: target });
    expect(picker).toHaveBeenCalledWith(expect.objectContaining({ multiple: false }));
    expect(JSON.stringify(picker.mock.calls)).toContain(kind === "json" ? ".json" : ".xlsx");
    if (kind === "json") expect(picker).toHaveBeenCalledWith(expect.objectContaining({ types: [expect.objectContaining({ description: "JSON file" })] }));
  });

  it("preserves exact JSON text including whitespace and rejects malformed or unreadable data", async () => {
    const files = new LocalFiles({});
    const { file, text } = handle();
    expect(await files.readJson(file)).toBe(text);
    await expect(files.readJson(new File(["{"], "bad.json"))).rejects.toBeInstanceOf(SyntaxError);
    const unreadable = new File([], "missing.json");
    vi.spyOn(unreadable, "text").mockRejectedValue(new DOMException("Unavailable", "NotFoundError"));
    await expect(files.readJson(unreadable)).rejects.toThrow("Unavailable");
  });

  it("selects Save As destination and writes exact text before committing", async () => {
    const { target, writable, text } = handle("copy.json");
    const picker = vi.fn(async () => target);
    const files = new LocalFiles({ isSecureContext: true, showSaveFilePicker: picker });
    expect(await files.destination("source.json")).toBe(target);
    expect(picker).toHaveBeenCalledWith(expect.objectContaining({ suggestedName: "source.json", types: [expect.objectContaining({ description: "JSON file" })] }));
    const saved = new File([text], "copy.json", { lastModified: 1234 });
    vi.mocked(target.getFile).mockResolvedValue(saved);
    expect(await files.write(target, text)).toBe(1234);
    expect(writable.write).toHaveBeenCalledWith(text);
    expect(writable.close).toHaveBeenCalledOnce();
    expect(writable.abort).not.toHaveBeenCalled();
    expect(writable.write.mock.invocationCallOrder[0]).toBeLessThan(writable.close.mock.invocationCallOrder[0]);
    expect(writable.close.mock.invocationCallOrder[0]).toBeLessThan(vi.mocked(target.getFile).mock.invocationCallOrder[0]);
  });

  it("compares disk timestamps and rejects unavailable checks without writing", async () => {
    const { target, file, writable } = handle();
    const files = new LocalFiles({});
    expect(await files.hasChanged(target, file.lastModified)).toBe(false);
    expect(await files.hasChanged(target, file.lastModified - 1)).toBe(true);
    await expect(files.hasChanged(target, undefined)).rejects.toThrow("previous modification time is unavailable");
    vi.mocked(target.getFile).mockRejectedValue(new DOMException("Access revoked", "NotAllowedError"));
    await expect(files.hasChanged(target, file.lastModified)).rejects.toThrow("Access revoked");
    expect(target.createWritable).not.toHaveBeenCalled();
    expect(writable.write).not.toHaveBeenCalled();
  });

  it("protects Save from different disk contents with identical timestamps", async () => {
    const { target, file } = handle();
    const files = new LocalFiles({});
    expect(await files.hasChanged(target, file.lastModified, await file.text())).toBe(false);
    expect(await files.hasChanged(target, file.lastModified, "other accepted content")).toBe(true);
    expect(target.createWritable).not.toHaveBeenCalled();
  });

  it("reports an unavailable timestamp after writing instead of claiming a reusable saved baseline", async () => {
    const { target, writable } = handle();
    vi.mocked(target.getFile).mockRejectedValue(new Error("File unavailable"));
    await expect(new LocalFiles({}).write(target, "{}")).rejects.toThrow("File unavailable");
    expect(writable.close).toHaveBeenCalledOnce();
    expect(writable.abort).not.toHaveBeenCalled();
  });

  it.each(["write", "close"] as const)("aborts an unsuccessful %s and preserves its error", async step => {
    const { target, writable } = handle();
    writable[step].mockRejectedValueOnce(new Error("Disk full"));
    await expect(new LocalFiles({}).write(target, "{}")).rejects.toThrow("Disk full");
    expect(writable.abort).toHaveBeenCalledOnce();
  });

  it("propagates cancelled pickers and denied or revoked write access without writing", async () => {
    const { target, writable } = handle();
    const cancelled = new DOMException("Cancelled", "AbortError");
    const files = new LocalFiles({ isSecureContext: true, showOpenFilePicker: vi.fn(async () => { throw cancelled; }), showSaveFilePicker: vi.fn(async () => { throw cancelled; }) });
    await expect(files.open("json")).rejects.toBe(cancelled);
    await expect(files.destination("x.json")).rejects.toBe(cancelled);
    vi.mocked(target.createWritable).mockRejectedValue(new DOMException("Denied", "NotAllowedError"));
    await expect(files.write(target, "{}")).rejects.toThrow("Denied");
    expect(writable.write).not.toHaveBeenCalled();
  });

  it("queries only on restoration and requests permission only for a user-initiated retry", async () => {
    const { target, file } = handle();
    const queryPermission = vi.fn(async (): Promise<PermissionState> => "prompt");
    const requestPermission = vi.fn(async (): Promise<PermissionState> => "granted");
    Object.assign(target, { queryPermission, requestPermission });
    const files = new LocalFiles({});
    await expect(files.reopen(target)).rejects.toThrow("requires your permission");
    expect(requestPermission).not.toHaveBeenCalled();
    expect(target.getFile).not.toHaveBeenCalled();
    expect(await files.reopen(target, true)).toBe(file);
    expect(requestPermission).toHaveBeenCalledWith({ mode: "read" });
    expect(queryPermission).toHaveBeenCalledOnce();
    requestPermission.mockResolvedValueOnce("denied");
    await expect(files.reopen(target, true)).rejects.toThrow("requires your permission");
    expect(target.getFile).toHaveBeenCalledOnce();
  });

  it("leaves missing permission APIs and missing disk files recoverable without requesting write access", async () => {
    const { target } = handle();
    const files = new LocalFiles({});
    await expect(files.reopen(target)).rejects.toThrow("permission checks are unavailable");
    Object.assign(target, { queryPermission: vi.fn(async () => "granted") });
    vi.mocked(target.getFile).mockRejectedValue(new DOMException("Deleted", "NotFoundError"));
    await expect(files.reopen(target)).rejects.toThrow("Deleted");
    expect(target.createWritable).not.toHaveBeenCalled();
  });

  it("deduplicates by entry identity rather than filename", async () => {
    const one = handle().target, two = handle().target;
    const files = new LocalFiles({});
    expect(await files.sameFile(one, undefined)).toBe(false);
    expect(await files.sameFile(one, one)).toBe(true);
    expect(await files.sameFile(one, two)).toBe(false);
    vi.mocked(one.isSameEntry).mockResolvedValueOnce(true);
    expect(await files.sameFile(one, two)).toBe(true);
  });
});
