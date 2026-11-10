import { useEffect, useRef, useState, type ChangeEvent, type Dispatch, type DragEvent, type RefObject, type SetStateAction } from "react";
import { type LocalFiles, type LocalFileHandle, type NativeFileDropItem } from "../services/localFiles";
import { templateFormat } from "../services/documentFormat";
import type { DiagnosticsTab, Problem, TemplateEntry } from "../types";

interface Props {
  files: LocalFiles;
  resetRevision: RefObject<number>;
  contentRevision: RefObject<number>;
  templateRecord: RefObject<TemplateEntry | undefined>;
  nextId: RefObject<number>;
  selectTemplateEntry(entry: TemplateEntry | undefined): void;
  setErrors: Dispatch<SetStateAction<Problem[]>>;
  fileError(error: unknown, title: string): void;
  openDiagnostics(tab: DiagnosticsTab): void;
  retainTemplateHandle(entry: TemplateEntry, handle: LocalFileHandle): void;
}

export function useTemplateFileActions({ files, resetRevision, contentRevision, templateRecord, nextId, selectTemplateEntry, setErrors, fileError, openDiagnostics, retainTemplateHandle }: Props) {
  const reopenTemplateId = useRef<string | undefined>(undefined);
  const [dragging, setDragging] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);
  const dropBinding = useRef(0);
  useEffect(() => () => { dropBinding.current += 1; }, []);

  async function openTemplate(replaceId?: string) {
    const reset = resetRevision.current;
    const binding = ++dropBinding.current;
    const previousId = templateRecord.current?.id;
    if (!files.canOpen) { reopenTemplateId.current = replaceId; fileInput.current?.click(); return; }
    try {
      const opened = await files.open("template");
      if (!opened || reset !== resetRevision.current || (replaceId && templateRecord.current?.id !== replaceId)) return;
      if (binding !== dropBinding.current || templateRecord.current?.id !== previousId) return;
      const current = templateRecord.current;
      const same = current && await files.sameFile(current.handle, opened.handle);
      if (binding !== dropBinding.current || reset !== resetRevision.current || templateRecord.current?.id !== previousId) return;
      await acceptTemplate(opened.file, opened.handle, same ? current.id : replaceId);
    } catch (error) { if (reset === resetRevision.current) fileError(error, "Template could not be opened"); }
  }

  async function recoverTemplate(record: TemplateEntry) {
    if (!record.needsPermission || !record.handle?.requestPermission) { await openTemplate(record.id); return; }
    const revision = contentRevision.current;
    try {
      const file = await files.reopen(record.handle, true);
      if (revision === contentRevision.current && templateRecord.current === record) await acceptTemplate(file, record.handle, record.id);
    } catch (error) { if (revision === contentRevision.current) fileError(error, "Template could not be reopened"); }
  }

  async function acceptTemplate(file: File, handle?: LocalFileHandle, replaceId?: string) {
    if (!templateFormat(file.name)) {
      setErrors([{ title: "Unsupported template format", detail: "Select a .docx or .xlsx file." }]);
      openDiagnostics("problems");
      return;
    }
    if (replaceId && templateRecord.current?.id !== replaceId) return;
    const binding = ++dropBinding.current;
    const reset = resetRevision.current;
    const previous = templateRecord.current;
    try {
      const stable = await files.stableTemplate(file, replaceId === previous?.id ? previous?.file : undefined);
      if (binding !== dropBinding.current || reset !== resetRevision.current || templateRecord.current?.id !== previous?.id) return;
      selectTemplateEntry({ id: replaceId ?? `template-${++nextId.current}`, name: file.name, file: stable, handle });
      return templateRecord.current;
    } catch (error) {
      if (binding === dropBinding.current && reset === resetRevision.current) fileError(error, "Template could not be read");
    }
  }

  function selected(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    const replaceId = reopenTemplateId.current;
    reopenTemplateId.current = undefined;
    if (file) acceptTemplate(file, undefined, replaceId);
    event.target.value = "";
  }

  function drop(event: DragEvent) {
    event.preventDefault();
    setDragging(false);
    const file = event.dataTransfer.files?.[0];
    if (!file) return;
    let handleRequest: ReturnType<NonNullable<NativeFileDropItem["getAsFileSystemHandle"]>> | undefined;
    const reset = resetRevision.current;
    if (templateFormat(file.name)) {
      const item = Array.from(event.dataTransfer.items ?? []).find(item => item.kind === "file") as NativeFileDropItem | undefined;
      try {
        // Native handle acquisition must start in the drop event's synchronous task.
        if (item?.getAsFileSystemHandle) {
          handleRequest = item.getAsFileSystemHandle();
        }
      } catch (error) {
        // The ordinary File selection remains available if native access fails.
      }
    }
    // Attach a rejection handler immediately, before the asynchronous File read.
    const native = handleRequest?.catch(() => undefined);
    const accepted = acceptTemplate(file);
    const binding = dropBinding.current;
    void Promise.all([accepted, native]).then(([entry, handle]) => {
      if (binding !== dropBinding.current || reset !== resetRevision.current || !entry || templateRecord.current !== entry) return;
      if (handle?.kind === "file" && handle.name === file.name) retainTemplateHandle(entry, handle);
    });
  }

  function clearTemplate() {
    dropBinding.current += 1;
    selectTemplateEntry(undefined);
  }

  return { openTemplate, recoverTemplate, acceptTemplate, selected, drop, clearTemplate, dragging, setDragging, fileInput, reopenTemplateId };
}
