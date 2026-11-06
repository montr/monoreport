import { useRef, useState, type ChangeEvent, type Dispatch, type DragEvent, type RefObject, type SetStateAction } from "react";
import type { LocalFiles, LocalFileHandle } from "../services/localFiles";
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
}

export function useTemplateFileActions({ files, resetRevision, contentRevision, templateRecord, nextId, selectTemplateEntry, setErrors, fileError, openDiagnostics }: Props) {
  const reopenTemplateId = useRef<string | undefined>(undefined);
  const [dragging, setDragging] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);

  async function openTemplate(replaceId?: string) {
    const reset = resetRevision.current;
    if (!files.canOpen) { reopenTemplateId.current = replaceId; fileInput.current?.click(); return; }
    try {
      const opened = await files.open("template");
      if (!opened || reset !== resetRevision.current || (replaceId && templateRecord.current?.id !== replaceId)) return;
      const current = templateRecord.current;
      if (current?.id !== replaceId && current && !current.restoreError && await files.sameFile(current.handle, opened.handle)) {
        if (reset !== resetRevision.current) return;
        selectTemplateEntry(current);
        return;
      }
      if (reset === resetRevision.current) acceptTemplate(opened.file, opened.handle, replaceId);
    } catch (error) { if (reset === resetRevision.current) fileError(error, "Template could not be opened"); }
  }

  async function recoverTemplate(record: TemplateEntry) {
    if (!record.needsPermission || !record.handle?.requestPermission) { await openTemplate(record.id); return; }
    const revision = contentRevision.current;
    try {
      const file = await files.reopen(record.handle, true);
      if (revision === contentRevision.current && templateRecord.current === record) acceptTemplate(file, record.handle, record.id);
    } catch (error) { if (revision === contentRevision.current) fileError(error, "Template could not be reopened"); }
  }

  function acceptTemplate(file: File, handle?: LocalFileHandle, replaceId?: string) {
    if (!templateFormat(file.name)) {
      setErrors([{ title: "Unsupported template format", detail: "Select a .docx or .xlsx file." }]);
      openDiagnostics("problems");
      return;
    }
    if (replaceId && templateRecord.current?.id !== replaceId) return;
    selectTemplateEntry({ id: `template-${++nextId.current}`, name: file.name, file, handle });
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
    if (file) acceptTemplate(file);
  }

  function clearTemplate() {
    selectTemplateEntry(undefined);
  }

  return { openTemplate, recoverTemplate, acceptTemplate, selected, drop, clearTemplate, dragging, setDragging, fileInput, reopenTemplateId };
}
