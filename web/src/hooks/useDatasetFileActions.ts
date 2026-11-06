import { useRef, useState, type Dispatch, type RefObject, type SetStateAction } from "react";
import { message, Modal } from "antd";
import type { LocalFiles, LocalFileHandle } from "../services/localFiles";
import type { Dataset, Problem } from "../types";

interface Props {
  files: LocalFiles;
  resetRevision: RefObject<number>;
  contentRevision: RefObject<number>;
  datasetRecord: RefObject<Dataset>;
  confirmDiscard(record: Dataset | undefined, title: string, content: string, okText: string): Promise<boolean>;
  addDataset(record: Omit<Dataset, "id">): void;
  setDatasetRecord(record: Dataset): void;
  setErrors: Dispatch<SetStateAction<Problem[]>>;
  fileError(error: unknown, title: string): void;
  downloadBlob(blob: Blob, filename: string): void;
  modal: ReturnType<typeof Modal.useModal>[0];
  feedback: ReturnType<typeof message.useMessage>[0];
}

export function useDatasetFileActions({ files, resetRevision, contentRevision, datasetRecord, confirmDiscard, addDataset, setDatasetRecord, setErrors, fileError, downloadBlob, modal, feedback }: Props) {
  const reopenDataId = useRef<string | undefined>(undefined);
  const [saving, setSaving] = useState(false);
  const savingRef = useRef(false);
  const dataInput = useRef<HTMLInputElement>(null);

  async function openData(file?: File, replaceId?: string, handle?: LocalFileHandle) {
    const reset = resetRevision.current;
    const revision = contentRevision.current;
    const record = datasetRecord.current;
    try {
      const opened = file ? { file, handle } : await files.open("json");
      if (!opened || reset !== resetRevision.current || revision !== contentRevision.current || (replaceId && record?.id !== replaceId)) return;
      if (record && !record.restoreError && await files.sameFile(record.handle, opened.handle)) return;
      const text = await files.readJson(opened.file);
      if (reset !== resetRevision.current || revision !== contentRevision.current) return;
      if (!await confirmDiscard(record, "Open dataset?", "Replace the current dataset, discarding unsaved changes?", "Replace")) return;
      if (revision !== contentRevision.current) return;
      addDataset({ name: opened.file.name, kind: opened.handle ? "file" : "upload", text, baseline: text, handle: opened.handle, lastModified: opened.handle ? opened.file.lastModified : undefined });
      setErrors([]);
    } catch (error) { if (reset === resetRevision.current) fileError(error, "Dataset could not be opened"); }
  }

  function chooseDataFile() {
    reopenDataId.current = undefined;
    if (files.canOpen) void openData();
    else dataInput.current?.click();
  }

  async function saveData(saveAs = false) {
    const snapshot = datasetRecord.current;
    if (!snapshot || snapshot.restoreError || savingRef.current) return;
    if (!files.canSave) {
      downloadBlob(new Blob([snapshot.text], { type: "application/json" }), snapshot.name);
      return;
    }
    savingRef.current = true;
    setSaving(true);
    try {
      const handle = !saveAs && snapshot.handle ? snapshot.handle : await files.destination(snapshot.name);
      if (await files.sameFile(snapshot.handle, handle) && await files.hasChanged(handle, snapshot.lastModified)) {
        if (!await modal.confirm({ title: "File changed on disk", content: `${snapshot.name} has been modified outside MonoReport. Overwrite it with this dataset?`, okText: "Overwrite", cancelText: "Cancel" })) return;
      }
      const lastModified = await files.write(handle, snapshot.text);
      const current = datasetRecord.current;
      if (current?.id === snapshot.id) {
        setDatasetRecord({ ...current, kind: "file", name: handle.name, handle, baseline: snapshot.text, lastModified });
      }
      void feedback.success("Saved.");
    } catch (error) { fileError(error, "Dataset could not be saved"); }
    finally { savingRef.current = false; setSaving(false); }
  }

  async function recoverData(record: Dataset) {
    if (!record.needsPermission || !record.handle?.requestPermission) {
      if (files.canOpen) await openData(undefined, record.id);
      else { reopenDataId.current = record.id; dataInput.current?.click(); }
      return;
    }
    const revision = contentRevision.current;
    try {
      const file = await files.reopen(record.handle, true);
      if (revision === contentRevision.current && datasetRecord.current === record) await openData(file, record.id, record.handle);
    } catch (error) { if (revision === contentRevision.current) fileError(error, "Dataset could not be reopened"); }
  }

  return { openData, chooseDataFile, saveData, recoverData, saving, savingRef, dataInput, reopenDataId };
}
