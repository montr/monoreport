import { useEffect, useRef, useState, type RefObject } from "react";
import type { LocalFileHandle, LocalFiles, FileRead } from "../services/localFiles";
import type { Dataset, TemplateEntry } from "../types";

const generationDebounceMs = 500;
const pollingIntervalMs = 1000;
type Source = "template" | "dataset";
export interface SourceStatus { id: string; handle?: LocalFileHandle; error?: string; diskText?: string; }
interface PendingRead { promise: Promise<FileRead>; key: string; owned: boolean; }
interface Props {
  files: LocalFiles;
  template: TemplateEntry | undefined;
  templateRecord: RefObject<TemplateEntry | undefined>;
  dataset: Dataset;
  datasetRecord: RefObject<Dataset>;
  saving: boolean;
  savingRef: RefObject<boolean>;
  saveRevision: RefObject<number>;
  contentRevision: RefObject<number>;
  busy: boolean;
  applyTemplate(file: File, owned: boolean): void;
  applyDataset(data: Dataset): void;
  generate(automatic?: boolean): Promise<unknown>;
  cancelAutomatic(): void;
  invalidateGeneration(): void;
  retryAutomatic(readRecovered?: boolean): void;
}

export function useLiveReload(props: Props) {
  const current = useRef(props);
  current.current = props;
  const [enabled, setEnabled] = useState(false);
  const enabledRef = useRef(false);
  const mounted = useRef(true);
  const epoch = useRef(0);
  const reads = useRef(new Map<LocalFileHandle, PendingRead>());
  const tasks = useRef(new Map<Source, Promise<void>>());
  const [statuses, setStatuses] = useState<Partial<Record<Source, SourceStatus>>>({});
  const statusRef = useRef(statuses);
  const [returnRevision, setReturnRevision] = useState(0);
  const [reconciledRevision, setReconciledRevision] = useState(0);
  const restoration = useRef<{ promise: Promise<void>; resolve(): void } | undefined>(undefined);

  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      epoch.current++;
      restoration.current?.resolve();
    };
  }, []);

  function status(source: Source, value: SourceStatus) {
    if (!mounted.current) return;
    const previous = statusRef.current[source];
    if (previous?.id === value.id && previous.handle === value.handle && previous.error === value.error && previous.diskText === value.diskText) return;
    statusRef.current = { ...statusRef.current, [source]: value };
    setStatuses(statusRef.current);
  }

  useEffect(() => {
    let next = statusRef.current;
    for (const source of ["template", "dataset"] as const) {
      const binding = source === "template" ? props.template : props.dataset;
      const previous = next[source];
      if (previous && (previous.id !== binding?.id || previous.handle !== binding?.handle)) {
        next = { ...next };
        delete next[source];
      }
    }
    if (next !== statusRef.current) { statusRef.current = next; setStatuses(next); }
  }, [props.template?.id, props.template?.handle, props.dataset.id, props.dataset.handle]);

  async function read(handle: LocalFileHandle, key: string, owned: boolean, valid: () => boolean): Promise<PendingRead> {
    const previous = reads.current.get(handle);
    if (previous) {
      if (previous.key === key) { previous.owned ||= owned; return previous; }
      // Native calls cannot be aborted. An obsolete binding must settle first.
      await previous.promise.catch(() => undefined);
      if (!valid()) throw new DOMException("Source binding changed.", "AbortError");
      return read(handle, key, owned, valid);
    }
    const pending = { promise: current.current.files.readFile(handle), key, owned };
    reads.current.set(handle, pending);
    void pending.promise.finally(() => {
      if (reads.current.get(handle) === pending) reads.current.delete(handle);
    }).catch(() => undefined);
    return pending;
  }

  async function refreshTemplate(owned = false, watching?: () => boolean) {
    const p = current.current;
    const entry = p.templateRecord.current;
    if (!entry?.handle) return;
    const revision = p.contentRevision.current;
    const expectedEpoch = epoch.current;
    const bound = () => mounted.current && (!watching || watching()) &&
      current.current.templateRecord.current?.id === entry.id && current.current.templateRecord.current?.handle === entry.handle;
    const valid = () => bound() && epoch.current === expectedEpoch;
    let pending: PendingRead | undefined;
    try {
      pending = await read(entry.handle, `${entry.id}:${expectedEpoch}`, owned, valid);
      const candidate = await pending.promise;
      if (owned && bound() && epoch.current !== expectedEpoch) return refreshTemplate(owned, watching);
      if (!valid() || (owned && revision !== p.contentRevision.current)) throw new DOMException("Template selection changed during reading.", "AbortError");
      const latest = current.current.templateRecord.current!;
      // An explicit preparation sharing this sample owns its revision commit.
      if (owned || !pending.owned) {
        const file = p.files.acceptTemplateRead(candidate, latest.file);
        if (file !== latest.file) current.current.applyTemplate(file, owned);
      }
      status("template", { id: entry.id, handle: entry.handle });
      current.current.retryAutomatic(true);
    } catch (error) {
      if (owned && bound() && epoch.current !== expectedEpoch) return refreshTemplate(owned, watching);
      if (valid() && !(error instanceof DOMException && error.name === "AbortError")) {
        status("template", { id: entry.id, handle: entry.handle, error: "Template unavailable; retrying while Live Reload is enabled. Use Reopen if access does not recover." });
        if (!owned && !pending?.owned) current.current.invalidateGeneration();
      }
      throw error;
    }
  }

  async function refreshDataset(valid: () => boolean) {
    const p = current.current;
    const entry = p.datasetRecord.current;
    const saveRevision = p.saveRevision.current;
    if (!entry.handle || p.savingRef.current || entry.restoreError) return;
    try {
      const pending = await read(entry.handle, `${entry.id}:${epoch.current}:${saveRevision}`, false, valid);
      const candidate = await pending.promise;
      if (!valid() || current.current.savingRef.current || saveRevision !== current.current.saveRevision.current) return;
      const latest = current.current.datasetRecord.current;
      const disk = p.files.inspectDatasetRead(candidate);
      status("dataset", { id: entry.id, handle: entry.handle, diskText: disk.text,
        error: disk.valid ? undefined : "Disk JSON is invalid; editor text was retained." });
      if (!disk.valid) return;
      if (latest.text === latest.baseline || disk.text === latest.text) {
        if (latest.text !== disk.text || latest.baseline !== disk.text || latest.lastModified !== candidate.file.lastModified) {
          current.current.applyDataset({ ...latest, text: disk.text, baseline: disk.text, lastModified: candidate.file.lastModified });
        }
      } else if (disk.text === latest.baseline && latest.lastModified !== candidate.file.lastModified) {
        current.current.applyDataset({ ...latest, lastModified: candidate.file.lastModified });
      }
    } catch (error) {
      if (valid()) status("dataset", { id: entry.id, handle: entry.handle, error: "Dataset unavailable; editor text was retained. Use Reopen if access does not recover." });
    }
  }

  function setLiveReloadEnabled(value: boolean) {
    if (value === enabledRef.current) return;
    enabledRef.current = value;
    // Disabling polling must not obsolete a separate explicit preparation.
    if (!value) current.current.cancelAutomatic();
    else {
      epoch.current++;
      statusRef.current = {};
      setStatuses({});
      current.current.retryAutomatic();
    }
    restoration.current?.resolve();
    restoration.current = undefined;
    setEnabled(value);
  }

  useEffect(() => {
    if (!enabled) return;
    let active = true;
    const expectedEpoch = epoch.current;
    const ownedTasks = new Map<Source, Promise<void>>();
    const template = props.template;
    const dataset = props.dataset;
    const valid = (source: Source) => active && mounted.current && enabledRef.current && epoch.current === expectedEpoch && (source === "template"
      ? current.current.templateRecord.current?.id === template?.id && current.current.templateRecord.current?.handle === template?.handle
      : current.current.datasetRecord.current.id === dataset.id && current.current.datasetRecord.current.handle === dataset.handle);
    function sample(source: Source) {
      if (!valid(source)) return Promise.resolve();
      const existing = tasks.current.get(source);
      // Busy ticks do not queue; a new binding's initial read waits for old work.
      if (existing) return existing;
      const task = (source === "template" ? refreshTemplate(false, () => valid(source)) : refreshDataset(() => valid(source))).catch(() => undefined);
      tasks.current.set(source, task);
      ownedTasks.set(source, task);
      void task.finally(() => { if (tasks.current.get(source) === task) tasks.current.delete(source); });
      return task;
    }
    // Wait for obsolete samples before reconciling this binding, without overlap.
    const initial = Promise.all((["template", "dataset"] as const).map(async source => {
      await tasks.current.get(source);
      if (active) await sample(source);
    }));
    void initial.finally(() => {
      if (active && restoration.current) {
        restoration.current.resolve();
        restoration.current = undefined;
        setReconciledRevision(value => value + 1);
      }
    });
    const timer = window.setInterval(() => { void sample("template"); void sample("dataset"); }, pollingIntervalMs);
    return () => {
      active = false;
      window.clearInterval(timer);
      for (const [source, task] of ownedTasks) if (tasks.current.get(source) === task) tasks.current.delete(source);
    };
  }, [enabled, props.template?.id, props.template?.handle, props.dataset.id, props.dataset.handle, props.saving, returnRevision]);

  useEffect(() => {
    if (!enabled) return;
    function restore() {
      if (document.visibilityState === "hidden" || restoration.current) return;
      epoch.current++;
      statusRef.current = {};
      setStatuses({});
      current.current.cancelAutomatic();
      let resolve!: () => void;
      const promise = new Promise<void>(done => { resolve = done; });
      restoration.current = { promise, resolve };
      setReturnRevision(value => value + 1);
    }
    document.addEventListener("visibilitychange", restore);
    window.addEventListener("pageshow", restore);
    return () => {
      document.removeEventListener("visibilitychange", restore);
      window.removeEventListener("pageshow", restore);
    };
  }, [enabled]);

  async function reconcileInputs() {
    await restoration.current?.promise;
    await Promise.all([...tasks.current.values()]);
    const entry = current.current.templateRecord.current;
    const state = statusRef.current.template;
    if (entry?.handle && state?.id === entry.id && state.handle === entry.handle && state.error) throw new Error(state.error);
  }

  const templateStatus = statuses.template?.id === props.template?.id && statuses.template?.handle === props.template?.handle ? statuses.template : undefined;
  const datasetStatus = statuses.dataset?.id === props.dataset.id && statuses.dataset?.handle === props.dataset.handle ? statuses.dataset : undefined;
  useEffect(() => {
    if (!enabled || props.busy || !props.template?.file || props.dataset.restoreError || templateStatus?.error || restoration.current) return;
    try { JSON.parse(props.dataset.text); } catch { return; }
    const timer = window.setTimeout(() => { void current.current.generate(true); }, generationDebounceMs);
    return () => window.clearTimeout(timer);
  }, [enabled, props.busy, props.template?.file, props.dataset.id, props.dataset.text, props.dataset.restoreError, templateStatus?.error, returnRevision, reconciledRevision]);

  return { liveReloadEnabled: enabled, setLiveReloadEnabled, refreshTemplate, reconcileInputs, templateStatus, datasetStatus };
}
