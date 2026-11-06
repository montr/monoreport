import { Splitter } from "antd";
import { ReactNode, useEffect, useRef, useState } from "react";
import { workspaceStorage } from "../services/workspacePersistence";

type PanelState = [boolean, boolean];

function readRatio(key: string, fallback: number) {
  let value: number;
  try { value = Number(localStorage.getItem(key)); }
  catch { return fallback; }
  return Number.isFinite(value) && value > 0 && value < 100 ? value : fallback;
}

function readCollapsed(key: string): PanelState {
  try {
    const value: unknown = JSON.parse(localStorage.getItem(key) ?? "null");
    if (Array.isArray(value) && value.length === 2 && value.every(item => typeof item === "boolean")) return value as PanelState;
  } catch { /* Use the expanded defaults when preference storage is unavailable or invalid. */ }
  return [false, false];
}

function writePreference(key: string, value: string) {
  try { localStorage.setItem(key, value); }
  catch { /* Keep native resizing usable when preference storage is unavailable. */ }
}

interface Props {
  data: ReactNode;
  document: ReactNode;
  problems: ReactNode;
  problemsOpen: boolean;
  onProblemsOpenChange(open: boolean): void;
}

export function WorkspaceLayout({ data, document, problems, problemsOpen, onProblemsOpenChange }: Props) {
  const [mobile, setMobile] = useState(() => window.matchMedia("(max-width: 850px)").matches);
  const [value, setValue] = useState(() => readRatio(mobile ? workspaceStorage.mobile : workspaceStorage.desktop, mobile ? 50 : 45));
  const [collapsed, setCollapsed] = useState<PanelState>(() => readCollapsed(mobile ? workspaceStorage.mobileCollapsed : workspaceStorage.desktopCollapsed));
  const [problemsSize, setProblemsSize] = useState(() => readRatio(workspaceStorage.diagnosticsSize, 25));
  const mainCollapsed = useRef<PanelState>(collapsed);
  const problemsCollapsed = useRef(!problemsOpen);

  useEffect(() => {
    const media = window.matchMedia("(max-width: 850px)");
    const update = () => {
      setMobile(media.matches);
      setValue(readRatio(media.matches ? workspaceStorage.mobile : workspaceStorage.desktop, media.matches ? 50 : 45));
      const restored = readCollapsed(media.matches ? workspaceStorage.mobileCollapsed : workspaceStorage.desktopCollapsed);
      mainCollapsed.current = restored;
      setCollapsed(restored);
    };
    media.addEventListener("change", update);
    return () => media.removeEventListener("change", update);
  }, []);

  useEffect(() => { problemsCollapsed.current = !problemsOpen; }, [problemsOpen]);

  function updateCollapsed(next: boolean[]) {
    const state: PanelState = [Boolean(next[0]), Boolean(next[1])];
    mainCollapsed.current = state;
    setCollapsed(state);
    writePreference(mobile ? workspaceStorage.mobileCollapsed : workspaceStorage.desktopCollapsed, JSON.stringify(state));
  }

  function updateProblemsSize(sizes: number[]) {
    const total = sizes[0] + sizes[1];
    if (!problemsCollapsed.current && total && sizes[0] > 0 && sizes[1] > 0) {
      setProblemsSize(sizes[1] / total * 100);
    }
  }

  function saveProblemsSize(sizes: number[]) {
    const total = sizes[0] + sizes[1];
    if (!problemsCollapsed.current && total && sizes[0] > 0 && sizes[1] > 0) {
      writePreference(workspaceStorage.diagnosticsSize, String(sizes[1] / total * 100));
    }
  }

  function updateMainSize(sizes: number[]) {
    const total = sizes[0] + sizes[1];
    if (!mainCollapsed.current.some(Boolean) && total && sizes[0] > 0 && sizes[1] > 0) {
      setValue(sizes[0] / total * 100);
    }
  }

  function saveMainSize(sizes: number[]) {
    const total = sizes[0] + sizes[1];
    if (!mainCollapsed.current.some(Boolean) && total && sizes[0] > 0 && sizes[1] > 0) {
      writePreference(
        mobile ? workspaceStorage.mobile : workspaceStorage.desktop,
        String(sizes[0] / total * 100)
      );
    }
  }

  const mainSizes = collapsed[0] ? ["0%", "100%"] : collapsed[1] ? ["100%", "0%"] : [`${value}%`, `${100 - value}%`];
  return (
    <Splitter
      orientation="vertical"
      onResize={updateProblemsSize}
      onResizeEnd={saveProblemsSize}
      onCollapse={next => {
        problemsCollapsed.current = next[1];
        onProblemsOpenChange(!next[1]);
      }}
    >
      <Splitter.Panel size={problemsOpen ? `${100 - problemsSize}%` : "100%"}>
        <Splitter
          orientation={mobile ? "vertical" : "horizontal"}
          onResize={updateMainSize}
          onResizeEnd={saveMainSize}
          onCollapse={updateCollapsed}
        >
          <Splitter.Panel size={mainSizes[0]} collapsible>
            {data}
          </Splitter.Panel>
          <Splitter.Panel size={mainSizes[1]} collapsible>
            {document}
          </Splitter.Panel>
        </Splitter>
      </Splitter.Panel>
      <Splitter.Panel size={problemsOpen ? `${problemsSize}%` : 0} collapsible={{ start: true }}>
        {problems}
      </Splitter.Panel>
    </Splitter>
  );
}
