import { useEffect, useRef, useState } from "react";
import type { XlsxSelectionContext, XlsxViewer, XlsxWorkbook } from "@silurus/ooxml/xlsx";
import { XlsxNamedRanges, type XlsxNamedRange } from "../services/xlsxNamedRanges";

type PreviewStatus = { kind: "loading" | "error"; message: string };
type FormulaBar = { address: string; text: string };

export function useXlsxPreview(source: Blob | File, label: string) {
  const host = useRef<HTMLDivElement>(null);
  const navigation = useRef<((id: string) => Promise<void>) | undefined>(undefined);
  const [status, setStatus] = useState<PreviewStatus>();
  const [formulaBar, setFormulaBar] = useState<FormulaBar>();
  const [namedRanges, setNamedRanges] = useState<XlsxNamedRange[]>([]);

  useEffect(() => {
    const container = host.current;
    if (!container) return;
    let current = true;
    let constructingViewer = false;
    let viewer: Omit<XlsxViewer, "load"> | undefined;
    let workbook: XlsxWorkbook | undefined;
    let names: XlsxNamedRanges | undefined;
    let catalog: XlsxNamedRange[] = [];
    let catalogRequest = 0;
    let navigationRequest = 0;
    let destination: number | undefined;
    const sessionHost = document.createElement("div");
    sessionHost.className = "xlsx-preview-session";
    container.replaceChildren(sessionHost);
    setFormulaBar(undefined);
    setNamedRanges([]);
    setStatus({ kind: "loading", message: `Loading ${label.toLowerCase()} preview…` });

    function dispose() {
      navigation.current = undefined;
      if (constructingViewer) return;
      const ownedViewer = viewer;
      const ownedWorkbook = workbook;
      viewer = undefined;
      workbook = undefined;
      try { ownedViewer?.destroy(); }
      finally {
        ownedWorkbook?.destroy();
        sessionHost.remove();
      }
    }

    function failed() {
      if (!current) return;
      current = false;
      setFormulaBar(undefined);
      setNamedRanges([]);
      setStatus({ kind: "error", message: `${label} XLSX preview could not be rendered. The workbook may be unsupported or exceed preview limits. You can still use the original file.` });
      dispose();
    }

    function handleSelectionContextChange(context: XlsxSelectionContext | null) {
      if (!current) return;
      if (context?.kind !== "range") {
        setFormulaBar(undefined);
        return;
      }
      const active = context.selection.activeCell;
      let column = "";
      for (let col = active.col; col > 0; col = Math.floor((col - 1) / 26)) {
        column = String.fromCharCode(65 + (col - 1) % 26) + column;
      }
      const cell = context.cells.find(cell => cell.address.row === active.row && cell.address.col === active.col);
      setFormulaBar({
        address: `${column}${active.row}`,
        text: cell ? cell.formula !== undefined ? `=${cell.formula.replace(/^=+/, "")}` : cell.displayText : ""
      });
    }

    function handleSheetChange(index: number) {
      if (!current) return;
      if (destination !== index) navigationRequest++;
      const request = ++catalogRequest;
      catalog = [];
      setNamedRanges([]);
      setStatus(undefined);
      void names!.getForSheet(index).then(ranges => {
        if (!current || request !== catalogRequest || viewer?.sheetIndex !== index) return;
        catalog = ranges;
        setNamedRanges(ranges);
      }).catch(() => {
        if (current && request === catalogRequest && viewer?.sheetIndex === index) failed();
      });
    }

    navigation.current = async id => {
      const range = catalog.find(item => item.id === id);
      const activeViewer = viewer;
      if (!current || !range || !activeViewer) return;
      const request = ++navigationRequest;
      destination = range.sheetIndex;
      try {
        if (activeViewer.sheetIndex !== range.sheetIndex) await activeViewer.goToSheet(range.sheetIndex);
        if (!current || request !== navigationRequest || activeViewer.sheetIndex !== range.sheetIndex) return;
        await activeViewer.scrollToCell(range.firstCell, { align: "nearest" });
        if (!current || request !== navigationRequest || activeViewer.sheetIndex !== range.sheetIndex) return;
        destination = undefined;
        activeViewer.setSelection(range.range);
      } catch {
        if (current && request === navigationRequest) failed();
      } finally {
        if (request === navigationRequest) destination = undefined;
      }
    };

    void (async () => {
      const { XlsxViewer, XlsxWorkbook } = await import("@silurus/ooxml/xlsx");
      if (!current) return;
      const bytes = await source.arrayBuffer();
      if (!current) return;
      const loaded = await XlsxWorkbook.load(bytes, {
        mode: "worker", useGoogleFonts: false, workerTimeoutMs: 30_000
      });
      if (!current) { loaded.destroy(); return; }
      workbook = loaded;
      names = new XlsxNamedRanges(loaded);
      constructingViewer = true;
      try {
        viewer = XlsxViewer.fromWorkbook(sessionHost, loaded, {
          hiddenSheetMode: "skip",
          enableHyperlinks: false,
          resizable: false,
          onSheetChange: handleSheetChange,
          onSelectionStateChange: selection => {
            if (current && selection) { navigationRequest++; destination = undefined; }
          },
          onSelectionContextChange: handleSelectionContextChange,
          onError: failed
        });
      } finally {
        constructingViewer = false;
        // A synchronous error callback must not release the borrowed workbook
        // until construction returns and its viewer can be destroyed first.
        if (!current) dispose();
      }
    })().catch(failed);

    return () => {
      current = false;
      dispose();
    };
  }, [source, label]);

  return { host, status, formulaBar, namedRanges, navigateToName: (id: string) => { void navigation.current?.(id); } };
}
