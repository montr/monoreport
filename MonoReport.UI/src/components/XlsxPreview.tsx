import { useEffect, useRef, useState } from "react";
import type { XlsxViewer } from "@silurus/ooxml/xlsx";

interface Props {
  source: Blob | File;
  label: string;
}

type PreviewStatus = { kind: "loading" | "error"; message: string };
type FormulaBar = { address: string; text: string };

export function XlsxPreview({ source, label }: Props) {
  const host = useRef<HTMLDivElement>(null);
  const [status, setStatus] = useState<PreviewStatus>();
  const [formulaBar, setFormulaBar] = useState<FormulaBar>();

  useEffect(() => {
    const container = host.current;
    if (!container) return;
    let current = true;
    let viewer: XlsxViewer | undefined;
    const sessionHost = document.createElement("div");
    sessionHost.className = "xlsx-preview-session";
    container.replaceChildren(sessionHost);
    setFormulaBar(undefined);
    setStatus({ kind: "loading", message: `Loading ${label.toLowerCase()} preview…` });

    function dispose() {
      const owned = viewer;
      viewer = undefined;
      owned?.destroy();
      sessionHost.remove();
    }

    function failed() {
      if (!current) return;
      current = false;
      setFormulaBar(undefined);
      setStatus({ kind: "error", message: `${label} XLSX preview could not be rendered. The workbook may be unsupported or exceed preview limits. You can still use the original file.` });
      dispose();
    }

    void (async () => {
      const { XlsxViewer } = await import("@silurus/ooxml/xlsx");
      if (!current) return;
      const bytes = await source.arrayBuffer();
      if (!current) return;
      viewer = new XlsxViewer(sessionHost, {
        mode: "worker",
        hiddenSheetMode: "skip",
        useGoogleFonts: false,
        enableHyperlinks: false,
        resizable: false,
        workerTimeoutMs: 30_000,
        onSelectionContextChange: context => {
          if (!current) return;
          if (context?.kind !== "range") { setFormulaBar(undefined); return; }
          const active = context.selection.activeCell;
          let column = "";
          for (let col = active.col; col > 0; col = Math.floor((col - 1) / 26)) {
            column = String.fromCharCode(65 + (col - 1) % 26) + column;
          }
          const cell = context.cells.find(cell => cell.address.row === active.row && cell.address.col === active.col);
          setFormulaBar({
            address: `${column}${active.row}`,
            text: cell ? cell.formula !== undefined ? `=${cell.formula.replace(/^=+/, "")}` : cell.displayText
              : ""
          });
        },
        onError: failed
      });
      // Self-owned loading lets destroy() release the workbook and its workers.
      if (!current) { dispose(); return; }
      await viewer.load(bytes);
      if (current) setStatus(undefined);
    })().catch(failed);

    return () => {
      current = false;
      dispose();
    };
  }, [source, label]);

  return <div className="xlsx-preview">
    {status && <div className={status.kind === "error" ? "preview-warning" : "preview-loading"} role="status">{status.message}</div>}
    <div className="xlsx-formula-bar" role="group" aria-label={`${label} formula bar`}>
      <span className="xlsx-cell-address" aria-label={`${label} active cell`}>{formulaBar?.address ?? ""}</span>
      <span className="xlsx-formula-label" aria-hidden="true">fx</span>
      <textarea rows={1} readOnly aria-label={`${label} formula or value`} value={formulaBar?.text ?? ""} />
    </div>
    <div className="xlsx-preview-host" ref={host} data-testid={`${label.toLowerCase()}-xlsx-preview`} />
  </div>;
}
