import { useEffect, useRef, useState } from "react";
import type { XlsxViewer } from "@silurus/ooxml/xlsx";

interface Props {
  source: Blob | File;
  label: string;
}

export function XlsxPreview({ source, label }: Props) {
  const host = useRef<HTMLDivElement>(null);
  const [status, setStatus] = useState<string>();

  useEffect(() => {
    const container = host.current;
    if (!container) return;
    let current = true;
    let viewer: XlsxViewer | undefined;
    const sessionHost = document.createElement("div");
    sessionHost.className = "xlsx-preview-session";
    container.replaceChildren(sessionHost);
    setStatus(`Loading ${label.toLowerCase()} preview…`);

    function dispose() {
      const owned = viewer;
      viewer = undefined;
      owned?.destroy();
      sessionHost.remove();
    }

    function failed() {
      if (!current) return;
      current = false;
      setStatus(`${label} XLSX preview could not be rendered. The workbook may be unsupported or exceed preview limits. You can still use the original file.`);
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
    {status && <div className="preview-warning" role="status">{status}</div>}
    <div className="xlsx-preview-host" ref={host} data-testid={`${label.toLowerCase()}-xlsx-preview`} />
  </div>;
}
