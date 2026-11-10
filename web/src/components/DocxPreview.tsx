import { Alert } from "antd";
import { useEffect, useRef, useState } from "react";
import { renderAsync } from "docx-preview";

interface Props {
  source: Blob | File;
  label: string;
  sourceId?: string;
}

export function DocxPreview({ source, label, sourceId = label }: Props) {
  const host = useRef<HTMLDivElement>(null);
  const position = useRef<{ id: string; top: number; left: number } | undefined>(undefined);
  const [warning, setWarning] = useState<string>();

  useEffect(() => {
    const container = host.current;
    if (!container) return;
    let current = true;
    const canvas = container.closest<HTMLElement>(".canvas");
    const saved = position.current?.id === sourceId ? position.current : undefined;
    const renderHost = document.createElement("div");
    container.replaceChildren(renderHost);
    setWarning(undefined);
    void renderAsync(source, renderHost, undefined, { inWrapper: true, ignoreWidth: false, ignoreHeight: false, renderAltChunks: false })
      .then(() => {
        if (!current || !canvas) return;
        canvas.scrollTop = Math.min(saved?.top ?? 0, Math.max(0, canvas.scrollHeight - canvas.clientHeight));
        canvas.scrollLeft = Math.min(saved?.left ?? 0, Math.max(0, canvas.scrollWidth - canvas.clientWidth));
      })
      .catch(() => { if (current) setWarning(`${label} preview could not be rendered.`); });
    return () => {
      current = false;
      if (canvas) position.current = { id: sourceId, top: canvas.scrollTop, left: canvas.scrollLeft };
      renderHost.remove();
    };
  }, [source, label, sourceId]);

  return (
    <div className="document-preview">
      {warning && (
        <Alert type="warning" showIcon role="status" title={warning} />
      )}
      <div
        ref={host}
        data-testid={`${label.toLowerCase()}-preview`}
      />
    </div>
  );
}
