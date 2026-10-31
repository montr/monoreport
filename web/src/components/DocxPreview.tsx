import { useEffect, useRef, useState } from "react";
import { renderAsync } from "docx-preview";

interface Props {
  source: Blob | File;
  label: string;
}

export function DocxPreview({ source, label }: Props) {
  const host = useRef<HTMLDivElement>(null);
  const [warning, setWarning] = useState<string>();

  useEffect(() => {
    const container = host.current;
    if (!container) return;
    let current = true;
    const renderHost = document.createElement("div");
    container.replaceChildren(renderHost);
    setWarning(undefined);
    void renderAsync(source, renderHost, undefined, { inWrapper: true, ignoreWidth: false, ignoreHeight: false })
      .catch(() => { if (current) setWarning(`${label} preview could not be rendered.`); });
    return () => {
      current = false;
      renderHost.remove();
    };
  }, [source, label]);

  return (
    <div className="document-preview">
      {warning && <div className="preview-warning" role="status">{warning}</div>}
      <div ref={host} data-testid={`${label.toLowerCase()}-preview`} />
    </div>
  );
}
