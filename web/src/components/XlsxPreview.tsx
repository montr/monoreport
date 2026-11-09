import { Alert, Flex, Input, Select, Spin, Typography } from "antd";
import { useXlsxPreview } from "../hooks/useXlsxPreview";

interface Props {
  source: Blob | File;
  label: string;
}

export function XlsxPreview({ source, label }: Props) {
  const { host, status, formulaBar, namedRanges, navigateToName } = useXlsxPreview(source, label);

  return (
    <div className="xlsx-preview">
      {status && (
        status.kind === "error"
          ? <Alert type="warning" showIcon role="status" title={status.message} />
          : (
            <Flex align="center" gap="small" role="status" className="preview-loading">
              <Spin size="small" />
              <Typography.Text type="secondary">{status.message}</Typography.Text>
            </Flex>
          )
      )}
      <div className="xlsx-formula-bar" role="group" aria-label={`${label} formula bar`}>
        <Select<string>
          className="xlsx-name-box"
          size="small"
          aria-label={`${label} name box`}
          showSearch={false}
          value={null}
          disabled={namedRanges.length === 0}
          placeholder={<span aria-label={`${label} active cell`}>{formulaBar?.address ?? ""}</span>}
          options={namedRanges.map(range => ({ value: range.id, label: range.name }))}
          onSelect={navigateToName}
        />
        <span className="xlsx-formula-label" aria-hidden="true">fx</span>
        <Input.TextArea
          rows={1}
          readOnly
          wrap="off"
          aria-label={`${label} formula or value`}
          value={formulaBar?.text ?? ""}
          styles={{
            root: { flex: 1, minWidth: 0 },
            textarea: { whiteSpace: "pre", resize: "none", overflow: "auto" }
          }}
        />
      </div>
      <div
        className="xlsx-preview-host"
        ref={host}
        data-testid={`${label.toLowerCase()}-xlsx-preview`}
      />
    </div>
  );
}
