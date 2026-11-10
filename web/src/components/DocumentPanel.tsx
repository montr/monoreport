import { Alert, Button, Dropdown, Empty, Flex, Radio, Spin, Tag, Tooltip, Upload } from "antd";
import { EllipsisOutlined, ExclamationCircleOutlined, UploadOutlined } from "@ant-design/icons";
import type { DragEvent } from "react";
import { XlsxPreview } from "./XlsxPreview";
import { DocxPreview } from "./DocxPreview";
import { templateAccept, templateFormat } from "../services/documentFormat";
import { workspaceDropdownProps } from "./workspaceDropdownProps";
import { PanelFilename } from "./PanelFilename";
import type { Workspace } from "../hooks/useWorkspace";

export function DocumentPanel({ model }: { model: Workspace["documentPanel"] }) {
  const { liveReloadEnabled, sourceStatus, active, setActive, result, selectedTemplate, resultInputs, pending, exampleLoading, resultCurrent, dragging, setDragging, recoverTemplate, drop, files, openTemplate, openLabel, acceptTemplate, template, downloadBlob, clearTemplate, closeResult, fileInput, selected } = model;
  const source = active === "result" ? result?.blob : template;
  const format = active === "result" ? result?.format : template && templateFormat(template.name);
  const label = active === "result" ? "Result" : "Template";
  function getDocumentMenuItems() {
    return [
      ...(active === "template"
        ? [{ key: "upload", label: openLabel, onClick: () => void openTemplate() }]
        : []),
      {
        key: "download",
        label: "Download",
        disabled: !source,
        onClick: () => {
          if (active === "result" && result) downloadBlob(result.blob, result.filename);
          else if (template) downloadBlob(template, template.name);
        }
      },
      { type: "divider" as const, key: "file-actions-end" },
      {
        key: "close",
        label: "Close",
        disabled: active === "template" ? !selectedTemplate : !source,
        onClick: active === "template" ? clearTemplate : closeResult
      }
    ].map(item => "type" in item ? item : { ...item, icon: <span aria-hidden="true" /> });
  }

  function handleDragEnter(event: DragEvent<HTMLDivElement>) {
    if ((source || active === "result") && event.dataTransfer.types.includes("Files")) {
      event.preventDefault();
      setDragging(true);
    }
  }

  function handleBeforeUpload(file: File) {
    setDragging(false);
    acceptTemplate(file);
    return false;
  }

  return (
    <section className="preview" aria-label="Preview">
      <input ref={fileInput} type="file" hidden accept={templateAccept} onChange={selected} />
      <Flex className="pane-header" align="center" gap="small" wrap>
        <Radio.Group
          size="small"
          value={active}
          onChange={event => setActive(event.target.value)}
          aria-label="Document view"
        >
          <Radio.Button value="template">Template</Radio.Button>
          <Radio.Button value="result">Result</Radio.Button>
        </Radio.Group>
        <Flex className="context-name" align="center" gap="small" wrap={false}>
          <PanelFilename
            name={active === "result" ? result?.filename : selectedTemplate?.name}
            source={active === "result"
              ? resultInputs?.templateSource
              : selectedTemplate?.exampleId
                ? "Example file"
                : "Local file"}
          />
        </Flex>
        {liveReloadEnabled && selectedTemplate && (
          <Tag role="note" aria-label="Template Live Reload" color={sourceStatus?.error ? "warning" : undefined}>
            {sourceStatus?.error ? "Template unavailable" : !selectedTemplate.handle ? "Template snapshot" : sourceStatus ? "Template monitored" : "Reading Template…"}
          </Tag>
        )}
        {(pending || exampleLoading) && (
          <Spin
            size="small"
            aria-label={exampleLoading
              ? "Loading example"
              : pending === "generate"
                ? "Generating document"
                : pending === "analyze"
                  ? "Analyzing template"
                  : "Generating Dataset"}
          />
        )}
        {result && active === "result" && !resultCurrent && (
          <Tooltip title="The Dataset or Template changed since this Result was generated.">
            <Tag
              color="warning"
              variant="outlined"
              icon={<ExclamationCircleOutlined />}
              style={{ flex: "none" }}
            >
              Outdated
            </Tag>
          </Tooltip>
        )}
        <Dropdown
          {...workspaceDropdownProps}
          trigger={["click"]}
          menu={{ items: getDocumentMenuItems() }}
        >
          <Button
            size="small"
            type="text"
            icon={<EllipsisOutlined />}
            aria-label="Document menu"
          />
        </Dropdown>
      </Flex>
      {selectedTemplate && (selectedTemplate.restoreError || sourceStatus?.error) && (
        <Alert
          type="warning"
          title={`${selectedTemplate.name}: ${selectedTemplate.restoreError ?? sourceStatus?.error}`}
          action={(
            <Button size="small" onClick={() => void recoverTemplate(selectedTemplate)}>
              {selectedTemplate.needsPermission && selectedTemplate.handle?.requestPermission
                ? "Grant access"
                : "Reopen…"}
            </Button>
          )}
        />
      )}
      <div
        className={`canvas ${format === "xlsx" ? "xlsx-view" : ""} ${dragging ? "drop-target" : ""}`}
        onDragEnter={handleDragEnter}
        onDragOver={event => {
          if (event.dataTransfer.types.includes("Files")) event.preventDefault();
        }}
        onDragLeave={event => {
          if (!event.currentTarget.contains(event.relatedTarget as Node)) setDragging(false);
        }}
        onDropCapture={event => {
          event.stopPropagation();
          drop(event);
        }}
      >
        {source
          ? format === "xlsx"
            ? <XlsxPreview source={source} label={label} sourceId={active === "template" ? selectedTemplate?.id : "result"} />
            : <DocxPreview source={source} label={label} sourceId={active === "template" ? selectedTemplate?.id : "result"} />
          : active === "result"
            ? <Empty description="Not generated" />
            : files.canOpen
              ? (
                <Empty description="Select a DOCX or XLSX template">
                  <Button onClick={() => void openTemplate()}>{openLabel}</Button>
                </Empty>
              )
              : (
                <Upload.Dragger
                  accept={templateAccept}
                  showUploadList={false}
                  onDrop={event => event.stopPropagation()}
                  beforeUpload={handleBeforeUpload}
                >
                  <Empty description="Select a DOCX or XLSX template">
                    <Button icon={<UploadOutlined aria-hidden="true" />}>Upload…</Button>
                  </Empty>
                </Upload.Dragger>
              )}
      </div>
    </section>
  );
}
