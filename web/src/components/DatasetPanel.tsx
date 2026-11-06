import { Alert, Badge, Button, Dropdown, Flex, Typography } from "antd";
import { CheckOutlined, EllipsisOutlined } from "@ant-design/icons";
import type { ChangeEvent } from "react";
import { MonacoJsonEditor } from "./MonacoJsonEditor";
import { workspaceDropdownProps } from "./workspaceDropdownProps";
import { PanelFilename } from "./PanelFilename";
import type { Workspace } from "../hooks/useWorkspace";

export function DatasetPanel({ model }: { model: Workspace["datasetPanel"] }) {
  const { dataset, modified, token, recoverData, editor, json, changeJson, wordWrap, changeWordWrap, chooseDataFile, editorAvailable, saveData, openLabel, commandKey, saveAsShortcut, formatShortcut, saving, formatJson, copyJson, dataInput, reopenDataId, openData } = model;
  function getDataMenuItems() {
    return [
      { key: "open", label: openLabel, extra: editorAvailable ? `${commandKey}O` : undefined, onClick: chooseDataFile },
      { key: "save", label: "Save", extra: editorAvailable ? `${commandKey}S` : undefined, disabled: !editorAvailable || saving, onClick: () => void saveData() },
      { key: "save-as", label: "Save As…", extra: editorAvailable ? saveAsShortcut : undefined, disabled: !editorAvailable || saving, onClick: () => void saveData(true) },
      { type: "divider" as const, key: "file-actions-end" },
      { key: "format", label: "Format", extra: editorAvailable ? formatShortcut : undefined, disabled: !editorAvailable, onClick: () => void formatJson() },
      { key: "wrap", label: "Word wrap", role: "menuitemcheckbox", "aria-checked": wordWrap, disabled: !editorAvailable, onClick: () => editor.current?.toggleWrap() },
      { key: "copy", label: "Copy", disabled: !editorAvailable, onClick: () => void copyJson() }
    ].map(item => "type" in item
      ? item
      : {
        ...item,
        icon: item.key === "wrap" && wordWrap
          ? <CheckOutlined aria-hidden="true" />
          : <span aria-hidden="true" />
      });
  }

  function handleDataFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    const replaceId = reopenDataId.current;
    reopenDataId.current = undefined;
    event.target.value = "";

    if (file) void openData(file, replaceId);
  }

  return (
    <section className="source" aria-label="Input dataset">
      <input
        ref={dataInput}
        type="file"
        hidden
        accept=".json,application/json"
        aria-label="Upload dataset"
        onChange={handleDataFileChange}
      />
      <Flex className="pane-header" align="center" gap="small">
        <Flex align="center" gap="middle" flex="auto" style={{ minWidth: 0 }}>
          <Typography.Text strong style={{ flexShrink: 0 }}>Dataset</Typography.Text>
          <Flex className="context-name" align="center" gap="small" wrap={false}>
            <PanelFilename
              name={dataset?.name}
              source={dataset?.kind === "example"
                ? "Example file"
                : dataset?.kind === "file" || dataset?.kind === "upload"
                  ? "Local file"
                  : undefined}
            />
            {modified && (
              <Badge
                status="default"
                color={token.colorTextSecondary}
                aria-label="Dataset modified"
                style={{ flex: "none" }}
              />
            )}
          </Flex>
        </Flex>
        <Dropdown
          {...workspaceDropdownProps}
          trigger={["click"]}
          menu={{ items: getDataMenuItems() }}
        >
          <Button
            size="small"
            type="text"
            icon={<EllipsisOutlined />}
            aria-label="Dataset menu"
          />
        </Dropdown>
      </Flex>
      {dataset?.restoreError && (
        <Alert
          type="warning"
          title={`${dataset.name}: ${dataset.restoreError}`}
          action={(
            <Button size="small" onClick={() => void recoverData(dataset)}>
              {dataset.needsPermission && dataset.handle?.requestPermission ? "Grant access" : "Reopen…"}
            </Button>
          )}
        />
      )}
      <MonacoJsonEditor
        ref={editor}
        datasetId={dataset?.id}
        value={json}
        onChange={changeJson}
        wordWrap={wordWrap}
        onWordWrapChange={changeWordWrap}
        fileCommands={{
          open: chooseDataFile,
          ...(editorAvailable
            ? { save: () => void saveData(), saveAs: () => void saveData(true) }
            : {})
        }}
      />
    </section>
  );
}
