import { Button, Dropdown, Flex, Space, Tooltip } from "antd";
import { BulbOutlined, CheckOutlined, DownOutlined, EllipsisOutlined, FileExcelFilled, FileWordFilled } from "@ant-design/icons";
import { workspaceDropdownProps } from "./workspaceDropdownProps";
import { ShortcutKeys } from "./ShortcutKeys";
import type { DocumentFormat } from "../services/documentFormat";
import type { Workspace } from "../hooks/useWorkspace";

function FormatIcon({ format }: { format: DocumentFormat }) {
  const Icon = format === "docx" ? FileWordFilled : FileExcelFilled;
  return (
    <Icon
      className={`format-icon format-icon-${format}`}
      style={{ color: format === "docx" ? "#185abd" : "#217346" }}
      aria-hidden="true"
    />
  );
}

function FormatLabel({ format, title }: { format: DocumentFormat; title: string }) {
  return (
    <span className="format-label">
      <FormatIcon format={format} />
      <span className="format-title">{title}</span>
    </span>
  );
}

export function WorkspaceHeader({ model }: { model: Workspace["header"] }) {
  const { examples, pendingExample, selectedExample, chooseExample, generateShortcut, token, pending, dataset, selectedTemplate, generate, template, analyze, setJsonSettingsOpen, isMac, setShortcutsOpen, saving, clearing, clearLocalData } = model;
  const { liveReloadEnabled, setLiveReloadEnabled } = model;
  function getExampleMenuItems() {
    return (["docx", "xlsx"] as const).map(format => ({
      type: "group" as const,
      key: format,
      label: format.toUpperCase(),
      children: examples.filter(example => example.format === format).map(example => ({
        key: example.id,
        label: <FormatLabel format={format} title={example.title} />
      }))
    })).filter(group => group.children.length > 0);
  }

  function getGenerationMenuItems() {
    return [
      {
        key: "analyze",
        label: "Analyze Template",
        disabled: !template || Boolean(pending),
        onClick: () => void analyze()
      },
      {
        key: "generate-json",
        label: "Generate Dataset...",
        disabled: !template || Boolean(pending),
        onClick: () => setJsonSettingsOpen(true)
      },
      { type: "divider" as const },
      {
        key: "live-reload",
        label: <Tooltip title="Automatically generates from editor changes and reads linked files every second. Background updates may be delayed."><span>Live Reload</span></Tooltip>,
        role: "menuitemcheckbox",
        "aria-checked": liveReloadEnabled,
        icon: liveReloadEnabled ? <CheckOutlined aria-hidden="true" /> : undefined,
        onClick: () => setLiveReloadEnabled(!liveReloadEnabled)
      }
    ];
  }

  function getApplicationMenuItems() {
    return [
      {
        key: "api",
        label: <a href="/docs/" target="_blank" rel="noopener noreferrer">API documentation</a>
      },
      {
        key: "shortcuts",
        label: "Keyboard shortcuts",
        extra: isMac ? "⇧ /" : "Shift+/",
        onClick: () => setShortcutsOpen(true)
      },
      { type: "divider" as const },
      {
        key: "clear-local",
        label: "Clear Local Data…",
        disabled: saving || clearing,
        onClick: () => void clearLocalData()
      }
    ];
  }

  return (
    <Flex
      component="header"
      className="app-header"
      align="center"
      justify="space-between"
      gap="small"
      wrap
    >
      <span className="brand">monoreport</span>
      <Flex
        align="center"
        justify="flex-end"
        gap="small"
        wrap
        flex="auto"
        style={{ maxWidth: "100%" }}
        aria-label="Workspace actions"
      >
        <Tooltip title="Explore examples to see how MonoReport works">
          <Dropdown
            {...workspaceDropdownProps}
            trigger={["click"]}
            menu={{
              items: getExampleMenuItems(),
              selectedKeys: [pendingExample ?? selectedExample ?? ""],
              onClick: ({ key }) => void chooseExample(key)
            }}
          >
            <Button
              aria-label="Examples"
              disabled={examples.length === 0}
              icon={<BulbOutlined aria-hidden="true" />}
            >
              Examples <DownOutlined aria-hidden="true" />
            </Button>
          </Dropdown>
        </Tooltip>

        <Flex align="center" gap="small">
          <Space.Compact>
            <Tooltip
              title={(
                <Space size="small">
                  <span>Generate document</span>
                  <ShortcutKeys value={generateShortcut} color={token.colorTextLightSolid} />
                </Space>
              )}
            >
              <Button
                type="primary"
                aria-label="Generate document"
                loading={pending === "generate"}
                disabled={
                  Boolean(pending) ||
                  Boolean(dataset?.restoreError) ||
                  Boolean(selectedTemplate?.restoreError)
                }
                onClick={() => void generate()}
              >
                Generate
              </Button>
            </Tooltip>
            <Dropdown
              {...workspaceDropdownProps}
              trigger={["click"]}
              menu={{ items: getGenerationMenuItems() }}
            >
              <Button
                type="primary"
                aria-label="Generate actions"
                icon={<DownOutlined />}
              />
            </Dropdown>
          </Space.Compact>

          <Dropdown
            {...workspaceDropdownProps}
            trigger={["click"]}
            menu={{ items: getApplicationMenuItems() }}
          >
            <Button type="text" icon={<EllipsisOutlined />} aria-label="Application menu" />
          </Dropdown>
        </Flex>
      </Flex>
    </Flex>
  );
}
