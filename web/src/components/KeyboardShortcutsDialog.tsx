import { Flex, Modal, Typography } from "antd";
import { ShortcutKeys } from "./ShortcutKeys";
import type { Workspace } from "../hooks/useWorkspace";

export function KeyboardShortcutsDialog({ model }: { model: Workspace["dialogs"]["shortcuts"] }) {
  const { shortcutsOpen, setShortcutsOpen, shortcutGroups } = model;
  return (
    <Modal
      title="Keyboard shortcuts"
      open={shortcutsOpen}
      onCancel={() => setShortcutsOpen(false)}
      footer={null}
      width={440}
      style={{ maxWidth: "calc(100vw - 32px)" }}
    >
      <Flex vertical gap="small">
        {shortcutGroups.map(group => (
          <section key={group.id} aria-labelledby={`${group.id}-shortcuts-title`}>
            <Typography.Title
              level={5}
              id={`${group.id}-shortcuts-title`}
              style={{ marginTop: 0 }}
            >
              {group.label}
            </Typography.Title>
            <Flex vertical gap="small">
              {group.items.map(item => (
                <Flex
                  key={item.action}
                  justify="space-between"
                  align="center"
                  gap="middle"
                  role="group"
                  aria-label={item.action}
                >
                  <Typography.Text>{item.action}</Typography.Text>
                  <ShortcutKeys value={item.shortcut} />
                </Flex>
              ))}
            </Flex>
          </section>
        ))}
      </Flex>
    </Modal>
  );
}
