import { Space, Typography } from "antd";

export function ShortcutKeys({ value, color }: { value: string; color?: string }) {
  const keys = value.split(/[ +]+/);
  const names: Record<string, string> = { "⇧": "Shift", "⌘": "Command", "⌥": "Option", "⎋": "Escape" };
  return (
    <Space
      size={4}
      wrap={false}
      className="shortcut"
      role="group"
      aria-label={keys.map(key => names[key] ?? key).join(" + ")}
    >
      {keys.map((key, index) => (
        <Typography.Text
          keyboard
          aria-hidden="true"
          key={index}
          style={color ? { color } : undefined}
        >
          {key}
        </Typography.Text>
      ))}
    </Space>
  );
}

