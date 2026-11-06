import { theme, Tooltip, Typography } from "antd";

export function PanelFilename({ name, source }: { name?: string; source?: "Example file" | "Local file" }) {
  const { token } = theme.useToken();
  return (
    <Tooltip title={source}>
      <Typography.Text
        ellipsis
        title={name}
        style={{
          minWidth: 0,
          flex: "0 1 auto",
          color: source === "Example file" ? token.colorTextTertiary : token.colorText
        }}
      >
        {name}
      </Typography.Text>
    </Tooltip>
  );
}

