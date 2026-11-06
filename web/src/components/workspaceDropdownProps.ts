import type { DropdownProps } from "antd";

// Native alignment and shifting keep right-edge triggers and long menus inside the viewport.
export const workspaceDropdownProps = {
  placement: "bottomRight",
  autoAdjustOverflow: true,
  align: { overflow: { adjustX: true, adjustY: true, shiftX: true, shiftY: true } },
  styles: { root: { maxWidth: "calc(100vw - 16px)" } }
} satisfies Partial<DropdownProps>;
