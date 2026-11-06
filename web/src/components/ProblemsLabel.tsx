import { Badge, Space } from "antd";

export function ProblemsLabel({ count }: { count: number }) {
  return (
    <Space size="small">
      Problems{count > 0 && <Badge count={count} size="small" />}
    </Space>
  );
}
