import { Button, ConfigProvider, Flex, Space, Spin, Typography } from "antd";
import { CheckCircleFilled, CloseCircleFilled, GithubFilled, StopOutlined } from "@ant-design/icons";
import { ProblemsLabel } from "./ProblemsLabel";
import type { Workspace } from "../hooks/useWorkspace";

const repositoryUrl = "https://github.com/montr/monoreport";

export function WorkspaceFooter({ model }: { model: Workspace["footer"] }) {
  const { token, operationStatus, errors, openDiagnostics, problemsOpen, diagnosticsTab } = model;
  const statusLabel = operationStatus.outcome === "running"
    ? operationStatus.action === "Generate" ? "Generating..."
      : operationStatus.action === "Generate Dataset" ? "Generating sample..." : "Analyzing..."
    : operationStatus.outcome === "ready" ? "Ready"
      : operationStatus.outcome === "success" ? "Success"
        : operationStatus.outcome === "error" ? "Error" : "Cancelled";
  const statusIcon = operationStatus.outcome === "success" ? <CheckCircleFilled aria-hidden="true" style={{ color: token.colorSuccess }} />
    : operationStatus.outcome === "error" ? <CloseCircleFilled aria-hidden="true" style={{ color: token.colorError }} />
      : operationStatus.outcome === "cancelled" ? <StopOutlined aria-hidden="true" style={{ color: token.colorWarning }} />
        : undefined;
  return (
    <ConfigProvider
      componentSize="small"
      theme={{
        token: { fontSize: 12, fontSizeSM: 12, fontSizeLG: 12 },
        components: { Button: { contentFontSizeSM: 12 } }
      }}
    >
      <Flex component="footer" align="center" gap="small" wrap style={{ fontSize: 12 }}>
        <Button
          size="small"
          type="text"
          style={{ paddingInlineStart: 0 }}
          onClick={() => openDiagnostics("problems")}
          aria-label={`Workspace status: ${statusLabel}`}
        >
          <Space size="small">
            <span role="status" aria-label={statusLabel}>
              <Flex align="center" gap={token.paddingXXS}>
                <span style={{ display: "inline-flex", alignItems: "center" }}>
                  {operationStatus.outcome === "running" ? <Spin size="small" /> : statusIcon}
                </span>
                <span>{statusLabel}</span>
              </Flex>
            </span>
            {operationStatus.outcome !== "running" && operationStatus.duration !== undefined && (
              <Typography.Text type="secondary">
                {operationStatus.duration.toFixed(0)} ms
              </Typography.Text>
            )}
          </Space>
        </Button>

        {!problemsOpen && (
          <>
            <Button
              size="small"
              type="text"
              aria-expanded={problemsOpen && diagnosticsTab === "problems"}
              onClick={() => openDiagnostics("problems")}
            >
              <ProblemsLabel count={errors.length} />
            </Button>
            <Button
              size="small"
              type="text"
              aria-expanded={problemsOpen && diagnosticsTab === "analysis"}
              onClick={() => openDiagnostics("analysis")}
            >
              Analysis
            </Button>
          </>
        )}

        <span className="project-links">
          <a
            href={repositoryUrl}
            target="_blank"
            rel="noreferrer"
            title="MonoReport on GitHub"
            style={{ color: token.colorTextTertiary }}
          >
            <GithubFilled aria-hidden="true" />MonoReport
          </a>
          {__APP_VERSION__ === "dev"
            ? <span style={{ color: token.colorTextTertiary }}>dev</span>
            : (
              <a
                href={`${repositoryUrl}/releases/tag/v${__APP_VERSION__}`}
                target="_blank"
                rel="noreferrer"
                style={{ color: token.colorTextTertiary }}
              >
                v{__APP_VERSION__}
              </a>
            )}
        </span>
      </Flex>
    </ConfigProvider>
  );
}
