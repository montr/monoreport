import { Button, ConfigProvider, Empty, Flex, Tabs, Tooltip } from "antd";
import { CloseCircleFilled, CloseOutlined, CopyOutlined, DeleteOutlined, ExclamationCircleFilled } from "@ant-design/icons";
import { AnalysisView } from "./AnalysisView";
import { ProblemsLabel } from "./ProblemsLabel";
import type { Workspace } from "../hooks/useWorkspace";

export function DiagnosticsPanel({ model }: { model: Workspace["diagnosticsPanel"] }) {
  const { token, diagnosticsTab, setDiagnosticsTab, changeProblemsOpen, errors, setErrors, copyProblems,
    copyAnalysis, analysis } = model;
  function renderProblem(error: typeof errors[number], index: number) {
    const severity = error.severity ?? "Error";
    const Icon = severity === "Warning" ? ExclamationCircleFilled : CloseCircleFilled;
    const code = error.code ?? error.title;
    const title = error.title === code ? undefined : error.title;

    return (
      <div className="issue" key={`${code}-${index}`}>
        <div className="issue-content">
          <div className="issue-heading">
            <div
              role="img"
              aria-label={severity}
              className="issue-icon"
              style={{ color: severity === "Warning" ? token.colorWarning : token.colorError }}
            >
              <Icon aria-hidden="true" />
            </div>
            <strong>{code}</strong>
            {title && <span>{title}</span>}
          </div>
          <p>
            {[error.field, error.fieldCode, error.dataPath, error.part, error.location,
              error.formatter,
              error.removedCharacterCount === undefined
                ? undefined
                : `removed ${error.removedCharacterCount}`,
              error.traceId === undefined ? undefined : `trace ${error.traceId}`,
              error.limitReached ? "error limit reached" : undefined,
              error.detail]
              .filter(Boolean)
              .join(" · ")}
          </p>
        </div>
      </div>
    );
  }

  return (
    <ConfigProvider
      componentSize="small"
      theme={{
        token: { fontSize: token.fontSizeSM, fontSizeSM: token.fontSizeSM, fontSizeHeading5: token.fontSizeSM },
        components: {
          Tabs: {
            horizontalItemPaddingSM: `${token.paddingXXS}px ${token.paddingSM}px`,
            horizontalItemGutter: token.marginXS,
            horizontalMargin: "0"
          },
          Table: {
            cellPaddingBlockSM: token.paddingXXS,
            cellPaddingInlineSM: token.paddingXS
          },
          Tree: { titleHeight: token.controlHeightSM - token.paddingXXS },
          Typography: { titleMarginTop: token.marginXS, titleMarginBottom: token.marginXXS }
        }
      }}
    >
      <section className="problems" aria-label="Diagnostics">
        <Tabs
          size="small"
          styles={{
            root: { height: "100%", overflow: "hidden" },
            header: { background: token.colorFillQuaternary, paddingBlock: 0 },
            body: { height: "100%" },
            content: { height: "100%", overflow: "auto" }
          }}
          activeKey={diagnosticsTab}
          onChange={key => setDiagnosticsTab(key as Workspace["diagnosticsPanel"]["diagnosticsTab"])}
          tabBarExtraContent={(
            <Flex
              role="toolbar"
              aria-label={diagnosticsTab === "problems" ? "Problems actions" : "Analysis actions"}
              align="center"
              gap={token.paddingXXS}
            >
              <Tooltip title="Copy">
                <Button
                  size="small"
                  type="text"
                  icon={<CopyOutlined />}
                  aria-label="Copy"
                  onClick={() => void (diagnosticsTab === "problems" ? copyProblems() : copyAnalysis())}
                />
              </Tooltip>
              {diagnosticsTab === "problems" && (
                <Tooltip title="Clear">
                  <Button
                    size="small"
                    type="text"
                    icon={<DeleteOutlined />}
                    aria-label="Clear"
                    onClick={() => setErrors([])}
                  />
                </Tooltip>
              )}
              <Tooltip title="Hide diagnostics">
                <Button
                  size="small"
                  type="text"
                  icon={<CloseOutlined />}
                  aria-label="Hide diagnostics"
                  onClick={() => changeProblemsOpen(false)}
                />
              </Tooltip>
            </Flex>
          )}
          items={[
            {
              key: "problems",
              label: <ProblemsLabel count={errors.length} />,
              children: errors.length
                ? <div className="diagnostic-list">{errors.map(renderProblem)}</div>
                : <Empty image={Empty.PRESENTED_IMAGE_SIMPLE} description="No problems" />
            },
            { key: "analysis", label: "Analysis", children: <AnalysisView analysis={analysis} /> }
          ]}
        />
      </section>
    </ConfigProvider>
  );
}
