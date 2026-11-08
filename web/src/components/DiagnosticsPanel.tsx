import { Button, ConfigProvider, Empty, Flex, Tabs, Tooltip, Typography } from "antd";
import { CloseCircleFilled, CloseOutlined, CopyOutlined, DeleteOutlined, ExclamationCircleFilled } from "@ant-design/icons";
import { AnalysisView } from "./AnalysisView";
import { ProblemsLabel } from "./ProblemsLabel";
import type { Workspace } from "../hooks/useWorkspace";

export function DiagnosticsPanel({ model }: { model: Workspace["diagnosticsPanel"] }) {
  const { token, diagnosticsTab, setDiagnosticsTab, changeProblemsOpen, errors, setErrors, copyProblems,
    copyAnalysis, analysis, analysisWireResult } = model;
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

  function getTabItems() {
    return [
      {
        key: "problems",
        label: <ProblemsLabel count={errors.length} />,
        children: (
          <>
            <Flex
              role="toolbar"
              aria-label="Problems actions"
              align="center"
              gap={token.paddingXXS}
              style={{ marginTop: token.paddingXXS, marginBottom: token.marginXXS }}
            >
              <Tooltip title="Copy">
                <Button
                  size="small"
                  type="text"
                  icon={<CopyOutlined />}
                  aria-label="Copy"
                  onClick={() => void copyProblems()}
                />
              </Tooltip>
              <Tooltip title="Clear">
                <Button
                  size="small"
                  type="text"
                  icon={<DeleteOutlined />}
                  aria-label="Clear"
                  onClick={() => setErrors([])}
                />
              </Tooltip>
            </Flex>
            {errors.length
              ? <div className="diagnostic-list">{errors.map(renderProblem)}</div>
              : <Empty image={Empty.PRESENTED_IMAGE_SIMPLE} description="No problems" />}
          </>
        )
      },
      {
        key: "analysis",
        label: "Analysis",
        children: (
          <>
            <Flex
              role="toolbar"
              aria-label="Analysis actions"
              align="center"
              gap={token.paddingXXS}
              style={{ marginTop: token.paddingXXS, marginBottom: token.marginXXS }}
            >
              <Tooltip title="Copy">
                <Button
                  size="small"
                  type="text"
                  icon={<CopyOutlined />}
                  aria-label="Copy"
                  onClick={() => void copyAnalysis()}
                />
              </Tooltip>
            </Flex>
            <AnalysisView analysis={analysis} />
          </>
        )
      }
    ];
  }

  return (
    <ConfigProvider
      componentSize="small"
      theme={{
        token: { fontSize: 12, fontSizeSM: 12, fontSizeLG: 12, fontSizeHeading5: 12 },
        components: {
          Button: { contentFontSizeSM: 12 },
          Tabs: {
            titleFontSizeSM: 12,
            horizontalItemPaddingSM: `${token.paddingXXS}px ${token.paddingSM}px`,
            horizontalItemGutter: token.marginXS,
            horizontalMargin: "0"
          },
          Table: {
            cellFontSizeSM: 12,
            cellPaddingBlockSM: token.paddingXXS,
            cellPaddingInlineSM: token.paddingXS
          },
          Typography: { titleMarginTop: token.marginXS, titleMarginBottom: token.marginXXS }
        }
      }}
    >
      <section className="problems" aria-label="Diagnostics" style={{ fontSize: 12 }}>
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
            <Flex role="toolbar" aria-label="Diagnostics actions" align="center" gap={token.paddingXXS}>
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
          items={getTabItems()}
        />
      </section>
    </ConfigProvider>
  );
}
