import { ConfigProvider, Descriptions, Empty, Table, theme, Tree, Typography } from "antd";
import type { TreeDataNode } from "antd";
import type { Diagnostic, TemplateAnalysis, TemplateRequirement } from "../types";

function requirementNodes(fields: TemplateRequirement[], parent = ""): TreeDataNode[] {
  return fields.map((field, index) => ({
    key: `${parent}/${index}`,
    title: `${field.name} · ${field.type}`,
    children: field.fields ? requirementNodes(field.fields, `${parent}/${index}`) : undefined
  }));
}

export function AnalysisView({ analysis }: { analysis?: TemplateAnalysis }) {
  const { token } = theme.useToken();
  if (!analysis) return <Empty image={Empty.PRESENTED_IMAGE_SIMPLE} description="No analysis. Analyze a template to inspect its requirements." />;
  return <ConfigProvider componentSize="small" theme={{ algorithm: theme.compactAlgorithm, components: { Typography: { titleMarginTop: token.marginXS, titleMarginBottom: token.marginXXS } } }}>
    <Descriptions size="small" items={[
      { key: "valid", label: "Validity", children: analysis.valid ? "Valid" : "Invalid" },
      { key: "errors", label: "Errors", children: analysis.errors.length },
      { key: "warnings", label: "Warnings", children: analysis.warnings.length }
    ]} />
    <Typography.Title level={5}>Requirements</Typography.Title>
    {analysis.requirements.length ? <Tree key={JSON.stringify(analysis.requirements)} defaultExpandAll selectable={false} treeData={requirementNodes(analysis.requirements)} /> : <Typography.Paragraph>No requirements</Typography.Paragraph>}
    <Typography.Title level={5}>Occurrences</Typography.Title>
    <Table size="small" pagination={false} scroll={{ x: true }} dataSource={analysis.occurrences.map((item, key) => ({ ...item, key }))} columns={[
      { title: "Kind", dataIndex: "kind" }, { title: "Expression", dataIndex: "expression" }, { title: "Path", dataIndex: "path" },
      { title: "Part", dataIndex: "part" }, { title: "Location", dataIndex: "location" }
    ]} />
    {(["errors", "warnings"] as const).map(kind => analysis[kind].length > 0 && <div key={kind}>
      <Typography.Title level={5}>{kind === "errors" ? "Errors" : "Warnings"}</Typography.Title>
      <Table size="small" pagination={false} scroll={{ x: true }} dataSource={analysis[kind].map((item: Diagnostic, key) => ({ ...item, key }))} columns={[
        { title: "Title", dataIndex: "title" }, { title: "Detail", dataIndex: "detail" }, { title: "Code", dataIndex: "code" },
        { title: "Field", dataIndex: "field" }, { title: "Data path", dataIndex: "dataPath" }, { title: "Part", dataIndex: "part" },
        { title: "Location", dataIndex: "location" }, { title: "Formatter", dataIndex: "formatter" }
      ]} />
    </div>)}
  </ConfigProvider>;
}
