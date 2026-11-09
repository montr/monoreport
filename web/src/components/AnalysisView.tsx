import { Empty, Flex, Table, theme, Tree, Typography } from "antd";
import type { TreeDataNode } from "antd";
import type { Diagnostic, TemplateAnalysis, TemplateRequirement } from "../types";

function requirementNodes(fields: TemplateRequirement[], parent = ""): TreeDataNode[] {
  return fields.map((field, index) => ({
    key: `${parent}/${index}`,
    title: (
      <Flex gap={4} align="baseline">
        <Typography.Text>{field.name}</Typography.Text>
        <Typography.Text type="secondary">{field.type}</Typography.Text>
      </Flex>
    ),
    children: field.fields ? requirementNodes(field.fields, `${parent}/${index}`) : undefined
  }));
}

export function AnalysisView({ analysis }: { analysis?: TemplateAnalysis }) {
  const { token } = theme.useToken();
  if (!analysis) return <Empty image={Empty.PRESENTED_IMAGE_SIMPLE} description="No analysis" />;
  return (
    <>
      <Flex
        align="center"
        gap="middle"
        wrap
        style={{ marginTop: token.paddingXS, marginBottom: token.marginSM }}
      >
        <Typography.Text>
          <Typography.Text type="secondary">
            Validity{" "}
          </Typography.Text>
          <Typography.Text
            strong
            style={{ color: analysis.valid ? token.colorSuccess : token.colorError }}
          >
            {analysis.valid ? "Valid" : "Invalid"}
          </Typography.Text>
        </Typography.Text>
        <Typography.Text>
          <Typography.Text type="secondary">Errors{" "}</Typography.Text>
          <Typography.Text strong>{analysis.errors.length}</Typography.Text>
        </Typography.Text>
        <Typography.Text>
          <Typography.Text type="secondary">Warnings{" "}</Typography.Text>
          <Typography.Text strong>{analysis.warnings.length}</Typography.Text>
        </Typography.Text>
      </Flex>

      <Typography.Title level={5}>Requirements</Typography.Title>
      <Typography.Text type="secondary">
        Requirement types describe JSON structure; semantic types are advisory and describe value meaning.
      </Typography.Text>
      {analysis.requirements.length
        ? (
          <Tree
            key={JSON.stringify(analysis.requirements)}
            defaultExpandAll
            selectable={false}
            styles={{ item: { marginBlockEnd: 0 } }}
            treeData={requirementNodes(analysis.requirements)}
          />
        )
        : <Typography.Paragraph>No requirements</Typography.Paragraph>}

      <Typography.Title level={5}>Occurrences</Typography.Title>
      <Table
        size="small"
        tableLayout="fixed"
        pagination={false}
        scroll={{ x: true }}
        dataSource={analysis.occurrences.map((item, key) => ({ ...item, key }))}
        columns={[
          { title: "Kind", dataIndex: "kind", width: 88 },
          { title: "Expression", dataIndex: "expression", width: 240, ellipsis: { showTitle: true } },
          { title: "Path", dataIndex: "path", width: 220, ellipsis: { showTitle: true } },
          { title: "Semantic type", dataIndex: "semanticType", width: 120 },
          { title: "Location", dataIndex: "location" }
        ]}
      />

      {(["errors", "warnings"] as const).map(kind => analysis[kind].length > 0 && (
        <div key={kind}>
          <Typography.Title level={5}>
            {kind === "errors" ? "Errors" : "Warnings"}
          </Typography.Title>
          <Table
            size="small"
            pagination={false}
            scroll={{ x: true }}
            styles={{ body: { cell: { verticalAlign: "top" } } }}
            dataSource={analysis[kind].map((item: Diagnostic, key) => ({ ...item, key }))}
            columns={[
              {
                title: "Title",
                dataIndex: "title",
                render: (title: string, item: Diagnostic) => title === item.code ? null : title
              },
              { title: "Detail", dataIndex: "detail" },
              { title: "Code", dataIndex: "code" },
              { title: "Field", dataIndex: "field" },
              {
                title: "Data path",
                dataIndex: "dataPath",
                onHeaderCell: () => ({ style: { whiteSpace: "nowrap" } })
              },
              { title: "Part", dataIndex: "part" },
              { title: "Location", dataIndex: "location" },
              { title: "Formatter", dataIndex: "formatter" }
            ]}
          />
        </div>
      ))}
    </>
  );
}
