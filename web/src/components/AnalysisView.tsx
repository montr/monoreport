import { ConfigProvider, Empty, Flex, Table, theme, Tree, Typography } from "antd";
import type { TreeDataNode } from "antd";
import type { Diagnostic, TemplateAnalysis, TemplateRequirement } from "../types";

function requirementNodes(fields: TemplateRequirement[], colors: { name: string; type: string }, parent = ""): TreeDataNode[] {
  return fields.map((field, index) => ({
    key: `${parent}/${index}`,
    title: (
      <Flex gap={4} align="baseline">
        <Typography.Text style={{ color: colors.name, fontSize: 12 }}>
          {field.name}
        </Typography.Text>
        <Typography.Text style={{ color: colors.type, fontSize: 12 }}>
          {field.type}
        </Typography.Text>
      </Flex>
    ),
    children: field.fields ? requirementNodes(field.fields, colors, `${parent}/${index}`) : undefined
  }));
}

export function AnalysisView({ analysis }: { analysis?: TemplateAnalysis }) {
  const { token } = theme.useToken();
  const headingStyle = { fontSize: 12, marginTop: token.marginXS, marginBottom: token.marginXXS };
  if (!analysis) return <Empty image={Empty.PRESENTED_IMAGE_SIMPLE} description="No analysis yet." />;
  return (
    <ConfigProvider
      componentSize="small"
      theme={{
        algorithm: theme.compactAlgorithm,
        token: { fontSize: 12, fontSizeSM: 12, fontSizeLG: 12, fontSizeHeading5: 12 },
        components: {
          Table: {
            cellFontSizeSM: 12,
            cellPaddingBlockSM: token.paddingXXS,
            cellPaddingInlineSM: token.paddingXS
          },
          Tree: { titleHeight: token.controlHeightSM - token.paddingXXS },
          Typography: { titleMarginTop: token.marginXS, titleMarginBottom: token.marginXXS }
        }
      }}
    >
      <Flex
        align="center"
        gap="middle"
        wrap
        style={{ marginTop: token.paddingXS, marginBottom: token.marginSM, fontSize: 12 }}
      >
        <Typography.Text style={{ fontSize: 12 }}>
          <Typography.Text type="secondary" style={{ fontSize: 12 }}>
            Validity{" "}
          </Typography.Text>
          <Typography.Text
            strong
            style={{ color: analysis.valid ? token.colorSuccess : token.colorError, fontSize: 12 }}
          >
            {analysis.valid ? "Valid" : "Invalid"}
          </Typography.Text>
        </Typography.Text>
        <Typography.Text style={{ fontSize: 12 }}>
          <Typography.Text type="secondary" style={{ fontSize: 12 }}>Errors{" "}</Typography.Text>
          <Typography.Text strong style={{ fontSize: 12 }}>{analysis.errors.length}</Typography.Text>
        </Typography.Text>
        <Typography.Text style={{ fontSize: 12 }}>
          <Typography.Text type="secondary" style={{ fontSize: 12 }}>Warnings{" "}</Typography.Text>
          <Typography.Text strong style={{ fontSize: 12 }}>{analysis.warnings.length}</Typography.Text>
        </Typography.Text>
      </Flex>

      <Typography.Title level={5} style={headingStyle}>Requirements</Typography.Title>
      {analysis.requirements.length
        ? (
          <Tree
            key={JSON.stringify(analysis.requirements)}
            defaultExpandAll
            selectable={false}
            styles={{ item: { marginBlockEnd: 0 } }}
            treeData={requirementNodes(analysis.requirements, {
              name: token.colorText,
              type: token.colorTextSecondary
            })}
          />
        )
        : <Typography.Paragraph>No requirements</Typography.Paragraph>}

      <Typography.Title level={5} style={headingStyle}>Occurrences</Typography.Title>
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
          { title: "Location", dataIndex: "location" }
        ]}
      />

      {(["errors", "warnings"] as const).map(kind => analysis[kind].length > 0 && (
        <div key={kind}>
          <Typography.Title level={5} style={headingStyle}>
            {kind === "errors" ? "Errors" : "Warnings"}
          </Typography.Title>
          <Table
            size="small"
            pagination={false}
            scroll={{ x: true }}
            styles={{ body: { cell: { verticalAlign: "top" } } }}
            dataSource={analysis[kind].map((item: Diagnostic, key) => ({ ...item, key }))}
            columns={[
              { title: "Title", dataIndex: "title" },
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
    </ConfigProvider>
  );
}
