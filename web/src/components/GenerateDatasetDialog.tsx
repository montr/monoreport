import { Alert, Button, ConfigProvider, Flex, Form, InputNumber, Modal, Radio, Typography } from "antd";
import { CopyOutlined } from "@ant-design/icons";
import { MonacoJsonEditor } from "./MonacoJsonEditor";
import type { Workspace } from "../hooks/useWorkspace";

export function GenerateDatasetDialog({ model }: { model: Workspace["dialogs"]["generateDataset"] }) {
  const { jsonSettingsOpen, pending, submitJsonSettings, invalidateOperation, setJsonSettingsOpen, token,
    jsonSettings, sampleCandidate, sampleFailure, acceptSample, copySampleJson, settingsChanged, captureSampleValidation,
    modified } = model;
  async function generateSample() {
    const validationIsCurrent = captureSampleValidation();
    try {
      const settings = await jsonSettings.validateFields();
      if (!validationIsCurrent()) return;
      await submitJsonSettings(settings);
    } catch {
      // Form displays field validation errors.
    }
  }
  return (
    <Modal
      title="Generate Dataset"
      open={jsonSettingsOpen}
      okText={sampleCandidate ? "Replace Dataset" : "Generate sample"}
      footer={(_, { CancelBtn }) => (
        <Flex vertical gap="small">
          <div style={{ minHeight: token.controlHeightSM }}>
            {sampleCandidate && modified && (
              <Typography.Text type="warning">Replacing the Dataset will discard unsaved changes.</Typography.Text>
            )}
          </div>
          <Flex justify="end" gap="small">
            <CancelBtn />
            {sampleCandidate
              ? <Button style={{ width: 136 }} type="primary" onClick={acceptSample}>Replace Dataset</Button>
              : <Button style={{ width: 136 }} type="primary" loading={pending === "generate-json"} onClick={() => void generateSample()}>Generate sample</Button>}
          </Flex>
        </Flex>
      )}
      width="min(540px, calc(100vw - 32px))"
      onOk={() => sampleCandidate ? acceptSample() : void generateSample()}
      onCancel={() => {
        invalidateOperation();
        setJsonSettingsOpen(false);
      }}
    >
      <ConfigProvider
        componentSize="small"
        theme={{ components: { Form: { itemMarginBottom: token.marginSM } } }}
      >
        <Form
          form={jsonSettings}
          layout="horizontal"
          labelCol={{ span: 8 }}
          wrapperCol={{ span: 16 }}
          labelAlign="start"
          colon={false}
          initialValues={{ itemsPerCollection: 10, values: "filled" }}
          onValuesChange={settingsChanged}
        >
          <Form.Item
            name="itemsPerCollection"
            label="Items per collection"
            extra="Nested collections multiply the total generated elements. The overall limit is 1,000,000."
            rules={[{
              validator: (_, value) => typeof value === "number" && Number.isInteger(value) && value > 0
                ? Promise.resolve()
                : Promise.reject(new Error("Enter a positive integer."))
            }]}
          >
            <InputNumber changeOnBlur={false} />
          </Form.Item>
          <Form.Item name="values" label="Values">
            <Radio.Group
              options={[
                { label: "Filled", value: "filled" },
                { label: "Empty", value: "empty" }
              ]}
            />
          </Form.Item>
        </Form>
      </ConfigProvider>
      {sampleCandidate && (
        <>
          <Flex justify="space-between" align="center" style={{ marginTop: token.marginSM }}>
            <Typography.Text strong>Sample preview</Typography.Text>
            <Button
              size="small"
              type="text"
              icon={<CopyOutlined />}
              aria-label="Copy"
              onClick={() => void copySampleJson()}
            />
          </Flex>
          {sampleCandidate.warnings.length > 0 && (
            <Alert
              type="warning"
              showIcon
              title={`${sampleCandidate.warnings.length} sample warning(s)`}
              description={sampleCandidate.warnings
                .map(item => [item.title, item.detail, item.part, item.location].filter(Boolean).join(": "))
                .join(" · ")}
              style={{ marginTop: token.marginXS }}
            />
          )}
          <div className="sample-preview-editor" style={{ marginTop: token.marginXS }}>
            <MonacoJsonEditor
              value={sampleCandidate.json}
              onChange={() => undefined}
              readOnly
              ariaLabel="Sample JSON preview"
              minimap
              stickyScroll={false}
            />
          </div>
        </>
      )}
      {!sampleCandidate && sampleFailure && sampleFailure.length > 0 && (
        <Alert
          type="error"
          showIcon
          title={`${sampleFailure.length} sample error(s)`}
          description={sampleFailure
            .map(item => [item.title, item.detail, item.part, item.location].filter(Boolean).join(": "))
            .join(" · ")}
          style={{ marginTop: token.marginXS }}
        />
      )}
    </Modal>
  );
}
