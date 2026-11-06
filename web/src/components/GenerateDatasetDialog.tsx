import { ConfigProvider, Form, InputNumber, Modal, Radio } from "antd";
import type { Workspace } from "../hooks/useWorkspace";

export function GenerateDatasetDialog({ model }: { model: Workspace["dialogs"]["generateDataset"] }) {
  const { jsonSettingsOpen, pending, submitJsonSettings, invalidateOperation, setJsonSettingsOpen, token, jsonSettings } = model;
  return (
    <Modal
      title="Generate Dataset"
      open={jsonSettingsOpen}
      okText="OK"
      footer={(_, { OkBtn }) => <OkBtn />}
      width={400}
      okButtonProps={{ loading: pending === "generate-json" }}
      onOk={() => void submitJsonSettings()}
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
          labelCol={{ flex: "150px" }}
          wrapperCol={{ flex: "auto" }}
          labelAlign="start"
          colon={false}
          initialValues={{ itemsPerCollection: 10, values: "empty" }}
        >
          <Form.Item
            name="itemsPerCollection"
            label="Items per collection"
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
                { label: "Empty", value: "empty" },
                { label: "Random", value: "random" }
              ]}
            />
          </Form.Item>
        </Form>
      </ConfigProvider>
    </Modal>
  );
}
